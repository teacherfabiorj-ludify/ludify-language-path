// ============================================================================
// RPL — CORE RULEBOOK — single-manuscript build script
//
// This file is the ONE source of truth for the whole Core Rulebook.
// Every chapter lives here, as data + layout, in the order it appears in
// the book (CH1, CH2, ... Appendix A).
//
// To add or revise a chapter: edit this file, then run `node build.js`.
// It always regenerates the ENTIRE book from scratch into one docx —
// there is never a separate file per chapter, and never a "v2"/"final"
// filename. Version history lives in GitHub commits on this one file.
//
// To rename the system once the final name is locked in: change GAME_NAME
// below. Every mention in every chapter updates automatically.
// ============================================================================

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, PageBreak,
  ImageRun, Footer, PageNumber,
} = require("docx");
const { readFileSync } = require("fs");
const sizeOf = (() => {
  // Minimal PNG dimension reader (avoids adding a new npm dependency)
  return (path) => {
    const buf = readFileSync(path);
    const width = buf.readUInt32BE(16);
    const height = buf.readUInt32BE(20);
    return { width, height };
  };
})();

const GAME_NAME = "Ludify RPL"; // locked 12/08/2026 — Roleplaying Language

// ---- Palette (shared across Core Rulebook AND Master's Guide for visual consistency) ----
const ACCENT = "2A78D6";
const GOOD = "0CA30C";
const WARN = "FAB219";
const CRIT = "D03B3B";
const INK = "0B0B0B";
const INK_SECONDARY = "52514E";
const MUTED = "898781";
const BRAND = "D2691E"; // laranja da marca — só para "Ludify RPL"
const BOX_BG = "F2F2F0";
const ZEBRA = "F7F7F5";
const WHITE = "FFFFFF";

const PAGE_W = 12240; // US Letter
const PAGE_H = 15840;
const MARGIN = 1080; // 0.75in

// ---------------------------------------------------------------------------
// Shared layout helpers
// ---------------------------------------------------------------------------

function noBorder() {
  return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
}

function spacer(h = 120) {
  return new Paragraph({ spacing: { after: h }, children: [] });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function eyebrow(text) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({ text: text.toUpperCase(), bold: true, color: ACCENT, size: 20, characterSpacing: 20 }),
    ],
  });
}

function chapterTitle(text) {
  return new Paragraph({
    spacing: { after: 200 },
    children: [ new TextRun({ text, bold: true, color: INK, size: 64 }) ],
  });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 280, after: 100 },
    keepNext: true,
    children: [ new TextRun({ text, bold: true, color: INK, size: 26 }) ],
  });
}


// ---------------------------------------------------------------------------
// GAME TERMS — every named Focus, Move, Signature Move, Boon and resource gets
// picked out of running prose in accent bold, so a reader can tell "you Parley"
// (a Move) from "you parley" (a verb) without having to guess from context.
// Matching is case-sensitive and word-bounded; longest terms first so that
// "Read the Room" wins over "Read the Scene" and neither is cut in half.
// ---------------------------------------------------------------------------
const GAME_TERMS = [
  // Signature Moves
  "Nothing Gets Past Me", "Everyone Has a Price", "Speak for the Table",
  "Angles and Openings", "Already Knew That", "I Planned for This",
  "Nothing Slips By", "One Step Ahead", "Shield the Line", "Hold the Door",
  "Read the Room", "Follow Me",
  // The six core Moves
  "Act Under Pressure", "Persuade or Manipulate", "Help or Interfere",
  "Read the Scene", "Face Danger", "Parley",
  // System nouns
  "Growth Moment", "Growth Level", "Growth Ledger", "Language Focus",
  "Language Point", "Spotlight Token", "Signature Move", "Homework Bonus",
  "Session Zero", "Cross-Training", "Focus Shift", "Legacy Boon",
  // The four Archetypes
  "Vanguard", "Diplomat", "Strategist", "Scout",
  // The four Focuses
  "Courage", "Empathy", "Instinct", "Wit",
  // What you carry, and how far away it is
  "Within reach", "Out of sight", "Far away", "Nearby", "Boon", "Kit", "Pack",
];

const TERM_RE = new RegExp(
  "\\b(" + GAME_TERMS
    .slice()
    .sort((a, b) => b.length - a.length)
    .map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|") + ")(s\\b|\\b)",
  "g"
);

// One pass finds both the brand name and every game term, so a term can never
// be swallowed by the brand match or vice versa.
const MARK_RE = new RegExp(
  "(Ludify RPL)|\\b(" + GAME_TERMS
    .slice()
    .sort((a, b) => b.length - a.length)
    .map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|") + ")(s\\b|\\b)",
  "g"
);

// Split a string into TextRuns: brand name in brand orange, game terms in
// accent blue, everything else untouched.
function markTerms(text, base = {}) {
  if (!text) return [new TextRun({ text: text || "", ...base })];
  const runs = [];
  let last = 0;
  let m;
  MARK_RE.lastIndex = 0;
  while ((m = MARK_RE.exec(text)) !== null) {
    if (m.index > last) {
      runs.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    }
    const isBrand = m[1] !== undefined;
    runs.push(new TextRun({
      text: m[0], ...base, bold: true, color: isBrand ? BRAND : ACCENT,
    }));
    last = m.index + m[0].length;
  }
  if (last < text.length) {
    runs.push(new TextRun({ text: text.slice(last), ...base }));
  }
  return runs.length ? runs : [new TextRun({ text, ...base })];
}

function bodyPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 160 },
    children: markTerms(text, { color: INK, size: 22, italics: opts.italics || false }),
  });
}

function flavorQuote(lines) {
  // Accepts a single string, or an array of strings rendered as separate lines.
  const arr = Array.isArray(lines) ? lines : [lines];
  return new Table({
    width: { size: 10080, type: WidthType.DXA },
    columnWidths: [10080],
    borders: {
      top: noBorder(), bottom: noBorder(), right: noBorder(),
      left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT },
      insideHorizontal: noBorder(), insideVertical: noBorder(),
    },
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 10080, type: WidthType.DXA },
            margins: { top: 40, bottom: 40, left: 240, right: 120 },
            children: arr.map((text, i) => new Paragraph({
              spacing: { after: i === arr.length - 1 ? 0 : 100 },
              children: markTerms(text, { italics: true, color: INK_SECONDARY, size: 22 }),
            })),
          }),
        ],
      }),
    ],
  });
}

// Callout colour is decided by FUNCTION, never by taste:
//   example  (green)  — a concrete case showing the rule in motion
//   clarify  (blue)   — why a rule works this way, or a finer point
//   warn     (amber)  — a limit, a trap, or something easy to get wrong
const BOX_KINDS = { example: GOOD, clarify: ACCENT, warn: WARN };

function calloutBox(label, text, kind = "clarify") {
  const color = BOX_KINDS[kind] || kind;
  return new Table({
    width: { size: 10080, type: WidthType.DXA },
    columnWidths: [10080],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color },
      bottom: { style: BorderStyle.SINGLE, size: 6, color },
      left: { style: BorderStyle.SINGLE, size: 6, color },
      right: { style: BorderStyle.SINGLE, size: 6, color },
      insideHorizontal: noBorder(), insideVertical: noBorder(),
    },
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 10080, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX_BG },
            margins: { top: 140, bottom: 140, left: 200, right: 200 },
            children: [
              new Paragraph({
                spacing: { after: 60 },
                children: [ new TextRun({ text: label.toUpperCase(), bold: true, color, size: 18, characterSpacing: 15 }) ],
              }),
              new Paragraph({ children: markTerms(text, { size: 20, color: INK }) }),
            ],
          }),
        ],
      }),
    ],
  });
}

// Embed a PNG at a given display width (DXA), auto-scaling height to match its
// real aspect ratio, with an optional small italic caption underneath.
function figure(path, opts = {}) {
  const dxaWidth = opts.width ?? 10080; // full content width by default
  const pxWidth = Math.round(dxaWidth / 15); // 1px ≈ 15 DXA at 96dpi/1440-per-inch
  const { width: nativeW, height: nativeH } = sizeOf(path);
  const pxHeight = Math.round(pxWidth * (nativeH / nativeW));

  const parts = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: opts.caption ? 40 : 160 },
      children: [
        new ImageRun({
          type: "png",
          data: readFileSync(path),
          transformation: { width: pxWidth, height: pxHeight },
        }),
      ],
    })
  ];
  if (opts.caption) {
    parts.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [ new TextRun({ text: opts.caption, italics: true, size: 18, color: MUTED }) ],
    }));
  }
  return parts;
}

// Generic 3-column reference table (used by the book roadmap and the Ch.4 quick reference)
function threeColTable(headers, rows, widths) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: ACCENT },
        margins: { top: 90, bottom: 90, left: 120, right: 120 },
        children: [ new Paragraph({ children: [ new TextRun({ text: h, bold: true, color: WHITE, size: 18 }) ] }) ],
      })
    ),
  });
  const bodyRows = rows.map((r, idx) => new TableRow({
    cantSplit: true,
    children: r.map((cell, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: idx % 2 ? ZEBRA : WHITE },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [ new Paragraph({ children: i === 0
        ? [ new TextRun({ text: cell, size: 19, color: INK, bold: true }) ]
        : markTerms(cell, { size: 19, color: INK_SECONDARY }) }) ],
    })),
  }));
  return new Table({
    width: { size: 10080, type: WidthType.DXA },
    columnWidths: widths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: MUTED },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: MUTED },
      left: { style: BorderStyle.SINGLE, size: 4, color: MUTED },
      right: { style: BorderStyle.SINGLE, size: 4, color: MUTED },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "E1E0D9" },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "E1E0D9" },
    },
    rows: [headerRow, ...bodyRows],
  });
}

// A thin accent rule across the foot of every page, with the page number in
// accent bold sitting just under its right-hand end.
function pageFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 120 },
        border: {
          top: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 6 },
        },
        children: [
          new TextRun({ children: [PageNumber.CURRENT], bold: true, color: ACCENT, size: 20 }),
        ],
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// TITLE PAGE
// ---------------------------------------------------------------------------

function titlePage() {
  const children = [];
  children.push(spacer(1400));
  children.push(new Paragraph({
    spacing: { after: 80 },
    children: [ new TextRun({ text: "Ludify RPL", bold: true, color: BRAND, size: 88 }) ],
  }));
  children.push(new Paragraph({
    spacing: { after: 220 },
    children: [ new TextRun({ text: "Roleplaying Language", color: INK_SECONDARY, size: 30 }) ],
  }));
  children.push(new Paragraph({
    spacing: { after: 40 },
    children: [ new TextRun({ text: "CORE RULEBOOK", bold: true, color: ACCENT, size: 26, characterSpacing: 60 }) ],
  }));
  children.push(spacer(1800));
  children.push(new Paragraph({
    children: [ new TextRun({ text: "Ludify — Idiomas com diversão e propósito", color: MUTED, size: 20 }) ],
  }));
  children.push(pageBreak());
  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 1 — Welcome to [Game Name]
// ---------------------------------------------------------------------------

// Contents. The page column is filled by a two-pass build: the book is
// generated once, the real first page of every chapter is read back out of the
// PDF, and those numbers are written in here. Change a chapter and the numbers
// have to be re-read — see PAGES below.
const PAGES = {
  "Ch. 1": "2",
  "Ch. 2": "4",
  "Ch. 3": "6",
  "Ch. 4": "9",
  "Ch. 5": "12",
  "Ch. 6": "21",
  "Ch. 7": "25",
  "Ch. 8": "28",
  "Ch. 9": "30",
  "Ch. 10": "32",
  "Ch. 11": "34",
  "Ch. 12": "35",
  "Appendix A": "36",
};

const contentsRows = [
  ["Ch. 1 — Welcome", "What this game is, and how a table with mixed English levels plays together."],
  ["Ch. 2 — How a Turn Works", "The 2d6 engine, the three outcome bands, one worked example."],
  ["Ch. 3 — Your Character", "Focuses, your Language Focus, and how to read your own sheet."],
  ["Ch. 4 — Moves", "The six core Moves everyone shares, no matter the world."],
  ["Ch. 5 — Archetypes & Growth", "Four Archetypes, and the twelve-level track your character climbs as you do."],
  ["Ch. 6 — What You Carry", "Your Kit, six Pack slots, Boons, money and distance — all measured in words."],
  ["Ch. 7 — Making a Character", "Ana builds hers from scratch, one decision at a time."],
  ["Ch. 8 — Language Points & Spotlight Tokens", "How you earn them, how you spend them."],
  ["Ch. 9 — The Worlds You Can Play In", "Six settings this same engine runs, at a glance."],
  ["Ch. 10 — Table Etiquette", "The four house rules that keep this a safe place to make mistakes."],
  ["Ch. 11 — Your Responsibilities", "Homework, punctuality, showing up ready."],
  ["Ch. 12 — Session Zero Checklist", "What to agree on before your very first adventure."],
  ["Appendix A — As Regras em Português", "The core mechanics, summarised in Portuguese, for players still starting out."],
];

const bookContents = contentsRows.map(([title, blurb]) => {
  const key = title.split(" — ")[0];
  return [title, blurb, String(PAGES[key] ?? "—")];
});

function chapter1() {
  const children = [];
  children.push(eyebrow("Chapter 1"));
  children.push(chapterTitle(`Welcome to ${GAME_NAME}`));

  children.push(flavorQuote(
    `The lantern gutters. Oren's hand hasn't left the pommel of his sword since you mentioned the merchant's name. Whatever you say next actually matters — and it has to be your own words, right now, in English.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `${GAME_NAME} is a tabletop roleplaying game built for one purpose: getting genuinely better at English by using it for something that matters in the moment — not by studying it. You'll sit down with your teacher (who runs the game as the Game Master, or GM) and a small group of classmates, and together you'll play through a story that unfolds entirely through conversation, choices, and a couple of dice.`
  ));
  children.push(bodyPara(
    `You're not going to fill out a worksheet in the middle of a scene. You're going to talk your way past a suspicious guard, convince a skeptical ally, or figure out what's really going on in a room full of secrets — and the English you need for that is exactly the English you're already working on this week.`
  ));

  children.push(sectionHeading("Why Narrative Comes First"));
  children.push(bodyPara(
    `Every rule in this book exists to serve the story, not the other way around. That's not a stylistic choice — it's the actual engine of how this game teaches. You produce more language, and better language, when you're communicating because a situation demands it, not because you were told to practice. Most of the time, you won't feel like you're “doing grammar” — even in the moments when that's exactly what's happening.`
  ));
  children.push(bodyPara(
    `This also means ${GAME_NAME} stays light on the tactical side by design. There's no miniature-and-grid combat to calculate, no long list of numbers to optimize. Conflict resolves through description and a roll of two six-sided dice, so the table's energy stays where it belongs: on talking to each other.`
  ));

  children.push(sectionHeading("One Table, Every Level"));
  children.push(bodyPara(
    `Tables in ${GAME_NAME} mix students at different levels — someone just starting out might be playing right next to someone who's nearly fluent. That's by design, not a compromise. The game never grades you against the other players at your table. Instead, each of you carries a personal language target into every session — your Language Focus, covered fully in Chapter 3 — and your GM builds moments in the spotlight around it. The story is shared. The language work is yours.`
  ));

  children.push(sectionHeading("Contents"));
  children.push(threeColTable(
    ["CHAPTER", "WHAT'S INSIDE", "PAGE"],
    bookContents,
    [3600, 5280, 1200]
  ));

  children.push(spacer(180));
  children.push(sectionHeading("Before You Play"));
  children.push(bodyPara(
    `This game runs on trying, not on being right — more on exactly what that means in Chapter 10. Your actual starting point isn't Chapter 2, though: it's the Session Zero Checklist at the very end of this book (Chapter 12), where your table aligns on tone, expectations, and how everyone wants to play before the first adventure begins.`
  ));
  children.push(bodyPara(
    `And if the English in these chapters is still ahead of you, turn to Appendix A at the back. Two pages, in Portuguese, covering every rule you need to play your first session. Use it for as long as you need it — and know that the day you stop needing it is one of the things this course is for.`
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 2 — How a Turn Works
// ---------------------------------------------------------------------------

const ASSETS = `${__dirname}/assets`;

function chapter2() {
  const children = [];
  children.push(eyebrow("Chapter 2"));
  children.push(chapterTitle("How a Turn Works"));

  children.push(flavorQuote(
    `Two dice hit the table. For a second, nobody knows what happens next — not even the GM.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Whenever you attempt a Move (Chapter 4), the same three steps happen, every time, no matter which Move it is or which world you're playing in. Learn this once and you already know how to play.`
  ));
  children.push(...figure(`${ASSETS}/flow_diagram.png`, {
    caption: "The three steps of a turn — the same sequence, every single time."
  }));

  children.push(sectionHeading("The Engine: 2d6 + Focus"));
  children.push(bodyPara(
    `When a Move calls for a roll, you roll two six-sided dice (2d6), add the two numbers together, then add your rating in whichever Focus that Move uses — Courage, Empathy, Wit, or Instinct. That total is what you compare against the three outcome bands below.`
  ));
  children.push(bodyPara(
    `Why two dice instead of one? Because two dice don't spread results evenly — some totals are far more likely than others. Roll a single d12 and every number from 1 to 12 has exactly the same one-in-twelve chance. Roll two d6 and add them, and the middle totals (7, 8, 9) come up far more often than the extremes (2 or 12), simply because there are more ways to make them. Here's what that actually looks like:`,
    { after: 100 }
  ));
  children.push(...figure(`${ASSETS}/bands_chart.png`, {
    caption: "All 36 possible 2d6 combinations, sorted by total and colored by outcome band."
  }));
  children.push(bodyPara(
    `That's exactly the shape a good scene needs. Clean, total failure should be rare. A perfect, no-cost win should also be rare. Most of the time, life hands you something in between — you get what you wanted, but it costs you something. That middle band isn't a design compromise; it's the most common outcome on purpose.`
  ));
  children.push(calloutBox(
    "In Practice",
    `Same dice, same math, for every player at the table — the beginner and the near-fluent student roll identically. Your English level was never meant to be “character power”; only your Focus rating changes what number you add.`,
    "clarify"
  ));

  children.push(sectionHeading("The Three Outcome Bands"));
  children.push(bodyPara(
    `Once you have your total, read it against these three bands. Every Move in Chapter 4 is written using exactly this structure — once you know what each band means in general, you can read any Move on sight.`,
    { after: 100 }
  ));
  children.push(
    new Table({
      width: { size: 10080, type: WidthType.DXA },
      columnWidths: [1600, 8480],
      borders: {
        top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder(),
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "FFFFFF" },
        insideVertical: noBorder(),
      },
      rows: [
        new TableRow({ children: [
          new TableCell({ width:{size:1600,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,color:"auto",fill:GOOD}, verticalAlign:"center", margins:{top:100,bottom:100,left:100,right:100}, children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"10+",bold:true,color:WHITE,size:22})]})] }),
          new TableCell({ width:{size:8480,type:WidthType.DXA}, margins:{top:100,bottom:100,left:180,right:180}, children:[
            new Paragraph({ children:[ new TextRun({text:"Strong Hit — ", bold:true, size:21, color:INK}), new TextRun({text:"you get exactly what you were going for, cleanly, with no complication attached.", size:21, color:INK_SECONDARY}) ]})
          ] }),
        ]}),
        new TableRow({ children: [
          new TableCell({ width:{size:1600,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,color:"auto",fill:WARN}, verticalAlign:"center", margins:{top:100,bottom:100,left:100,right:100}, children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"7–9",bold:true,color:INK,size:22})]})] }),
          new TableCell({ width:{size:8480,type:WidthType.DXA}, margins:{top:100,bottom:100,left:180,right:180}, children:[
            new Paragraph({ children:[ new TextRun({text:"Mixed Result — ", bold:true, size:21, color:INK}), new TextRun({text:"you get it, but it costs you something — the exact cost is spelled out by the Move itself.", size:21, color:INK_SECONDARY}) ]})
          ] }),
        ]}),
        new TableRow({ children: [
          new TableCell({ width:{size:1600,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,color:"auto",fill:CRIT}, verticalAlign:"center", margins:{top:100,bottom:100,left:100,right:100}, children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"6−",bold:true,color:WHITE,size:22})]})] }),
          new TableCell({ width:{size:8480,type:WidthType.DXA}, margins:{top:100,bottom:100,left:180,right:180}, children:[
            new Paragraph({ children:[ new TextRun({text:"Miss — ", bold:true, size:21, color:INK}), new TextRun({text:"it doesn't work out. The GM makes a move of their own: a complication, a new danger, or a turn you didn't expect.", size:21, color:INK_SECONDARY}) ]})
          ] }),
        ]}),
      ],
    })
  );
  children.push(spacer(160));

  children.push(sectionHeading("A Turn, Start to Finish"));
  children.push(bodyPara(
    `Here's what all three steps look like at a real table. Ana, playing at A2 level, is at the gates of a walled town with a suspicious gatekeeper blocking the way.`,
    { after: 100 }
  ));
  children.push(flavorQuote([
    `GM: "You arrive at the gates. A tired gatekeeper named Oren blocks your cart. 'Merchants, are you? Papers.' He eyes your bags with suspicion." "Ana, what do you tell him?"`,
    `Ana: "If you let us in, we can show you the trade permit inside."`,
  ]));
  children.push(spacer(100));
  children.push(bodyPara(
    `Step 1 — Trigger: Ana is making a direct request to an NPC, backed by something he wants (proof of the permit). That's Parley, using Empathy.`,
    { after: 80 }
  ));
  children.push(bodyPara(
    `Step 2 — Roll: Ana rolls 2d6, gets a 7, and adds her Empathy of +1. Total: 8.`,
    { after: 80 }
  ));
  children.push(bodyPara(
    `Step 3 — Outcome: 8 falls in the 7–9 band — a Mixed Result. The GM narrates the cost: “Oren squints. 'Fine — but someone needs to vouch for you at the guild first. Who's coming with me?'” Ana got what she wanted, but the scene isn't over — there's a string attached, exactly as the band promised.`,
    { after: 80 }
  ));
  children.push(bodyPara(
    `That's the whole engine. Everything else in this book — the six Moves, your Focuses, your Language Focus — just decides which situations call for this sequence, and which number you add when they do.`,
    { italics: true }
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 3 — Your Character
// ---------------------------------------------------------------------------

function chapter3() {
  const children = [];
  children.push(eyebrow("Chapter 3"));
  children.push(chapterTitle("Your Character"));

  children.push(flavorQuote(
    `A blank sheet isn't a test. It's four numbers, a name, and a world waiting for you to walk into it.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Every character in ${GAME_NAME}, no matter the setting, is built from the same small set of parts: four Focuses, a handful of Moves, one Language Focus that's yours alone, and a couple of resources that refill every session. This chapter walks through what each part means and how to read it on your own sheet.`
  ));

  children.push(sectionHeading("The Four Focuses"));
  children.push(bodyPara(
    `Your Focuses are the four numbers you add when you roll (Chapter 2). Every Move in the game is tied to exactly one of them.`,
    { after: 100 }
  ));
  children.push(...figure(`${ASSETS}/focus_cards.png`, { width: 10080 }));

  children.push(sectionHeading("Building Your Focus Array"));
  children.push(bodyPara(
    `At character creation, you take four numbers — +2, +1, +0, and −1 — and assign one to each Focus, in whatever order you want. A Courage-forward character might take Courage +2, Instinct +1, Wit +0, Empathy −1; a smooth talker might put the +2 in Empathy instead. Every character, in every setting, starts from this exact same array. Nobody is mathematically stronger than anybody else — you're only choosing where you're good, not how good you get to be overall.`
  ));
  children.push(calloutBox(
    "In Practice",
    `This is also what keeps every Archetype fair against every other, no matter which world you're playing in — a Warrior in the Fantasy setting and a Netrunner in the Cyberpunk setting both build from the same four numbers. Chapter 5 suggests a starting array for each Archetype, but the choice of how you split your own numbers is always yours.`,
    "clarify"
  ));

  children.push(sectionHeading("Your Language Focus"));
  children.push(bodyPara(
    `This is the one field on your sheet that's yours and nobody else's. Your Language Focus is the grammar structure or vocabulary set you're actively working on this week — assigned by your GM, pulled from whatever you're covering in your regular classes, and updated regularly as you move forward. When a Move you're making happens to line up with it, your GM may ask you specifically to narrate using that structure. Nail it, and that's a Language Point (Chapter 8).`
  ));
  children.push(bodyPara(
    `If you are still in your first level of the course, your Language Focus card carries one line of support in your own language, so you always know exactly what you are aiming for. That single line is the only Portuguese anywhere in this game. Everything else — every rule, every Move, every word the GM or an NPC says — is in English from the first session.`
  ));
  children.push(calloutBox(
    "Example",
    `Language Focus: real conditionals, present/future. Card note: “if / when clauses — se você fizer algo, algo vai acontecer (condição).” Everything else about the scene — what the GM says, what your character says back — stays in English.`,
    "example"
  ));

  children.push(spacer(200));
  children.push(sectionHeading("Reading Your Sheet"));
  children.push(bodyPara(
    `Everything this chapter has described lives in one place. Here is a real sheet — Ana's, at the moment she finished making it — with every block numbered. The table underneath says what goes in each one, who fills it in, and how often it changes.`,
    { after: 60 }
  ));
  children.push(...figure("assets/sheet_annotated.png", {
    caption: "Ana's sheet at Session Zero. Grey fields are filled in by your GM; yellow fields are yours.",
  }));
  children.push(threeColTable(
    ["ON THE SHEET", "WHAT GOES THERE", "HOW OFTEN IT CHANGES"],
    [
      ["1 — Who you are", "Your name, the Setting your table is playing in, and your Archetype (Chapter 5). Your character's name is the only part you choose here.", "Once, when you make the character."],
      ["2 — Growth Level", "How far along the twelve-level track you are (Chapter 5). It is copied from your Growth Ledger, which your GM keeps — never the other way round.", "Once every six units of your course."],
      ["3 — Focuses", "Courage, Empathy, Wit and Instinct, holding your +2, +1, +0 and −1 in the order you chose.", "Almost never. Only a Focus Shift moves them, and nothing ever raises them."],
      ["4 — Language Focus", "The grammar point or vocabulary set you are working on right now, and the Evolve unit it came from.", "Every time you finish a unit. Your GM writes it."],
      ["5 — Signature Move", "The Move that comes with your Archetype, at whatever tier your Growth Level has reached.", "At Levels 3, 7 and 12."],
      ["6 — Your Moves", "The same six Moves every player has, and which Focus each one uses.", "Never. These six are the whole game."],
      ["7 — This session", "Spotlight Tokens and Language Points (Chapter 8).", "Every session. They reset, and they never bank forward."],
      ["8 — Kit and Pack", "Your Kit is the four or five things your Archetype always carries. Your Pack is six slots for everything you pick up (Chapter 6).", "Kit almost never. Pack whenever you take or drop something."],
      ["9 — Boons", "What you earned at each Growth Moment. Your GM writes these; you never add one yourself.", "Once every six units, and only sometimes."],
      ["10 — Money and distance", "Your coins, handfuls and bags, and the four distances this game uses instead of metres (Chapter 6).", "Money, whenever you spend or earn. The four distances, never."],
    ],
    [2400, 5080, 2600]
  ));

  children.push(spacer(180));
  children.push(sectionHeading("One Sheet, In Practice"));
  children.push(bodyPara(
    `Diego plays a Scholar-Mage in the Fantasy setting, at a C1 level. Here's the top of his sheet:`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["FOCUS", "VALUE", "SIGNATURE MOVE"],
    [
      ["Wit", "+2", "Arcane Insight — reflavors Read the Scene to use Wit instead of Instinct."],
      ["Instinct", "+1", "—"],
      ["Courage", "+0", "—"],
      ["Empathy", "−1", "—"],
    ],
    [2200, 1600, 6280]
  ));
  children.push(spacer(120));
  children.push(bodyPara(
    `Diego put his +2 in Wit on purpose — it's both his strongest Focus and the one his Signature Move runs on, so it comes up often. His Language Focus this week is commenting adverbs and the future perfect, straight from his C1 coursework. None of that changes how the dice work for him; it just tells his GM exactly where to aim the spotlight.`,
    { italics: true }
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 4 — Moves
// ---------------------------------------------------------------------------

function howToReadBox() {
  return new Table({
    width: { size: 10080, type: WidthType.DXA },
    columnWidths: [10080],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: ACCENT },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT },
      left: { style: BorderStyle.SINGLE, size: 6, color: ACCENT },
      right: { style: BorderStyle.SINGLE, size: 6, color: ACCENT },
      insideHorizontal: noBorder(),
      insideVertical: noBorder(),
    },
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 10080, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX_BG },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            children: [
              new Paragraph({
                spacing: { after: 100 },
                children: [ new TextRun({ text: "HOW TO READ A MOVE", bold: true, color: ACCENT, size: 20, characterSpacing: 15 }) ],
              }),
              new Paragraph({
                spacing: { after: 80 },
                children: [
                  new TextRun({ text: "Trigger — ", bold: true, size: 20, color: INK }),
                  new TextRun({ text: "the situation that calls for this Move.", size: 20, color: INK_SECONDARY }),
                ],
              }),
              new Paragraph({
                spacing: { after: 80 },
                children: [
                  new TextRun({ text: "Focus — ", bold: true, size: 20, color: INK }),
                  new TextRun({ text: "which of your four Focuses you roll: Courage, Empathy, Wit, or Instinct.", size: 20, color: INK_SECONDARY }),
                ],
              }),
              new Paragraph({
                spacing: { after: 0 },
                children: [
                  new TextRun({ text: "Roll 2d6 + Focus. ", bold: true, size: 20, color: INK }),
                  new TextRun({ text: "10+ is a Strong Hit, 7–9 is a Mixed Result, 6− is a Miss — read across the color bar below every Move.", size: 20, color: INK_SECONDARY }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function outcomeRow(label, color, textColor, text) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 1100, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: color },
        verticalAlign: "center",
        margins: { top: 90, bottom: 90, left: 100, right: 100 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [ new TextRun({ text: label, bold: true, color: textColor, size: 20 }) ],
          }),
        ],
      }),
      new TableCell({
        width: { size: 8980, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: WHITE },
        margins: { top: 90, bottom: 90, left: 160, right: 160 },
        children: [
          new Paragraph({ children: [ new TextRun({ text, color: INK, size: 21 }) ] }),
        ],
      }),
    ],
  });
}

function moveBlock(move) {
  const parts = [];
  parts.push(
    new Paragraph({
      spacing: { before: 260, after: 20 },
      keepNext: true,
      border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: ACCENT, space: 4 } },
      children: [ new TextRun({ text: move.name.toUpperCase(), bold: true, color: INK, size: 30 }) ],
    })
  );
  parts.push(
    new Paragraph({
      spacing: { after: 80 },
      children: [ new TextRun({ text: `FOCUS: ${move.focus.toUpperCase()}`, bold: true, color: ACCENT, size: 18, characterSpacing: 10 }) ],
    })
  );
  parts.push(
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({ text: "Trigger: ", bold: true, italics: true, size: 21, color: INK }),
        new TextRun({ text: move.trigger, italics: true, size: 21, color: INK_SECONDARY }),
      ],
    })
  );
  parts.push(
    new Table({
      width: { size: 10080, type: WidthType.DXA },
      columnWidths: [1100, 8980],
      borders: {
        top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder(),
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "FFFFFF" },
        insideVertical: noBorder(),
      },
      rows: [
        outcomeRow("10+", GOOD, WHITE, move.strong),
        outcomeRow("7–9", WARN, INK, move.mixed),
        outcomeRow("6−", CRIT, WHITE, move.miss),
      ],
    })
  );
  if (move.note) {
    parts.push(
      new Paragraph({
        spacing: { before: 100, after: 60 },
        children: [ new TextRun({ text: move.note, italics: true, size: 19, color: MUTED }) ],
      })
    );
  }
  return parts;
}

const moves = [
  {
    name: "Act Under Pressure",
    focus: "Instinct",
    trigger: "when you have to act fast, with no time to think it through.",
    strong: "You do exactly what you meant to do.",
    mixed: "You do it — but pick one: you hesitate, you're off-balance, or you reveal more than you wanted to.",
    miss: "You freeze, panic, or act on the wrong instinct. The GM decides what happens next.",
  },
  {
    name: "Face Danger",
    focus: "Courage",
    trigger: "when you step into harm's way, on purpose, to get something done.",
    strong: "You handle it — clean, no cost.",
    mixed: "You handle it — but pick one: you get hurt, you lose something, or you have to make a hard choice right now.",
    miss: "The danger wins this round. The GM makes a move against you.",
  },
  {
    name: "Read the Scene",
    focus: "Instinct",
    trigger: "when you stop and look closely at a person, place, or situation before acting.",
    strong: "Ask the GM two questions from the list below. They must answer honestly.",
    mixed: "Ask one question from the list.",
    miss: "The GM asks you a question instead — and you have to answer it out loud, in English.",
    note: "Questions: What's really going on here? What should I watch out for? Who's really in control here? What here isn't what it looks like?",
  },
  {
    name: "Persuade or Manipulate",
    focus: "Wit",
    trigger: "when you work an angle on someone — flattery, logic, a clever half-truth.",
    strong: "They buy it. They do what you want.",
    mixed: "They're close — but they want something from you first.",
    miss: "They catch on. Now they trust you less.",
  },
  {
    name: "Parley",
    focus: "Empathy",
    trigger: "when you make a direct request, backed by something they actually want or need from you.",
    strong: "They give you what you asked for — or a fair trade.",
    mixed: "They'll do it, but there's a catch: a smaller ask, a delay, a condition.",
    miss: "No deal. And they remember you tried.",
  },
  {
    name: "Help or Interfere",
    focus: "Empathy",
    trigger: "when you jump in to support — or block — another player's Move, before the dice are rolled.",
    strong: "They roll with +1. Nothing bad happens to you.",
    mixed: "They roll with +1 — but now you're caught up in it too.",
    miss: "You get in the way instead. They roll with −1.",
  },
];

function chapter4() {
  const children = [];
  children.push(eyebrow("Chapter 4"));
  children.push(chapterTitle("Moves"));
  children.push(bodyPara(
    `Every time your character does something risky, clever, or uncertain, you're making a Move. Moves are the bridge between the story and the dice — they turn “I try to convince the guard” into a real roll with real stakes. There are six Moves everyone shares, no matter what world you're playing in. Your Archetype adds a Signature Move of its own — you'll find those in Chapter 5.`
  ));
  children.push(spacer(160));
  children.push(howToReadBox());
  children.push(spacer(80));

  moves.forEach((m) => { moveBlock(m).forEach(p => children.push(p)); });

  children.push(pageBreak());
  children.push(eyebrow("Chapter 4 — Quick Reference"));
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [ new TextRun({ text: "All Six Moves at a Glance", bold: true, color: INK, size: 36 }) ],
  }));
  children.push(threeColTable(
    ["MOVE", "FOCUS", "TRIGGER"],
    moves.map(m => [m.name, m.focus, m.trigger]),
    [3200, 1800, 5080]
  ));
  children.push(spacer(200));
  children.push(bodyPara(
    "Reminder: any Move is also a chance to use your current Language Focus. If your Move lines up with what you're working on this week, your GM may invite you to narrate it using that structure — and if you nail it, that's a Language Point. See Chapter 8.",
    { italics: true }
  ));
  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 5 — Language Points & Spotlight Tokens
// ---------------------------------------------------------------------------

function chapterResources() {
  const children = [];
  children.push(eyebrow("Chapter 8"));
  children.push(chapterTitle("Language Points & Spotlight Tokens"));

  children.push(flavorQuote(
    `Diego could let the moment pass — nod, say something short, move on. Or he could spend the token burning a hole in his character sheet and make the next thirty seconds entirely his.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Under Resources on your sheet (Chapter 3) sit two small pools of currency: Spotlight Tokens and Language Points. Both refill at the start of every session. Both are entirely yours to spend, on your own timing, with no permission needed. And both exist for the same underlying reason — to put a little more control over the shape of a session, and a little more reward for stretching your English, directly in your hands.`
  ));
  children.push(bodyPara(
    `They're easy to tell apart once you know what each one is for. Spotlight Tokens are about airtime — they buy you room in the scene. Language Points are about the dice — they buy you a second chance on a roll. Neither one makes your Focus rating go up, and neither one is required to play; plenty of good sessions end with tokens or points left unspent.`
  ));

  children.push(sectionHeading("Spotlight Tokens"));
  children.push(bodyPara(
    `Every player starts each session with three Spotlight Tokens — the same three, whether you're just starting out or nearly fluent, no matter your Archetype. You don't earn more of them during play, and whatever's left over at the end of the session doesn't carry over to the next one. Spend them if the moment calls for it; don't hoard them for later.`
  ));
  children.push(bodyPara(
    `Spend one Spotlight Token to claim an extended turn: a bigger beat that's fully about your character, for as long as it takes you to play it out. That might be a longer speech instead of one line, walking the table through your search of the ruins step by step instead of summarizing it, or a short flashback that explains why your character reacts the way they just did. The GM's job, once a token is on the table, is to slow down and let you have the scene.`,
    { after: 100 }
  ));
  children.push(calloutBox(
    "In Practice",
    `A Spotlight Token doesn't change any dice roll — it changes how much time and attention the table gives you. It's there to make room for more English, not to buy you a mechanical edge. If a Move happens during your extended turn, you still roll for it normally.`,
    "clarify"
  ));

  children.push(sectionHeading("Language Points"));
  children.push(bodyPara(
    `You earn a Language Point every time you nail your Language Focus (Chapter 3) in the middle of a scene — your GM awards it on the spot, out loud, the moment it happens. There's no cap on how many you can earn in a session, and like Spotlight Tokens, anything unspent when the session ends is gone; they don't bank forward.`
  ));
  children.push(bodyPara(
    `Spend a Language Point to reroll a 2d6 you've already rolled for a Move. Re-add your Focus to the new roll and read the new total against the three outcome bands (Chapter 2) — even if it's worse. One point buys exactly one reroll; if you want to try again after that, it costs another point.`,
    { after: 100 }
  ));
  children.push(calloutBox(
    "Example",
    `Diego's Language Focus this week is commenting adverbs and the future perfect. Earlier in the session he used one flawlessly while warning the party about a trap — "Fortunately, I will have checked the mechanism before anyone touches it" — and banked a Language Point on the spot. Two scenes later he rolls a Persuade or Manipulate check and lands a 5. He spends the point, rerolls, and gets an 8 instead — a Mixed Result instead of a flat Miss.`,
    "example"
  ));

  children.push(sectionHeading("Two Resources, Side by Side"));
  children.push(bodyPara(
    `A quick reference for both — pin this next to your sheet until it's automatic.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["RESOURCE", "EARNED", "SPENT ON"],
    [
      ["Spotlight Tokens", "3 at the start of every session — flat, for everyone.", "Claiming an extended turn — the scene is yours for a bigger beat."],
      ["Language Points", "1 each time you nail your Language Focus in a live scene.", "Rerolling a 2d6 you've already rolled for a Move."],
    ],
    [2600, 3980, 3500]
  ));
  children.push(spacer(100));
  children.push(bodyPara(
    "Neither pool carries over between sessions — both reset to their starting state (three tokens, zero points) the next time you sit down to play.",
    { italics: true }
  ));

  children.push(spacer(180));
  children.push(sectionHeading("A Turn, With a Language Point in Play"));
  children.push(bodyPara(
    `Picking up right where Diego's example above left off — here's the full turn at the table, start to finish.`,
    { after: 100 }
  ));
  children.push(flavorQuote([
    `GM: "The quartermaster crosses his arms. 'I've heard every sob story in this camp, mage. Why should your lot get the last of the healing draughts?'"`,
    `Diego: "Because when the wounded start arriving tonight, you'll have wished you'd have given them to us instead of watched us fail to save someone who will have been beyond saving otherwise."`,
    `GM: "That's Persuade or Manipulate — roll it."`,
  ]));
  children.push(spacer(100));
  children.push(bodyPara(
    `Step 1 — Roll: Diego rolls 2d6, gets a 3, and adds his Wit of +2. Total: 5 — a Miss.`,
    { after: 80 }
  ));
  children.push(bodyPara(
    `Step 2 — Spend: Diego has a Language Point banked from earlier in the session. He spends it to reroll.`,
    { after: 80 }
  ));
  children.push(bodyPara(
    `Step 3 — Reroll: New 2d6 comes up 6, plus his +2 Wit. New total: 8 — a Mixed Result. The GM narrates the cost: "Fine. Take two — but the healer's watching you both, and she's not the forgiving type." Diego's future-perfect line is what earned him the point in the first place; spending it is what turned a flat Miss into a Mixed Result that keeps the scene moving forward.`,
    { after: 80 }
  ));
  children.push(bodyPara(
    `That's the loop this whole chapter is built on: play your Language Focus well, and the game hands you a tool that makes your next roll a little less likely to go against you. Nobody has to remind you to practice — the incentive is already sitting on your sheet.`,
    { italics: true }
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 6 — Table Etiquette
// ---------------------------------------------------------------------------

function chapterEtiquette() {
  const children = [];
  children.push(eyebrow("Chapter 10"));
  children.push(chapterTitle("Table Etiquette"));

  children.push(flavorQuote(
    `Diego stumbles halfway through a sentence, stops, restarts. Nobody jumps in. Nobody laughs. The GM just waits — because waiting is the rule.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `This chapter is short on purpose — four rules, not forty. ${GAME_NAME} only works if everyone at the table protects the same thing: the space to try, fail, and try again out loud, in front of other people. That takes more deliberate care in a language classroom than it does at an ordinary game night, so these rules exist to spell out exactly what "safe to make mistakes" looks like in practice.`
  ));

  children.push(sectionHeading("Mistakes Are How We Play"));
  children.push(bodyPara(
    `A wrong verb tense, a mixed-up preposition, a sentence that trails off halfway through — none of that stops the story. Nobody at the table corrects you mid-scene, including your GM. If your GM catches something worth fixing, they'll model the correct form right back to you in their own next line of narration — a technique called a recast — instead of pausing to explain a rule. You hear the right version without ever being told you were wrong.`
  ));
  children.push(calloutBox(
    "In Practice",
    `You: "If you letting us in, we show the papers." GM, narrating straight back: "Oren studies you a moment longer, weighing whether he'll let you in once you can prove it..." Same idea, correct grammar, zero interruption. You keep playing — the correction already did its job.`,
    "clarify"
  ));

  children.push(sectionHeading("One Scene, One Voice"));
  children.push(bodyPara(
    `Anyone can jump into someone else's Move — that's what Help or Interfere (Chapter 4) is for. But talking over another player mid-sentence isn't the same thing as helping. Let people finish a thought before adding yours, even when your English comes out faster than theirs.`
  ));

  children.push(sectionHeading("Yes, And"));
  children.push(bodyPara(
    `Borrowed straight from improv: build on what another player adds instead of shutting it down. If Ana says there's a hidden door behind the bookshelf, don't declare there isn't one — find a way to make her addition true, then add something of your own on top of it. The story only gets richer this way, and it keeps the table's energy generous instead of competitive.`
  ));

  children.push(sectionHeading("In Character, In English"));
  children.push(bodyPara(
    `Once a scene is rolling, everything anyone says as their character — dialogue, description, questions to the GM — happens in English. The only exception anywhere in this game is the single line of support on a first-level player's Language Focus card (Chapter 3), and that line retires the moment they finish their first level. If you get stuck mid-sentence, gesture, describe around the word, or just ask your GM "how do I say ___?" in English — that question is part of the game, not a break from it.`
  ));

  children.push(pageBreak());
  children.push(eyebrow("Chapter 10 — Quick Reference"));
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [ new TextRun({ text: "The Four Rules, At a Glance", bold: true, color: INK, size: 36 }) ],
  }));
  children.push(threeColTable(
    ["RULE", "IN ONE LINE"],
    [
      ["Mistakes Are How We Play", "Nobody corrects you mid-scene — your GM recasts the right form back to you instead."],
      ["One Scene, One Voice", "Let a player finish their moment before jumping in."],
      ["Yes, And", "Build on what someone adds to the story — don't shut it down."],
      ["In Character, In English", "Everything you say as your character stays in English, no exceptions but your Language Focus card."],
    ],
    [3000, 7080]
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 7 — Your Responsibilities
// ---------------------------------------------------------------------------

function chapterResponsibilities() {
  const children = [];
  children.push(eyebrow("Chapter 11"));
  children.push(chapterTitle("Your Responsibilities"));

  children.push(flavorQuote(
    `The dice don't care if you studied. But the story remembers who shows up ready.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `${GAME_NAME} only works if you bring a little bit of yourself to the table every time — not talent, just preparation and presence. This chapter covers the three things that are entirely on you, none of which take more than a few minutes.`
  ));

  children.push(sectionHeading("Homework Bonus"));
  children.push(bodyPara(
    `Your regular coursework and your character sheet are connected on purpose. Complete whatever homework you were assigned in your regular class before a session, and check the Homework Bonus box in the Progress section of your sheet (Chapter 3). Once per session, spend it to add +1 to the total of any single roll, after the dice land — enough to turn a Miss into a Mixed Result, or a Mixed Result into a Strong Hit, at exactly the moment you need it.`,
    { after: 100 }
  ));
  children.push(calloutBox(
    "Example",
    `Ana finished her homework the night before. Mid-session, she rolls a 8 on Parley — 2d6 came up 7, plus her Empathy of +1 — a Mixed Result, but she wanted more. She checks her Homework Bonus box, adds +1, and the total becomes 9... still short. She only gets one shot at it per session, so she banks the lesson for next time: save it for a roll that's already close to 10+.`,
    "example"
  ));

  children.push(sectionHeading("Punctuality"));
  children.push(bodyPara(
    `Every minute the table waits on one player is a minute of English practice everyone else loses. Session start times are fixed for a reason — if something's going to make you late, message your GM before the session starts, not after. A few minutes here and there might mean missing your character's opening scene entirely, and that's spotlight you don't get back.`
  ));

  children.push(sectionHeading("Showing Up Ready"));
  children.push(bodyPara(
    `Before you sit down: know what's on your Language Focus card (Chapter 3), know what happened last session, and put your phone somewhere that isn't your hand. None of this takes long, and all of it is the difference between playing the game and watching it happen to you.`,
    { after: 100 }
  ));
  children.push(calloutBox(
    "In Practice",
    `One habit covers all of it: reread your character sheet and the last line of the previous session's recap right before you sit down. That's the whole prep — nothing more is required.`,
    "clarify"
  ));

  children.push(spacer(180));
  children.push(sectionHeading("Your Responsibilities, At a Glance"));
  children.push(threeColTable(
    ["RESPONSIBILITY", "IN ONE LINE"],
    [
      ["Homework Bonus", "+1 to one roll this session, unlocked by finishing your regular coursework."],
      ["Punctuality", "Message your GM ahead of time if you'll be late — the table waits for no one."],
      ["Showing Up Ready", "Know your Language Focus, remember last session, phone away."],
    ],
    [3000, 7080]
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 8 — Session Zero Checklist
// ---------------------------------------------------------------------------

function chapterSessionZero() {
  const children = [];
  children.push(eyebrow("Chapter 12"));
  children.push(chapterTitle("Session Zero Checklist"));

  children.push(flavorQuote(
    `Before a single die is rolled, the table decides — together — what kind of story this is going to be.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Session Zero is the meeting — or the first block of your very first session — where everyone at the table builds characters and aligns expectations before anyone opens a scene. It trades guessing for a shared understanding, so the story can start strong instead of stalling out on logistics five minutes in. Work through this checklist, in order, before you play.`
  ));

  children.push(threeColTable(
    ["ALIGN ON", "WHAT THAT MEANS"],
    [
      ["Setting & Tone", "Which of the worlds in Chapter 9 is the table playing, and which campaign shape — Long Haul, Chronicle, or Anthology (Chapter 5)? What's in bounds, and what's off the table content-wise?"],
      ["Characters", "Everyone builds their sheet together — assign your Focus array (Chapter 3), then pick an Archetype and Signature Move (Chapter 5) and read the Kit that comes with it (Chapter 6). Chapter 7 walks through one from start to finish."],
      ["Growth Levels", "A brand-new table starts everyone at Level 1, whatever their English level (Chapter 5). A table taking in a player who has played here before reads their Growth Ledger. Levels may differ around the table — that's normal and changes nothing about the dice."],
      ["Language Focus", "Your GM assigns each player's starting Language Focus, pulled from whatever you're covering in regular class."],
      ["Schedule & Attendance", "Confirm the session day and time, and revisit the Punctuality expectation (Chapter 11)."],
      ["Table Etiquette Recap", "A quick read-through of the four rules in Chapter 10, out loud, together."],
      ["Resources", "Confirm that Spotlight Tokens and Language Points (Chapter 8) reset every session, and how the Homework Bonus (Chapter 11) works."],
      ["Comfort Check-In", `Agree on a simple word — "pause" works fine — that any player can say to skip or soften content that's making them uncomfortable. No explanation required in the moment.`],
      ["First Scene", "Agree on the opening image: where the story starts, and who's there."],
    ],
    [2600, 7480]
  ));

  children.push(spacer(180));
  children.push(sectionHeading("That's the Whole Book"));
  children.push(bodyPara(
    `Cover to cover, everything you need is behind you now — the 2d6 engine (Chapter 2), your character (Chapter 3), your Moves (Chapter 4), your Archetype and the twelve-level track it climbs (Chapter 5), everything you carry (Chapter 6), your resources (Chapter 8), the worlds you can play in (Chapter 9), and the etiquette and responsibilities that keep the table running (Chapters 10 and 11). Every future Setting Guide builds on top of exactly what you already know — new worlds, new dressing, the same four Focuses and the same six Moves underneath. Session Zero is the last thing standing between you and the lantern gutters, the gate, and Oren's hand on the pommel of his sword. Roll the dice.`,
    { italics: true }
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 5 — Archetypes & Growth
// ---------------------------------------------------------------------------
// Design note: growth is deliberately "toolbox-only". Nothing in this chapter
// ever raises a Focus number or adds a flat bonus to a roll, because Ch.2 and
// Ch.3 lock in the promise that the dice math is identical for every student
// regardless of English level. The five Guardrails below are the load-bearing
// rules that keep a Level 12 character interesting instead of unbeatable.
// ---------------------------------------------------------------------------

const archetypes = [
  {
    name: "The Vanguard",
    concept: "Steps into danger first, shields the people behind them, acts when talk runs out.",
    array: [["Courage", "+2"], ["Instinct", "+1"], ["Empathy", "+0"], ["Wit", "−1"]],
    moveName: "Shield the Line",
    tiers: [
      ["Tier 1 — Level 1", "Shield the Line", "When you Face Danger to protect another character, on a 7–9 you choose the complication yourself instead of the GM."],
      ["Tier 2 — Level 3", "Shield the Line", "As Tier 1, and on a 10+ the character you protected may immediately take an action of their own before the scene moves on."],
      ["Tier 3 — Level 7", "Hold the Door", "As Tier 2, and once per session, when an ally would suffer a consequence from a Miss, you may take it in their place instead."],
      ["Tier 4 — Level 12", "Nothing Gets Past Me", "As Tier 3, and whenever you take a consequence in an ally's place, describe the cost out loud, in character, in English. If you do, the whole table banks one shared reroll to spend before the session ends."],
    ],
  },
  {
    name: "The Diplomat",
    concept: "Reads what people actually want, negotiates instead of confronting, turns enemies into leverage.",
    array: [["Empathy", "+2"], ["Wit", "+1"], ["Instinct", "+0"], ["Courage", "−1"]],
    moveName: "Read the Room",
    tiers: [
      ["Tier 1 — Level 1", "Read the Room", "Before you decide what to offer in a Parley, you may ask your GM what the other party actually wants — they answer honestly."],
      ["Tier 2 — Level 3", "Read the Room", "As Tier 1, and you may instead ask what the other party is afraid of. Your choice, one question, answered honestly before you commit."],
      ["Tier 3 — Level 7", "Everyone Has a Price", "As Tier 2, and once per session, treat a 7–9 on Parley as a 10+ instead."],
      ["Tier 4 — Level 12", "Speak for the Table", "As Tier 3, and when you Parley on behalf of another player's character rather than your own — and they narrate their half of the offer in English too — you both earn a Language Point, whatever the dice say."],
    ],
  },
  {
    name: "The Strategist",
    concept: "Outthinks the problem before it becomes a fight — always three steps ahead, the clever one at the table.",
    array: [["Wit", "+2"], ["Instinct", "+1"], ["Empathy", "+0"], ["Courage", "−1"]],
    moveName: "Angles and Openings",
    tiers: [
      ["Tier 1 — Level 1", "Angles and Openings", "Once per session, before you roll Persuade or Manipulate, you may reroll one of your two dice."],
      ["Tier 2 — Level 3", "Angles and Openings", "As Tier 1, and the reroll is no longer limited to once per session — but every use after the first costs you one Language Point."],
      ["Tier 3 — Level 7", "Already Knew That", "As Tier 2, and once per session you may ask your GM one question about the current scene without rolling Read the Scene. They answer honestly, at no cost."],
      ["Tier 4 — Level 12", "I Planned for This", "As Tier 3, and at the start of a session you may name one thing you expect to go wrong. If it does, every player at the table may reroll one die during that scene."],
    ],
  },
  {
    name: "The Scout",
    concept: "Notices danger before anyone else does, moves quiet, trusts their gut over their plan.",
    array: [["Instinct", "+2"], ["Courage", "+1"], ["Wit", "+0"], ["Empathy", "−1"]],
    moveName: "Nothing Slips By",
    tiers: [
      ["Tier 1 — Level 1", "Nothing Slips By", "Even on a Miss, Read the Scene still lets you ask one question — the GM just answers it in a way that costs you something."],
      ["Tier 2 — Level 3", "Nothing Slips By", "As Tier 1, and on a 10+ you may aim one of your questions at a scene that hasn't happened yet — somewhere your character is headed."],
      ["Tier 3 — Level 7", "One Step Ahead", "As Tier 2, and once per session you may act first in a scene, ahead of where turn order would normally put you."],
      ["Tier 4 — Level 12", "Follow Me", "As Tier 3, and when you act first you may bring one other player with you. They act immediately after you — and describe what they see, in English, before they roll."],
    ],
  },
];

function archetypeBlock(a) {
  const parts = [];
  parts.push(sectionHeading(a.name));
  parts.push(bodyPara(a.concept, { italics: true, after: 100 }));
  parts.push(threeColTable(
    ["FOCUS", "SUGGESTED VALUE"],
    a.array,
    [5040, 5040]
  ));
  parts.push(spacer(120));
  parts.push(new Paragraph({
    keepNext: true,
    spacing: { after: 80 },
    children: [
      new TextRun({ text: "Signature Move track — ", bold: true, size: 21, color: INK }),
      new TextRun({ text: `${a.moveName}`, bold: true, italics: true, size: 21, color: ACCENT }),
      new TextRun({ text: " and what it becomes. Each tier replaces the one above it; you never carry more than one.", size: 21, color: INK_SECONDARY }),
    ],
  }));
  parts.push(threeColTable(
    ["UNLOCKS AT", "MOVE", "WHAT IT DOES"],
    a.tiers,
    [1700, 2200, 6180]
  ));
  return parts;
}

// The full career ladder. 12 Growth Levels, 11 Growth Moments between them,
// plus a Capstone for finishing the last half-book of the Evolve track.
const growthLadder = [
  ["Level 1", "the day you sit down", "Archetype, Focus array, Signature Move at Tier 1."],
  ["Level 2", "6 units", "A Boon."],
  ["Level 3", "12 units", "Signature Move steps up to Tier 2."],
  ["Level 4", "18 units", "A Boon."],
  ["Level 5", "24 units", "Cross-Training."],
  ["Level 6", "30 units", "A Boon."],
  ["Level 7", "36 units", "Signature Move steps up to Tier 3."],
  ["Level 8", "42 units", "A Boon."],
  ["Level 9", "48 units", "A Focus Shift."],
  ["Level 10", "54 units", "A Boon."],
  ["Level 11", "60 units", "Cross-Training."],
  ["Level 12", "66 units", "Signature Move steps up to Tier 4."],
  ["Capstone", "72 units", "A Legacy Boon — and your character's story closes, or hands itself to someone new."],
];

const growthKinds = [
  ["Boon", "A narrative gain, defined by you and your GM out of what has already happened in your story: an item, an ally, a place that will open its door to you, a reputation that precedes you. A Boon can open a door, change a fact, or give your GM a reason to say yes. It never adds a number to a roll."],
  ["Signature Move tier", "Your Archetype's Signature Move, rewritten sharper. The new tier replaces the old text on your sheet. You always have exactly one Archetype Signature Move, at every level, from Level 1 to Level 12."],
  ["Cross-Training", "Take the Tier 1 Signature Move of a different Archetype and add it to your sheet permanently. It stays at Tier 1 forever — it never upgrades. You may do this twice in a career, and never from the same Archetype twice."],
  ["Focus Shift", "Swap the values of any two of your Focuses. Your +2 and your −1 can trade places, or your +1 and your 0. Nothing goes up; something moves. This is your character changing, not your character improving."],
  ["Legacy Boon", "The one growth that outlives the character. It attaches to your table, not your sheet — a place named after you, an organisation you founded, a debt the world owes you — and it stays in the world for whoever plays there next, including you, with a new character."],
];

const guardrails = [
  ["You start at Level 1. Everybody does.", "Your Growth Level does not come from how good your English already is. A student who joins at a high level and a student who joins from zero both sit down at Level 1, with no Boons. What you knew before you got here is not something this table gets to reward — only what you do after."],
  ["Focus numbers never go up.", "The only thing that ever happens to your array is a Focus Shift, which moves a number without creating one. A Level 12 character and a Level 1 character roll against exactly the same odds. This is the promise from Chapter 2, kept all the way to the end of the track."],
  ["You only ever have one Archetype Signature Move.", "Tiers replace each other. Your sheet at Level 12 is not four Moves deep — it is one Move, four times sharper."],
  ["Cross-Training stays at Tier 1, twice, forever.", "Borrowed Moves give you range, not depth. Two of them, from two different Archetypes, and neither one ever upgrades."],
  ["Nothing may add a flat bonus to a roll.", "No growth in this game ever hands you a +1. The Homework Bonus (Chapter 11) is the only bonus of its kind that exists, it comes from studying, and it is capped at once per session."],
  ["One session is still one session.", "Almost every Signature Move is once per session. A Level 12 character holding three once-per-session tools still only gets three moments of leverage in a two-hour game. What grows is the number of interesting choices, not the size of the numbers."],
];

const campaignShapes = [
  ["The Long Haul", "One world, one character, the whole ladder. Every Growth Moment lands on the same person, and by the Capstone that character has a twelve-level history the table watched happen. Best for a stable group that wants one long story. This is the default."],
  ["The Chronicle", "The table plays in arcs. When an arc ends — or when the group simply wants a different world — everyone retires their character and builds a new one at their current Growth Level. Same level, same tier, same number of Boons. Only the fiction is new: each Boon gets reimagined for the new world, keeping what it does and changing what it is. A sword becomes a contact; a contact becomes an access code."],
  ["The Anthology", "Short settings, two to four sessions each, new characters every time. The Growth Ledger does all the work: whatever your level is when a new setting opens, that is what you build at. Best for tables that want to sample a lot of worlds, and the easiest shape to run alongside a class calendar."],
];

function chapterArchetypes() {
  const children = [];
  children.push(eyebrow("Chapter 5"));
  children.push(chapterTitle("Archetypes & Growth"));

  children.push(flavorQuote(
    `Every character starts somewhere. What matters more is where a year of real study takes them.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Your Archetype is the role your character plays at the table — how they tend to solve problems, and which Focus they usually lean on to do it. The four Archetypes in this chapter work in absolutely any setting, no matter which of the worlds in Chapter 9 your table picks. As Setting Guides get written, each will add its own world-flavored Archetypes on top of these — you already glimpsed one in Chapter 3's example, Diego's Scholar-Mage, which is exactly what a Fantasy-flavored Archetype looks like. The four below are different: they belong to no single world, so they are playable today, anywhere.`
  ));
  children.push(bodyPara(
    `One thing this chapter deliberately leaves out: species or ancestry — elf, android, ghost, whatever a given world calls for. That is setting flavor rather than a universal role, so it stays with each Setting Guide instead of living here.`
  ));

  children.push(sectionHeading("The Four Archetypes"));
  children.push(bodyPara(
    `Each Archetype suggests a Focus array to match its concept — but exactly as Chapter 3 says, the suggestion is a starting point, not a requirement. Split your numbers however fits the character you actually want to play. Each Archetype also comes with one Signature Move, and a track showing what that Move becomes as you grow. You get the Tier 1 version free, the moment you pick the Archetype.`,
    { after: 100 }
  ));

  archetypes.forEach((a, i) => {
    archetypeBlock(a).forEach(p => children.push(p));
    // Two Archetypes per page — each block is a heading plus two tables, and
    // letting them flow freely orphans single rows across the page boundary.
    if (i < archetypes.length - 1) children.push(pageBreak());
  });

  // -------------------------------------------------------------------------
  children.push(pageBreak());
  children.push(sectionHeading("Growth: Leveling Up With Your English"));
  children.push(bodyPara(
    `${GAME_NAME} ties character growth to something almost no other game can reach: the work you do on your own English. Finish six units of your course, pass the test that closes them, and your character hits a Growth Moment — no matter how many sessions that took, and no matter what happened at the table. Studying is what makes your character grow. Nothing else does.`
  ));
  children.push(bodyPara(
    `Every Growth Moment also retires your current Language Focus and replaces it with a new one, drawn by your GM from the material you are moving into. That part is automatic — it is not one of your choices, it is the reason the choices exist.`
  ));

  children.push(sectionHeading("The Twelve-Level Track"));
  children.push(bodyPara(
    `The ladder has twelve rungs. Eleven Growth Moments sit between Level 1 and Level 12, and one last Capstone waits at the top. Every rung costs the same thing: six more units finished.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["LEVEL", "UNITS DONE HERE", "WHAT YOU GAIN"],
    growthLadder,
    [1500, 2800, 5780]
  ));

  children.push(spacer(180));
  children.push(calloutBox(
    "In Practice",
    `Notice what the ladder alternates. Six of the eleven Growth Moments give you fiction — a Boon, then another Boon — and only five touch the rules at all. That rhythm is on purpose: it keeps growth constant without letting mechanics pile up.`,
    "clarify"
  ));

  children.push(pageBreak());
  children.push(sectionHeading("Everyone Starts at Level 1"));
  children.push(bodyPara(
    `Read the second column of that table again. It says units done here — not which book you are in, and not how good your English is.`
  ));
  children.push(bodyPara(
    `Four people can sit down at the same table on the same day with four different English levels. One is on their first unit ever; one is halfway through the third book; one is nearly finished. All four start at Growth Level 1, with no Boons, with their Signature Move at Tier 1. What you already knew when you walked in is not something this game pays you for.`
  ));
  children.push(bodyPara(
    `That is not a way of holding strong students back. It is what makes a Boon mean something. A Boon is proof of work you did here, at this table, over months — and a reward you were handed for a test you took somewhere else, years ago, would not be proof of anything.`
  ));
  children.push(calloutBox(
    "And when the books run out",
    `Six units is six units, whatever you are studying. A student who joins at a high level will finish the last Evolve book long before they reach Level 12 — and then they simply keep counting, six units at a time, through whatever course comes next. The ladder belongs to the game, not to any one book.`,
    "clarify"
  ));

  children.push(spacer(240));
  children.push(sectionHeading("The Five Kinds of Growth"));
  children.push(bodyPara(
    `Every rung on the ladder hands you one of these five things, and nothing else.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["GROWTH", "WHAT IT IS"],
    growthKinds,
    [2400, 7680]
  ));

  children.push(spacer(200));
  children.push(sectionHeading("Why a Level 12 Character Is Not Unbeatable"));
  children.push(bodyPara(
    `A twelve-level track sounds like the kind of thing that ends with an untouchable character. It does not, and the reason is five rules that hold from the first session to the last.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["THE GUARDRAIL", "WHAT IT PROTECTS"],
    guardrails,
    [3400, 6680]
  ));

  children.push(spacer(180));
  children.push(calloutBox(
    "The Short Version",
    `A Level 12 character has one very sharp Signature Move, two borrowed ones, a rearranged Focus array, and a pile of history. They still roll 2d6. They still Miss on a 6 or under. Nothing on the ladder was ever aimed at the dice — it was aimed at how many interesting things you can choose to do before you roll them.`,
    "clarify"
  ));

  children.push(spacer(180));
  children.push(calloutBox(
    "Example",
    `Ana — the player from Chapter 2's gate scene — is at Growth Level 2 with her Diplomat, Mira. She finishes her twelfth unit and passes the test: Growth Moment. That rung is a Signature Move upgrade, so Read the Room sharpens to Tier 2, and from now on she can ask her GM what an NPC is afraid of instead of what they want. Her Empathy is still +1. It was +1 last month, and it will be +1 at Level 12.`,
    "example"
  ));

  // -------------------------------------------------------------------------
  children.push(spacer(240));
  children.push(sectionHeading("Your Growth Level Belongs to You, Not Your Character"));
  children.push(bodyPara(
    `This is the rule that makes everything else portable, so it is worth stating plainly: Growth Levels are earned by the student, not by the character. They come from your coursework. Your coursework does not reset when a campaign ends, so neither does your level.`
  ));
  children.push(bodyPara(
    `This is why the Growth Ledger is not part of your character sheet, and must never be kept there. Your character sheet is temporary: it belongs to one character, in one world, in one campaign, and the day your table starts something new it gets replaced. Your Ledger is permanent. It lists your Growth Level, how many units you have finished here, every Archetype you have played, every character you have played them as, and every Boon you have ever earned.`
  ));

  children.push(bodyPara(
    `Your GM keeps the Ledger for the whole table, in one place that is not the character sheets — a shared page, a spreadsheet, a notebook. You should be able to read yours at any time, and you should never be the only copy of it.`
  ));
  children.push(calloutBox(
    "Example",
    `Ana has been playing for a year. Her table finishes its Fantasy campaign and votes to start again in Cyberpunk. Mira the Diplomat is retired, and Ana builds someone new. The character sheet is blank again — new name, new Kit, new Focus labels. Her Ledger is not: it still reads Growth Level 3, two Boons earned, Archetypes played: Diplomat. She builds her new character at Level 3, with her Signature Move already at Tier 2, and reimagines both Boons for the new world. Nothing she earned went anywhere.`,
    "example"
  ));
  children.push(calloutBox(
    "The one thing that would break this",
    `If your Growth Level lived only on your character sheet, retiring that character would delete a year of your work. That is the whole reason for the separation — not bookkeeping neatness, but making sure the table can change worlds as often as it likes without anybody paying for it.`,
    "warn"
  ));

  children.push(sectionHeading("Three Shapes a Campaign Can Take"));
  children.push(bodyPara(
    `Tables run at different speeds and get restless at different rates, and a track this long has to survive a group deciding, at session forty, that it wants to play something else. Pick one of these three shapes at Session Zero — and pick it again whenever the table wants to change.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["SHAPE", "HOW IT RUNS"],
    campaignShapes,
    [2400, 7680]
  ));

  children.push(spacer(200));
  children.push(sectionHeading("Rebuilding at Level"));
  children.push(bodyPara(
    `Whenever a Chronicle turns over or an Anthology setting opens, everyone rebuilds. The procedure is short. Take your Growth Level from your Ledger. Pick an Archetype — the one you had, or a different one. Assign a fresh Focus array, applying any Focus Shifts you have earned. Set your Signature Move to the tier your level says you have reached. Reclaim your Cross-Trained Moves if you have any. Then, with your GM, reimagine each Boon you have earned for the new world: a Boon keeps what it does and changes what it is.`
  ));
  children.push(calloutBox(
    "In Practice",
    `A player at Level 8 who moves from Fantasy to Cyberpunk does not restart. Their sword — a Boon meaning “a weapon nobody argues with” — becomes a corporate ID that nobody questions. Same function, new world. Their Signature Move is still Tier 3, because they still did the forty-two units that earned it.`,
    "clarify"
  ));

  children.push(spacer(180));
  children.push(pageBreak());
  children.push(sectionHeading("Different Levels at the Same Table"));
  children.push(bodyPara(
    `On day one, every table is level. Everybody is at Level 1, whatever book they are studying from. What splits the table later is not English — it is time and work. The player who has been here two years is further up the ladder than the player who arrived in March, and that is the whole point of the ladder.`
  ));
  children.push(bodyPara(
    `So yes, you will end up with a Level 10 character sitting next to a Level 3 one. That is fine, for the same reason everything else in this book is fine: none of it touches the dice. The Level 10 player has more tools to reach for. Neither player has better odds. The Level 3 player is not behind — they are simply earlier, on exactly the same road.`
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 6 — Making a Character (Ana's worked example)
// ---------------------------------------------------------------------------
// Purpose: everything needed to build a character is already in Ch.3 and Ch.5,
// but stated as rules. This chapter states it as a decision sequence, narrated,
// so a student with creative block has something to imitate instead of a blank
// sheet to stare at.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// CHAPTER 6 — WHAT YOU CARRY
// Gear, Boons, money and distance. The governing law of this chapter is that
// nothing here ever touches the 2d6 math (see Ch.2 and the equity rule): owning
// things changes what is POSSIBLE, never what is PROBABLE.
// ---------------------------------------------------------------------------

const priceExamples = [
  ["A hot meal, a bed for the night", "a coin"],
  ["A good rope, a lantern, a warm coat", "a coin"],
  ["A decent weapon, a week of lodging", "a handful"],
  ["A horse, a forged document, a bribe that works", "a bag"],
  ["A house, a ship, a name that opens doors", "a chest"],
];

const distanceLadder = [
  ["Within reach", "Close enough to touch. You can hand something over, or grab it."],
  ["Nearby", "Same room, a few steps away. You can speak normally and be heard."],
  ["Far away", "Across the hall, the street, the clearing. You have to move to get there."],
  ["Out of sight", "Behind a door, around the corner, gone. You cannot act on it at all."],
];

function chapterGear() {
  const children = [];
  children.push(eyebrow("Chapter 6"));
  children.push(chapterTitle("What You Carry"));

  children.push(flavorQuote([
    `Mira has a sealed letter she has not opened, a silver ring she will not explain, and four empty slots in her pack.`,
    `The ring is not magic. It just belongs to somebody who is going to want it back.`,
  ]));
  children.push(spacer(140));

  children.push(bodyPara(
    `Your character owns things. Some of it came with them, some they picked up along the way, and some was handed to them as a reward for how far they have come. All of it lives in four short blocks on your sheet, and the whole lot takes about ten seconds to read.`
  ));

  children.push(sectionHeading("The Rule Behind This Entire Chapter"));
  children.push(bodyPara(
    `Nothing you own changes the dice. Not your Kit, not the thing you took off a table last session, not a Boon, not a chest full of gold. There is no +1 anywhere in this chapter, and there never will be.`
  ));
  children.push(bodyPara(
    `What your possessions change is what is possible. A rope means you can climb the wall. A fine coat means the guard at the gate looks at your face instead of your boots. A letter with the right seal means the door opens without a word. None of those make you roll better — they change which situations come up, and which ones you can walk straight past without rolling at all.`
  ));
  children.push(calloutBox(
    "Why it works this way",
    `The player who happens to find the better sword should not roll better than the player who did not. Your dice belong to you and your Focuses, and nothing you can buy, loot or be given is allowed to touch them. That is the same promise the game makes about your English level in Chapter 2, kept one more time.`,
    "clarify"
  ));

  // -------------------------------------------------------------------------
  children.push(sectionHeading("Your Kit"));
  children.push(bodyPara(
    `Your Kit is the four or five things your character always has, printed on your sheet the day you make them. It comes from your Archetype and from the world you are playing in — a Scout in a Fantasy setting carries a rope, a knife, a lantern and a water bottle; a Scout in a Cyberpunk setting carries a grapple line, a knife, a flashlight and a grey hoodie. Same role, same job, different century.`
  ));
  children.push(bodyPara(
    `Your Kit never runs out and never needs tracking. You do not spend it, you do not count it, and you do not lose it unless the story takes it from you in a scene everybody watched happen.`
  ));
  children.push(bodyPara(
    `It exists so you never have to wonder whether you have the obvious thing. You look at your sheet, and you know. And because the items are already written down in English, a Kit item is a sentence you already know how to start: I use my rope to reach the window. I don't have a lantern — does anyone?`,
    { italics: false }
  ));

  // -------------------------------------------------------------------------
  children.push(sectionHeading("Your Pack — Six Slots"));
  children.push(bodyPara(
    `Everything else your character picks up, buys, steals or is handed goes into your Pack, and your Pack has exactly six slots.`
  ));
  children.push(bodyPara(
    `One thing per slot, whatever its size. A knife is a slot. A lantern is a slot. A stolen painting the size of a door is a slot. The game is not interested in weight, and neither are you — six is six.`
  ));
  children.push(bodyPara(
    `Your Kit does not use slots. Your Boons do not use slots. The six are only for what you have gathered since the story began.`
  ));

  children.push(sectionHeading("Dropping and Taking"));
  children.push(bodyPara(
    `When your Pack is full and you want something new, something old has to go. Before you take the new thing, say out loud — in English — what you are leaving behind and where you are leaving it.`
  ));
  children.push(bodyPara(
    `That sentence is the entire reason the six slots exist. A full pack in front of an open vault is the most reliable argument this game knows how to start, and it is a good argument: everyone at the table has an opinion, everyone has a reason, and all of it happens in English while the torches burn down. Your GM does not decide for you. Your table does.`
  ));
  children.push(calloutBox(
    "At the table",
    `“I'll leave the crowbar by the door — we can come back for it. I want the book.” That is a complete turn. Nobody rolled anything, and four people just negotiated in a second language about something they cared about.`,
    "example"
  ));

  // -------------------------------------------------------------------------
  children.push(spacer(200));
  children.push(sectionHeading("Boons"));
  children.push(bodyPara(
    `A Boon is not bought and not found. It is earned at a Growth Moment, when you cross into a new half of your course — the twelve-level track in Chapter 5. It is the only kind of possession in this game that arrives because of work you did outside the fiction.`
  ));
  children.push(bodyPara(
    `Boons live in their own block on your sheet. They do not take Pack slots, they are not lost in an ordinary scene, and they are always narratively strong: a door that opens for you, a person who owes you a favour and knows it, an object that does one impossible thing once. Strong, and still not a bonus to any roll.`
  ));
  children.push(bodyPara(
    `When your table rebuilds for a new world — a new Chronicle, a new Anthology setting, the procedure in Chapter 5 — your Boons come with you. A Boon keeps what it does and changes what it is. The blade that never dulls becomes the pistol that never jams; the debt an old captain owes you becomes a debt owed by somebody with a very different rank and the same expression.`
  ));

  // -------------------------------------------------------------------------
  children.push(sectionHeading("Money"));
  children.push(bodyPara(
    `Money in this game has four steps, and you will never do arithmetic with it.`
  ));
  children.push(threeColTable(
    ["STEP", "WHAT IT MEANS"],
    [
      ["A coin", "Enough for a meal, a bed, or one ordinary useful thing."],
      ["A handful", "Ten coins. Enough to matter for a week."],
      ["A bag", "Ten handfuls. Enough to change what you are able to attempt."],
      ["A chest", "Ten bags — and as much as any one person can carry. Spend it, store it, or give it away."],
    ],
    [3000, 7080]
  ));
  children.push(spacer(160));
  children.push(bodyPara(
    `Mark the boxes on your sheet as you go. When ten coins fill up, rub them out and mark one handful instead — the same way ten of anything becomes one of the next thing up, in any language you have ever counted in.`
  ));
  children.push(bodyPara(
    `Prices are spoken in steps, never calculated. Nothing in this game costs one hundred and thirty-seven of anything. A merchant says a sword costs two handfuls; you say that is too expensive; the two of you find out who is more stubborn. That conversation is the point of having money at all.`
  ));
  children.push(pageBreak());
  children.push(threeColTable(
    ["ROUGHLY WHAT THINGS COST", "STEP"],
    priceExamples,
    [7080, 3000]
  ));
  children.push(spacer(160));
  children.push(calloutBox(
    "Every world, same four steps",
    `Whatever your setting calls its money — gold, credits, scrip, cash, funds — the four steps never change their names. A handful of credits and a bag of scrip work exactly like a handful of gold and a bag of gold. You learn the ladder once and it follows you into every world you ever play in.`,
    "clarify"
  ));

  // -------------------------------------------------------------------------
  children.push(spacer(200));
  children.push(sectionHeading("One Last Measurement: How Far Is It?"));
  children.push(bodyPara(
    `Distance works the same way money does — in words, not numbers. This game has no metres, no feet and no squares on a grid. It has four steps, and they are the only four you need.`
  ));
  children.push(threeColTable(
    ["DISTANCE", "WHAT IT MEANS"],
    distanceLadder,
    [3000, 7080]
  ));
  children.push(spacer(160));
  children.push(bodyPara(
    `Going from Within reach to Nearby is free — take a step and keep talking. Going from Nearby to Far away, or back, uses your moment in the scene; and if something is actively trying to stop you, it is a Move and the dice come out. Out of sight is not a distance so much as a wall: you cannot act on what you cannot see, and getting eyes on it again is its own problem to solve.`
  ));
  children.push(bodyPara(
    `Ask your GM how far something is and you will get one of those four words back, every time. Not eighteen metres. Nearby.`
  ));

  // -------------------------------------------------------------------------
  children.push(pageBreak());
  children.push(eyebrow("Chapter 6 — Quick Reference"));
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [ new TextRun({ text: "Everything You Own, At a Glance", bold: true, color: INK, size: 36 }) ],
  }));
  children.push(threeColTable(
    ["BLOCK", "WHAT IT IS"],
    [
      ["Kit", "Four or five items from your Archetype and setting. Always with you. Never tracked, never counted."],
      ["Pack", "Six slots for everything you have gathered. One thing per slot, whatever its size."],
      ["Boons", "Earned at Growth Moments only. No slots. Narratively strong. They follow you into new worlds."],
      ["Money", "A coin, a handful, a bag, a chest. Ten of each makes one of the next. A chest is the ceiling."],
      ["Distance", "Within reach, Nearby, Far away, Out of sight. There are no other distances."],
      ["The one rule", "Nothing you own changes the dice — only what is possible, never what is probable."],
    ],
    [2600, 7480]
  ));

  return children;
}

const anaSheet = [
  ["Name", "Mira"],
  ["Archetype", "The Diplomat"],
  ["Courage", "−1"],
  ["Empathy", "+1"],
  ["Wit", "+0"],
  ["Instinct", "+2"],
  ["Signature Move", "Read the Room (Tier 1)"],
  ["Language Focus", "Past simple — narrating what happened"],
  ["Kit", "A sealed letter · a silver ring · a warm cloak · a small mirror"],
  ["Pack", "Six slots, all empty — she has not been anywhere yet"],
  ["Growth Level", "1 — her first day here"],
  ["Resources", "3 Spotlight Tokens · 0 Language Points · Homework Bonus unchecked"],
];

function chapterAna() {
  const children = [];
  children.push(eyebrow("Chapter 7"));
  children.push(chapterTitle("Making a Character"));

  children.push(flavorQuote(
    `The hardest roll in any roleplaying game is the first one, and it happens before you touch the dice.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Everything you need to build a character is already behind you. Chapter 3 gave you the Focuses; Chapter 5 gave you the Archetypes. What neither of them gave you is the feeling of sitting down with a blank sheet and no idea where to start — which is where most players actually get stuck, in every game, in every language.`
  ));
  children.push(bodyPara(
    `So here is one, built in front of you. Ana is a real student at an A2 level, and this is how she made Mira in about ten minutes at her table's Session Zero. Copy her order of operations, not her answers.`
  ));

  children.push(sectionHeading("Step 1 — Start With a Person, Not a Number"));
  children.push(bodyPara(
    `Ana does not open her sheet first. She asks herself one question: in a story, what would I actually enjoy doing? Not what is strongest — what is fun. Her answer is immediate, and it is talking. She likes the scenes where somebody has to be convinced, or calmed down, or read correctly before it is too late. She does not particularly want to fight anything.`
  ));
  children.push(calloutBox(
    "In Practice",
    `If you cannot answer this question, answer a smaller one: name a character from a book, film, or series whose scenes you always look forward to. You are not copying them. You are just finding out what kind of scene you want more of.`,
    "clarify"
  ));

  children.push(sectionHeading("Step 2 — Pick the Archetype That Matches the Answer"));
  children.push(bodyPara(
    `With “talking” as her answer, Ana reads the four Archetypes in Chapter 5 and stops at The Diplomat: reads what people actually want, negotiates instead of confronting, turns enemies into leverage. That is the description of the scenes she just said she wanted. She does not agonise over the other three. The Archetype is a starting point, not a life sentence — and at a Chronicle turnover she could pick a different one anyway.`
  ));

  children.push(sectionHeading("Step 3 — Assign Your Focuses, and Feel Free to Disagree"));
  children.push(bodyPara(
    `The Diplomat suggests Empathy +2, Wit +1, Instinct +0, Courage −1. Ana reads it and does not take it — not because she is optimising, but because the array describes a different person than the one in her head. Her Diplomat is not a clever negotiator who out-argues people. Her Diplomat reads a room the second she walks into it and works from her gut.`
  ));
  children.push(bodyPara(
    `So she moves two numbers: Instinct +2, Empathy +1, Wit +0, Courage −1. Same four values, in a different order, because that is the only thing you are ever allowed to do with them. Her −1 stays in Courage, and she is glad it does — a character who is bad at something is a character with scenes worth watching.`
  ));
  children.push(calloutBox(
    "The Rule Behind This",
    `You always place the same four numbers: +2, +1, +0, and −1. Every player at every table at every level does. Where you put them is the entire choice, and it is the only part of your sheet that will never change on its own.`,
    "clarify"
  ));

  children.push(pageBreak());
  children.push(sectionHeading("Step 4 — Take Your Signature Move (You Do Not Choose It)"));
  children.push(bodyPara(
    `Read the Room comes with the Diplomat, free, at Tier 1: before Ana decides what to offer in a Parley, she may ask her GM what the other party actually wants, and get an honest answer. There is nothing to pick here — the Archetype hands it over.`
  ));
  children.push(bodyPara(
    `Ana is at Growth Level 1, like everyone else at the table on their first day. Her English is stronger than Leo's and weaker than Tiago's, and none of that shows up here. Six units from now she gets her first Boon.`
  ));

  children.push(sectionHeading("Step 5 — Your Language Focus Is Assigned, Not Chosen"));
  children.push(bodyPara(
    `This is the one line on the sheet Ana does not write herself. Her GM fills it in from what she is covering in her regular class this week: past simple, for narrating what happened. From now on, whenever a scene calls for Mira to explain events that have already occurred, the GM will aim the spotlight at Ana specifically and ask her to tell it in that structure. Nail it and that is a Language Point (Chapter 8). It will change every few weeks, and it changes automatically at every Growth Moment.`
  ));

  children.push(sectionHeading("Step 6 — A Name and One True Detail"));
  children.push(bodyPara(
    `Ana names her Mira, and writes one sentence underneath: “She used to translate for a merchant caravan, and she has never once been paid what she was promised.” That is the whole backstory. It is enough. It tells her GM what Mira is good at, why she talks to strangers for a living, and exactly what will make her angry — three hooks out of one sentence, which is more than most three-page histories manage.`
  ));
  children.push(calloutBox(
    "In Practice",
    `One true detail beats a full biography, especially in a language you are still building. A sentence you can say confidently in English at the table is worth more than a paragraph you wrote with a dictionary and will never use.`,
    "clarify"
  ));

  children.push(sectionHeading("Step 7 — Your Kit Is Already Written"));
  children.push(bodyPara(
    `Ana does not choose her gear. The Diplomat's Kit for a Fantasy table is a sealed letter, a silver ring, a warm cloak and a small mirror, and it is printed on her sheet before she touches it. Her six Pack slots are empty, because she has not been anywhere yet.`
  ));
  children.push(bodyPara(
    `She reads the four items and does what every good player does with a Kit: she asks who the letter is addressed to. Her GM does not know yet. Neither does she. That is now a thing the campaign owes both of them.`,
    { italics: true }
  ));

  children.push(pageBreak());
  children.push(sectionHeading("Step 8 — Fill In the Rest and Stop"));
  children.push(bodyPara(
    `Three Spotlight Tokens, zero Language Points, Homework Bonus box empty until she earns it. Growth Level 1, and a brand new Growth Ledger with nothing on it yet. Done — she is finished before half the table has decided on a name.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["MIRA'S SHEET", "AT SESSION ZERO"],
    anaSheet,
    [3400, 6680]
  ));

  children.push(spacer(240));
  children.push(sectionHeading("What Ana Did Not Do"));
  children.push(bodyPara(
    `She did not check which Focus the dice favour, because none of them do. She did not read all four Archetypes twice looking for the strongest, because they are not ranked. She did not write a backstory longer than a sentence, and she did not wait for permission to be finished. Two sessions later she was standing in front of a tired gatekeeper named Oren with her Empathy of +1 and a permit she could not produce — and that scene, in Chapter 2, is what all of this was for.`,
    { italics: true }
  ));

  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 8 — The Worlds You Can Play In
// ---------------------------------------------------------------------------
// Introductory only, by design. Each of these gets a full module in the
// Master's Guide (reskinned Focuses, world Archetypes, adventure skeletons).
// This chapter exists so a student can choose one at Session Zero.
// ---------------------------------------------------------------------------

const settingSummary = [
  ["Fantasy", "Courage · Empathy · Wit · Instinct", "Requests, negotiation, describing places and people."],
  ["Cosmic Horror", "Nerve · Rapport · Lore · Dread", "Hedging and uncertainty, reported speech, describing what you are not sure you saw."],
  ["Supernatural Investigation", "Grit · Rapport · Deduction · Hunch", "Question forms, past tenses, deduction language: must have, can't have, might have."],
  ["Dystopian Superheroes", "Valor · Charisma · Ingenuity · Reflex", "Opinions and argument, modals of obligation, persuading a crowd."],
  ["Post-Apocalypse Survival", "Steel · Trust · Salvage · Survival", "Giving instructions, conditionals, stating needs plainly and fast."],
  ["Cyberpunk", "Edge · Face · Hacking · Street", "Future forms, technical description, bargaining and double-talk."],
];

function settingBlock(name, pitch, body) {
  const parts = [];
  parts.push(sectionHeading(name));
  parts.push(bodyPara(pitch, { italics: true, after: 80 }));
  parts.push(bodyPara(body, { after: 180 }));
  return parts;
}

function chapterSettings() {
  const children = [];
  children.push(eyebrow("Chapter 9"));
  children.push(chapterTitle("The Worlds You Can Play In"));

  children.push(flavorQuote(
    `Six worlds. One engine. The dice never notice which one you picked.`
  ));
  children.push(spacer(140));

  children.push(bodyPara(
    `Everything you have read so far is the engine, and the engine does not care where the story happens. The same four Focuses, the same six Moves, the same three outcome bands run a haunted lighthouse, a collapsing megacity, and a road out of a burned town equally well. What changes between worlds is the dressing: what the Focuses are called, what an Archetype looks like when it walks into a room, and what kind of trouble the GM is allowed to make.`
  ));
  children.push(bodyPara(
    `Below are the six worlds ${GAME_NAME} is built to run. This chapter is a menu, not a manual — enough for your table to choose one at Session Zero and know roughly what it is agreeing to. The full version of each, with its own Archetypes, adventure skeletons, and world-specific Moves, lives in the Master's Guide, in your GM's hands.`
  ));

  settingBlock(
    "Fantasy",
    "Walled towns, tired gatekeepers, a road that goes somewhere worse than it looks.",
    `The default world, and the one every example in this book is drawn from. Swords, guilds, old magic that nobody fully understands, and problems that can usually be solved by talking to the right person before they have to be solved any other way. It is the most forgiving world to start in: the tone is flexible, the stakes scale easily up or down, and nothing about it requires content limits beyond common sense.`
  ).forEach(p => children.push(p));

  settingBlock(
    "Cosmic Horror",
    "Something is wrong with this town, and explaining what is the dangerous part.",
    `Slow dread rather than sudden violence. Characters investigate something they should probably leave alone, and the game is at its best when nobody is certain what they saw. This world runs on doubt, which makes it unusually good language practice — a table spends most of it hedging, qualifying, and reporting what someone else claimed. It is also the world most in need of a firm content agreement at Session Zero.`
  ).forEach(p => children.push(p));

  settingBlock(
    "Supernatural Investigation",
    "The case is real, the client is lying, and one of the witnesses is not alive.",
    `Ghosts, hauntings, and things that leave evidence. Structurally the most satisfying world for a short campaign, because a case has a shape: a question at the start and an answer at the end. Every session is built out of asking things, which makes it the strongest fit for a table working on question forms and past tenses.`
  ).forEach(p => children.push(p));

  settingBlock(
    "Dystopian Superheroes",
    "You have powers. So does the government, and theirs are legal.",
    `Big abilities, bigger consequences, and a world where the real problem is never the fight. Characters argue in public, take sides, and answer for what they did. The loudest of the six worlds, and the best one for a table that likes debating — most scenes end up being about whether an action was justified rather than whether it worked.`
  ).forEach(p => children.push(p));

  children.push(pageBreak());

  settingBlock(
    "Post-Apocalypse Survival",
    "The water is three days away and the truck holds four people.",
    `Scarcity, hard choices, and a group that has to keep functioning. Resources matter here more than in any other world, and so does telling people plainly what to do — this is the world where instructions, warnings, and conditionals get a real workout. It shares Cosmic Horror's need for an honest content conversation up front.`
  ).forEach(p => children.push(p));

  settingBlock(
    "Cyberpunk",
    "Everyone is being paid by someone, and nobody is being paid enough.",
    `Neon, contracts, and the assumption that every deal has a second layer. Characters take jobs, get betrayed, and negotiate their way sideways out of it. The most transactional of the six worlds, which makes it a natural fit for bargaining language, future forms, and technical description.`
  ).forEach(p => children.push(p));

  children.push(sectionHeading("The Six, Side by Side"));
  children.push(bodyPara(
    `Focus names change between worlds; the Focuses themselves do not. Courage is always the first column, Empathy the second, Wit the third, Instinct the fourth — whatever a given world decides to call them. The last column is the one your GM will care about most.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["WORLD", "THE FOUR FOCUSES BECOME", "THE ENGLISH IT DRILLS HARDEST"],
    settingSummary,
    [2200, 3400, 4480]
  ));

  children.push(spacer(200));
  children.push(sectionHeading("Choosing One, and Changing Your Mind"));
  children.push(bodyPara(
    `Your table picks a world at Session Zero, along with the campaign shape from Chapter 5 — Long Haul, Chronicle, or Anthology. Those two decisions go together: a table that already suspects it wants to try several of these worlds should say so now and run a Chronicle or an Anthology, rather than promising itself one long Fantasy campaign and quietly getting bored of it at session twelve.`
  ));
  children.push(bodyPara(
    `And if you change your mind anyway, nothing is lost. Your Growth Level is yours, your Ledger travels, and your Boons get reimagined for wherever you land next. Switching worlds costs you fiction, not progress.`
  ));

  return children;
}

// ---------------------------------------------------------------------------
// APPENDIX A — THE RULES IN PORTUGUESE
// The one deliberate exception to the English-only rule of this book. It exists
// so an A1 student can sit down and play in week one without understanding the
// chapters yet. Rule NAMES stay in English on purpose — those are what gets
// said out loud at the table; only the explanation is translated.
// ---------------------------------------------------------------------------

const ptRoll = [
  ["Você rola", "2d6 e soma o Focus que aquele Move pedir."],
  ["10 ou mais", "Strong Hit. Dá certo do jeito que você queria, sem custo."],
  ["7, 8 ou 9", "Mixed Result. Dá certo, mas tem um preço — o GM diz qual, ou te dá uma escolha difícil."],
  ["6 ou menos", "Miss. Não dá certo, e o GM responde com alguma coisa que complica a cena."],
];

const ptFocuses = [
  ["Courage", "Encarar o perigo de frente. Coragem física, aguentar o tranco."],
  ["Empathy", "Ler gente. Negociar, pedir, entender o que o outro realmente quer."],
  ["Wit", "Pensar rápido e com esperteza. Achar o ângulo, resolver o problema."],
  ["Instinct", "Perceber. Notar o detalhe, reagir antes de raciocinar."],
];

const ptMoves = [
  ["Face Danger", "Courage", "Quando você entra no perigo de propósito, para conseguir alguma coisa."],
  ["Parley", "Empathy", "Quando você faz um pedido direto, oferecendo algo que a outra pessoa quer."],
  ["Help or Interfere", "Empathy", "Quando você entra para ajudar — ou atrapalhar — a jogada de outro jogador."],
  ["Persuade or Manipulate", "Wit", "Quando você trabalha alguém: elogio, lógica, meia-verdade bem colocada."],
  ["Act Under Pressure", "Instinct", "Quando você tem que agir rápido, sem tempo de pensar."],
  ["Read the Scene", "Instinct", "Quando você para e observa uma pessoa, um lugar ou uma situação antes de agir."],
];

const ptSheet = [
  ["Language Focus", "O ponto de gramática que você está estudando esta semana. É só seu. Use bem numa cena e você ganha um Language Point."],
  ["Language Point", "Ganho quando você acerta o seu Language Focus em cena. Gasta para rolar 2d6 de novo."],
  ["Spotlight Token", "Compra espaço na cena para você. Zera a cada sessão — não acumula."],
  ["Kit", "Os 4 ou 5 itens que seu personagem sempre carrega. Vêm do Archetype. Não gasta, não conta."],
  ["Pack", "Seis espaços para o que você pegar pelo caminho. Cheio, tem que largar algo para pegar outra coisa."],
  ["Boon", "Recompensa de Growth Moment. Abre porta na história — nunca dá bônus no dado."],
  ["Growth Level", "Sobe 1 a cada 6 unidades do seu curso terminadas aqui. Todo mundo começa no 1, seja qual for o seu inglês."],
];

const ptTable = [
  ["Mistakes Are How We Play", "Ninguém corrige você no meio da cena. O GM devolve a forma certa dentro da fala dele."],
  ["One Scene, One Voice", "Deixe o colega terminar o momento dele antes de entrar."],
  ["Yes, And", "Construa em cima do que o outro trouxe. Não derrube a ideia dele."],
  ["In Character, In English", "Tudo o que você diz como personagem é em inglês. Travou? Descreva por volta, ou pergunte ao GM: how do I say ___?"],
];

function chapterAppendixPT() {
  const children = [];
  children.push(eyebrow("Appendix A"));
  children.push(chapterTitle("As Regras em Português"));

  children.push(flavorQuote([
    `This is the only page in the book that is not in English. Use it, and then outgrow it.`,
  ]));
  children.push(spacer(140));

  children.push(bodyPara(
    `Este resumo existe por um motivo só: para você conseguir jogar desde a primeira sessão, mesmo sem dar conta de ler o livro inteiro ainda. Ele cobre a mecânica principal e mais nada — o resto do livro tem coisas que este apêndice não substitui.`
  ));
  children.push(bodyPara(
    `Repare que os nomes das regras continuam em inglês. Isso é de propósito: são esses nomes que vão ser ditos em voz alta na mesa, toda sessão. Você aprende os nomes agora e a explicação depois.`
  ));
  children.push(calloutBox(
    "Uma meta, não um detalhe",
    `Quando você conseguir ler o livro inteiro em inglês, este apêndice para de ser necessário. Esse dia é uma das conquistas do curso — e vai chegar antes do que você imagina.`,
    "clarify"
  ));

  children.push(sectionHeading("Como funciona uma rolagem"));
  children.push(threeColTable(["QUANDO", "O QUE ACONTECE"], ptRoll, [2600, 7480]));

  children.push(spacer(200));
  children.push(sectionHeading("Os quatro Focuses"));
  children.push(bodyPara(
    `Na criação do personagem você distribui +2, +1, +0 e −1 entre os quatro. Todo mundo recebe exatamente esses quatro números — ninguém é mais forte que ninguém, você só escolhe onde é bom.`,
    { after: 100 }
  ));
  children.push(threeColTable(["FOCUS", "O QUE É"], ptFocuses, [2600, 7480]));

  children.push(pageBreak());
  children.push(sectionHeading("Os seis Moves"));
  children.push(bodyPara(
    `Todo Move usa um Focus. Você rola 2d6 e soma esse Focus.`,
    { after: 100 }
  ));
  children.push(threeColTable(["MOVE", "FOCUS", "QUANDO USAR"], ptMoves, [2900, 1600, 5580]));

  children.push(spacer(200));
  children.push(sectionHeading("O que tem na sua ficha"));
  children.push(threeColTable(["CAMPO", "O QUE É"], ptSheet, [2600, 7480]));

  children.push(spacer(200));
  children.push(sectionHeading("Dinheiro e distância"));
  children.push(bodyPara(
    `Dinheiro tem quatro degraus: a coin, a handful (10 coins), a bag (10 handfuls) e a chest (10 bags, e é o máximo que dá para carregar). Preço se fala em degraus — uma espada custa two handfuls. Ninguém faz conta.`
  ));
  children.push(bodyPara(
    `Distância também tem quatro degraus, e o jogo não usa metros: Within reach (dá para tocar), Nearby (mesma sala), Far away (precisa se deslocar) e Out of sight (fora de vista, não dá para agir).`
  ));

  children.push(pageBreak());
  children.push(sectionHeading("As quatro regras da mesa"));
  children.push(threeColTable(["REGRA", "O QUE SIGNIFICA"], ptTable, [3000, 7080]));

  children.push(spacer(220));
  children.push(calloutBox(
    "O resto do livro",
    `Tudo o que não está aqui — Archetypes, a trilha de doze níveis, os settings, o que fazer quando alguém entra na campanha no meio — está em inglês, nos capítulos. Peça ajuda ao seu GM sempre que precisar. Perguntar faz parte do jogo.`,
    "example"
  ));

  return children;
}

// ---------------------------------------------------------------------------
// ASSEMBLE THE WHOLE BOOK
// ---------------------------------------------------------------------------

const children = [];
children.push(...titlePage());
children.push(...chapter1());
children.push(pageBreak());
children.push(...chapter2());
children.push(pageBreak());
children.push(...chapter3());
children.push(pageBreak());
children.push(...chapter4());
children.push(pageBreak());
children.push(...chapterArchetypes());
children.push(pageBreak());
children.push(...chapterGear());
children.push(pageBreak());
children.push(...chapterAna());
children.push(pageBreak());
children.push(...chapterResources());
children.push(pageBreak());
children.push(...chapterSettings());
children.push(pageBreak());
children.push(...chapterEtiquette());
children.push(pageBreak());
children.push(...chapterResponsibilities());
children.push(pageBreak());
children.push(...chapterSessionZero());
children.push(pageBreak());
children.push(...chapterAppendixPT());

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      },
    },
    footers: { default: pageFooter() },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("/home/claude/rpl_core_rulebook/CoreRulebook.docx", buf);
  console.log("written");
});
