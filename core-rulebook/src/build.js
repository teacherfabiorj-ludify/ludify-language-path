// ============================================================================
// RPL — CORE RULEBOOK — single-manuscript build script
//
// This file is the ONE source of truth for the whole Core Rulebook.
// Every chapter lives here, as data + layout, in the order it appears in
// the book (CH1, CH2, ... CH8) — even chapters not written yet get a
// placeholder entry in the roadmap table so the manuscript always shows
// the full planned shape of the book.
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
  ImageRun,
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

const GAME_NAME = "RPL"; // working title — not locked yet, see blueprint §11

// ---- Palette (shared across Core Rulebook AND Master's Guide for visual consistency) ----
const ACCENT = "2A78D6";
const GOOD = "0CA30C";
const WARN = "FAB219";
const CRIT = "D03B3B";
const INK = "0B0B0B";
const INK_SECONDARY = "52514E";
const MUTED = "898781";
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

function bodyPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 160 },
    children: [ new TextRun({ text, color: INK, size: 22, italics: opts.italics || false }) ],
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
              children: [ new TextRun({ text, italics: true, color: INK_SECONDARY, size: 22 }) ],
            })),
          }),
        ],
      }),
    ],
  });
}

function calloutBox(label, text, color = GOOD) {
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
              new Paragraph({ children: [ new TextRun({ text, size: 20, color: INK }) ] }),
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
    children: r.map((cell, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: idx % 2 ? ZEBRA : WHITE },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [ new Paragraph({ children: [ new TextRun({ text: cell, size: 19, color: i === 0 ? INK : INK_SECONDARY, bold: i === 0 }) ] }) ],
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

// ---------------------------------------------------------------------------
// COVER NOTE — draft status banner (not a real book page, just orientation
// for whoever opens the file while it's still being written)
// ---------------------------------------------------------------------------

function coverNote() {
  const children = [];
  children.push(eyebrow(`${GAME_NAME} — Core Rulebook`));
  children.push(new Paragraph({
    spacing: { after: 160 },
    children: [ new TextRun({ text: "Draft manuscript", bold: true, color: INK, size: 40 }) ],
  }));
  children.push(bodyPara(
    "This is a living document — chapters are added and revised in place as they're written, not necessarily in final reading order until the book is complete. Check the roadmap table at the end of Chapter 1 for what's written so far.",
    { italics: true }
  ));
  children.push(spacer(200));
  children.push(pageBreak());
  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 1 — Welcome to [Game Name]
// ---------------------------------------------------------------------------

const bookRoadmap = [
  ["Ch. 2 — How a Turn Works", "The 2d6 engine, the three outcome bands, one worked example.", "Written"],
  ["Ch. 3 — Your Character", "Focuses, your Language Focus, reading your own sheet.", "Written"],
  ["Ch. 4 — Moves", "The six core Moves everyone shares, no matter the world.", "Written"],
  ["Ch. 5 — Language Points & Spotlight Tokens", "How you earn them, how you spend them.", "Written"],
  ["Ch. 6 — Table Etiquette", "The house rules that keep this a safe place to make mistakes.", "Not yet written"],
  ["Ch. 7 — Your Responsibilities", "Homework, punctuality, showing up ready.", "Not yet written"],
  ["Ch. 8 — Session Zero Checklist", "What to align before your very first adventure.", "Not yet written"],
];

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

  children.push(sectionHeading("What You'll Find in This Book"));
  children.push(bodyPara(
    `Here's the shape of the whole Core Rulebook. Chapters marked “Written” are finished; the rest are on the way.`,
    { after: 120 }
  ));
  children.push(threeColTable(
    ["CHAPTER", "WHAT'S INSIDE", "STATUS"],
    bookRoadmap,
    [3400, 4680, 2000]
  ));

  children.push(spacer(180));
  children.push(sectionHeading("Before You Play"));
  children.push(bodyPara(
    `This game runs on trying, not on being right — more on exactly what that means in Chapter 6. Your actual starting point isn't Chapter 2, though: it's the Session Zero Checklist at the very end of this book (Chapter 8), where your table aligns on tone, expectations, and how everyone wants to play before the first adventure begins.`
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
    ACCENT
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
    `This is also what keeps every Archetype fair against every other, no matter which world you're playing in — a Warrior in the Fantasy setting and a Netrunner in the Cyberpunk setting both build from the same four numbers. Your Setting Guide will suggest an array that fits each Archetype's concept, but the choice is always yours.`,
    ACCENT
  ));

  children.push(sectionHeading("Your Language Focus"));
  children.push(bodyPara(
    `This is the one field on your sheet that's yours and nobody else's. Your Language Focus is the grammar structure or vocabulary set you're actively working on this week — assigned by your GM, pulled from whatever you're covering in your regular classes, and updated regularly as you move forward. When a Move you're making happens to line up with it, your GM may ask you specifically to narrate using that structure. Nail it, and that's a Language Point (Chapter 5).`
  ));
  children.push(bodyPara(
    `Your Language Focus card is also the one place in this entire book where you'll find Portuguese. Every rule, every Move, every line the GM or an NPC says stays in English — that's the whole point. But the grammar point itself, right there on your card, gets one line of support in your own language, so you always know exactly what you're aiming for. The only other place that happens is the short lore recap your GM gives you at the top of a session — never in the scene itself.`
  ));
  children.push(calloutBox(
    "Example",
    `Language Focus: real conditionals, present/future. Card note: “if / when clauses — se você fizer algo, algo vai acontecer (condição).” Everything else about the scene — what the GM says, what your character says back — stays in English.`,
    GOOD
  ));

  children.push(sectionHeading("Reading Your Sheet"));
  children.push(bodyPara(
    `Here's every section of your character sheet, and what lives in each one.`,
    { after: 100 }
  ));
  children.push(threeColTable(
    ["SECTION", "WHAT'S THERE", "CHANGES"],
    [
      ["Identity", "Your name, your Setting, your Archetype.", "Set once, at creation."],
      ["Focuses", "Courage, Empathy, Wit, Instinct — each from −1 to +2.", "Set once, at creation."],
      ["Moves", "The six base Moves everyone has, plus 1–2 Signature Moves from your Archetype.", "Signature Moves unlock at creation; the six base Moves never change."],
      ["Language Focus", "This week's grammar or vocabulary target, with its one line of Portuguese support.", "Updated as you progress through your course."],
      ["Resources", "Spotlight Tokens and Language Points (Chapter 5).", "Refill at the start of every session."],
      ["Consequence", "A simple track for setbacks your character picks up during play.", "Changes scene to scene."],
      ["Progress", "Your Homework Bonus checkbox for the week.", "Reset every session."],
      ["Bio", "Two or three lines about who your character is.", "Whenever you want to add to it."],
    ],
    [2200, 5480, 2400]
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
    `Every time your character does something risky, clever, or uncertain, you're making a Move. Moves are the bridge between the story and the dice — they turn “I try to convince the guard” into a real roll with real stakes. There are six Moves everyone shares, no matter what world you're playing in. Your Archetype may add one or two Signature Moves of its own — you'll find those in your Setting Guide.`
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
    "Reminder: any Move is also a chance to use your current Language Focus. If your Move lines up with what you're working on this week, your GM may invite you to narrate it using that structure — and if you nail it, that's a Language Point. See Chapter 5.",
    { italics: true }
  ));
  return children;
}

// ---------------------------------------------------------------------------
// CHAPTER 5 — Language Points & Spotlight Tokens
// ---------------------------------------------------------------------------

function chapter5() {
  const children = [];
  children.push(eyebrow("Chapter 5"));
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
    ACCENT
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
    GOOD
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
// ASSEMBLE THE WHOLE BOOK
// ---------------------------------------------------------------------------

const children = [];
children.push(...coverNote());
children.push(...chapter1());
children.push(pageBreak());
children.push(...chapter2());
children.push(pageBreak());
children.push(...chapter3());
children.push(pageBreak());
children.push(...chapter4());
children.push(pageBreak());
children.push(...chapter5());

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("/home/claude/rpl_core_rulebook/CoreRulebook.docx", buf);
  console.log("written");
});
