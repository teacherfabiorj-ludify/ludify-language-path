# RPL — Language Path (protótipo)

Protótipo clicável do conceito **Language Path**: painel onde o aluno acompanha a trilha de units do seu nível (baseado no scope-and-sequence do Evolve/Cambridge), marca os tópicos que já estudou, e tenta um quiz de desbloqueio — 70%+ libera a próxima unit.

Faz parte do design do **Projeto RPL (Role Playing Language)**, o sistema de RPG de mesa autoral da escola de inglês Ludify, criado por Fábio Ferreira Gomes.

## O que este protótipo mostra

- Trilha visual das 12 units do nível B1 (dados reais do Evolve), com status: completa / atual / bloqueada.
- Painel da unit atual, com checklist de self-check dos tópicos gramaticais/vocabulário.
- Quiz de desbloqueio funcional (5 perguntas reais da Unit 6 — "Impact": quantifiers e real conditionals).
- Ao atingir 70%+ no quiz, a próxima unit desbloqueia ao vivo na trilha.

## O que este protótipo NÃO é

- Não salva nada — é um demo local, os dados resetam a cada recarregamento da página.
- Não tem backend nem login — é só front-end (HTML/CSS/JS puro, sem dependências).
- Não é a ferramenta que os alunos vão usar de fato — é uma referência visual pra validar a ideia antes de decidir o caminho de implementação real (Google Forms+Sheets, Airtable, ou uma versão evoluída deste protótipo).

## Como rodar

Não precisa de instalação nem servidor. Basta abrir `index.html` num navegador — ou, se este repositório estiver publicado via GitHub Pages, acessar o link público direto.

## Stack

HTML, CSS e JavaScript puros, em um único arquivo (`index.html`). Sem frameworks, sem build step.

---

*Documento de design completo do Projeto RPL vive fora deste repositório, no workspace de projeto do Fábio no Claude.*
