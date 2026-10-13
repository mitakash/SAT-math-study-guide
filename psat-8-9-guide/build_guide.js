/*
 * build_guide.js
 * ----------------------------------------------------------------------------
 * PURPOSE
 *   Turns the content in content.js into the Word file PSAT_8-9_Math_Guide.docx.
 *   This file is ONLY layout. To change math, edit content.js, then re-run
 *   verify_answers.py, then re-run this file.
 *
 * RUN
 *   node build_guide.js            (writes PSAT_8-9_Math_Guide.docx next to it)
 *
 * DOCUMENT ORDER (top to bottom)
 *   1. Formula sheet (pages 1-2), grouped by the 4 test domains, Top 20 tagged
 *   2. "How this guide works" page: scoring, Top 20 checklist
 *   3. 30 topic blocks, grouped by domain. Each block:
 *        header bar -> The idea -> (formulas) -> Example -> Q1 Normal -> Q2 Boss
 *        -> Watch out
 *      After each domain: Spot the Mistake + Joke Decoder
 *   4. Answer key (two columns): step-by-step solutions, mistake fixes, punchlines
 *
 * LIBRARY: "docx" npm package (docx-js). Units: DXA, 1440 DXA = 1 inch.
 *          Font sizes are in half-points (20 = 10pt).
 * ----------------------------------------------------------------------------
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  BorderStyle, ShadingType, AlignmentType, HeadingLevel, TabStopType, PageBreak,
  Footer, PageNumber, SectionType, VerticalAlign, LevelFormat,
} = require('docx');
const { DOMAINS, FORMULA_SHEET, NOT_ON_TEST, TOPICS, MISTAKES, JOKES } = require('./content.js');

// ---------------------------------------------------------------------------
// Page geometry and typography constants
// ---------------------------------------------------------------------------
const PAGE_W = 12240;            // US Letter width (8.5 in)
const PAGE_H = 15840;            // US Letter height (11 in)
const MARGIN = 900;              // 0.625 in margins: roomy enough to write notes
const CONTENT_W = PAGE_W - 2 * MARGIN;   // 10440 DXA usable width
const FONT = 'Cambria';          // textbook serif; LibreOffice substitutes Caladea
const SIZE = 20;                 // 10pt body text
const INK = '222222';            // near-black body text
const GREY = '666666';           // secondary text
const RULE = 'BBBBBB';           // light rule lines

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };

// ---------------------------------------------------------------------------
// rich(text, base): parse the mini markup from content.js into TextRuns.
//   ^{..} superscript, _{..} subscript, **..** bold. "base" holds default run
//   options (size, color, bold, italics) applied to every piece.
// ---------------------------------------------------------------------------
function rich(text, base = {}) {
  const runs = [];
  const parts = String(text).split(/(\^\{[^}]*\}|_\{[^}]*\}|\*\*[^*]+\*\*)/);
  for (const part of parts) {
    if (!part) continue;
    if (part.startsWith('^{')) runs.push(new TextRun({ ...base, text: part.slice(2, -1), superScript: true }));
    else if (part.startsWith('_{')) runs.push(new TextRun({ ...base, text: part.slice(2, -1), subScript: true }));
    else if (part.startsWith('**')) runs.push(new TextRun({ ...base, text: part.slice(2, -2), bold: true }));
    else runs.push(new TextRun({ ...base, text: part }));
  }
  return runs;
}

// Visible length of a markup string (used to choose how to lay out choices).
const visibleLen = (s) => String(s).replace(/\^\{|_\{|\}|\*\*/g, '').length;

// Two-digit topic number label, e.g. 7 -> "07".
const pad = (n) => String(n).padStart(2, '0');

// Shortcut for a plain paragraph. opts: indent, before, after, keepNext, align
function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : rich(children, opts.run || {}),
    spacing: { before: opts.before ?? 0, after: opts.after ?? 60, line: opts.line ?? 264 },
    indent: opts.indent ? { left: opts.indent, hanging: opts.hanging || 0 } : undefined,
    keepNext: opts.keepNext ?? false,
    keepLines: true,
    alignment: opts.align,
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: opts.shade, color: 'auto' } : undefined,
    border: opts.border,
    tabStops: opts.tabStops,
  });
}

// A single table cell with explicit width (needed for Word AND Google Docs).
function cell(children, width, opts = {}) {
  return new TableCell({
    children,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    margins: { top: opts.padV ?? 40, bottom: opts.padV ?? 40, left: opts.padH ?? 80, right: opts.padH ?? 80 },
    borders: opts.borders,
    verticalAlign: opts.vAlign || VerticalAlign.CENTER,
    columnSpan: opts.span,
  });
}

// ===========================================================================
// PART 1: FORMULA SHEET (pages 1-2)
// ===========================================================================
function buildFormulaSheet() {
  const out = [];
  // Title, with a rule under it
  out.push(new Paragraph({
    children: [
      new TextRun({ text: 'PSAT 8/9 Math', bold: true, size: 40, color: '1F3A68', font: FONT }),
      new TextRun({ text: '   The Essential Formula Sheet', size: 32, color: '1F3A68', font: FONT }),
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: '1F3A68', space: 4 } },
    spacing: { after: 100 },
  }));
  // One-line test facts so students know the format
  out.push(para([
    new TextRun({ text: 'The test: ', bold: true, size: 18, font: FONT }),
    new TextRun({ text: '44 math questions in 70 minutes (two 35-minute modules). About 3 out of 4 are multiple choice; the rest are type-in. A calculator (Desmos is built in) is allowed for every question. A reference sheet with area and volume formulas is provided.', size: 18, font: FONT, color: INK }),
  ], { after: 60 }));
  // Legend explaining the TOP 20 tag and topic numbers
  out.push(para([
    new TextRun({ text: ' TOP 20 ', bold: true, size: 16, color: 'FFFFFF', shading: { type: ShadingType.CLEAR, fill: '333333', color: 'auto' }, font: FONT }),
    new TextRun({ text: '  = shows up on almost every test. Learn these first.      ', size: 18, font: FONT, color: INK }),
    new TextRun({ text: 'Topic #', bold: true, size: 18, font: FONT, color: GREY }),
    new TextRun({ text: ' = where to find the lesson and practice questions.', size: 18, font: FONT, color: INK }),
  ], { after: 120 }));

  const TAG_W = 900, NUM_W = 760, TEXT_W = CONTENT_W - TAG_W - NUM_W;
  for (const key of ['A', 'M', 'D', 'G']) {
    const d = DOMAINS[key];
    // Domain header bar: colored rule + name + share of test
    out.push(new Paragraph({
      children: [
        new TextRun({ text: d.name.toUpperCase(), bold: true, size: 22, color: 'FFFFFF', font: FONT }),
        new TextRun({ text: '\t' + d.share, size: 18, color: 'FFFFFF', font: FONT }),
      ],
      shading: { type: ShadingType.CLEAR, fill: d.dark, color: 'auto' },
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W - 120 }],
      spacing: { before: 140, after: 0 },
      indent: { left: 80, right: 80 },
      keepNext: true,
    }));
    if (key === 'G') {
      out.push(para('Most geometry formulas are printed on the test reference sheet. You do not need to memorize them; you need to know WHEN to use them.', { run: { size: 17, italics: true, color: GREY, font: FONT }, before: 40, after: 20, keepNext: true }));
    }
    // One row per formula line: [TOP 20 tag] [formula text] [topic #]
    const rows = FORMULA_SHEET[key].map((f, i) => {
      const line = { style: BorderStyle.SINGLE, size: 2, color: 'DDDDDD' };
      const b = { top: NONE, left: NONE, right: NONE, bottom: line };
      const tag = f.top
        ? [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'TOP 20', bold: true, size: 14, color: 'FFFFFF', font: FONT })] })]
        : [new Paragraph({ children: [] })];
      return new TableRow({
        cantSplit: true,
        children: [
          cell(tag, TAG_W, { fill: f.top ? '333333' : undefined, borders: b, padV: 30 }),
          cell([para(f.t, { run: { size: 19, font: FONT, color: INK }, after: 0 })], TEXT_W, { borders: b, fill: i % 2 ? 'FAFAFA' : undefined, padV: 50 }),
          cell([para('# ' + f.n, { run: { size: 16, color: GREY, font: FONT }, after: 0, align: AlignmentType.RIGHT })], NUM_W, { borders: b, padV: 30 }),
        ],
      });
    });
    out.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [TAG_W, TEXT_W, NUM_W], borders: NO_BORDERS, rows }));
  }
  // Closing notes: what NOT to study, and the Desmos tip
  out.push(para([
    new TextRun({ text: 'Skip: ', bold: true, size: 18, font: FONT }),
    new TextRun({ text: NOT_ON_TEST, size: 18, font: FONT, color: INK }),
  ], { before: 160, after: 40 }));
  out.push(para([
    new TextRun({ text: 'Desmos habit: ', bold: true, size: 18, font: FONT }),
    new TextRun({ text: 'for systems, quadratics, and "where is the max or min?" questions, graphing is often faster than algebra. Practice both ways.', size: 18, font: FONT, color: INK }),
  ], { after: 0 }));
  // No page break here: the "How This Guide Works" section fills the rest of page 2.
  return out;
}

// ===========================================================================
// PART 2: HOW THIS GUIDE WORKS + TOP 20 CHECKLIST
// ===========================================================================
function buildIntro() {
  const out = [];
  out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 80 }, keepNext: true, children: [new TextRun({ text: 'How This Guide Works', font: FONT })] }));
  const bullets = [
    ['Each topic has: ', 'the idea in one or two sentences, one worked example, then two questions. Question 1 is Normal level. Question 2 is Boss level (a little harder, more like a real test question).'],
    ['Question types match the test. ', 'Multiple choice has four answers, A to D. "Type-in" questions have no choices: you work out the number yourself. On the real test you can type fractions like 5/11 and negative numbers like −6.'],
    ['Answers are at the back, ', 'with every step written out. Try the question first. If you get it wrong, read the steps and find exactly which step went differently.'],
    ['After each section: ', 'a Spot the Mistake problem (find where a student went wrong) and a Joke Decoder (your answers spell out the punchline, but only if they are right).'],
  ];
  for (const [b, t] of bullets) {
    out.push(new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text: b, bold: true, font: FONT, size: SIZE }), new TextRun({ text: t, font: FONT, size: SIZE, color: INK })],
      spacing: { after: 40 },
    }));
  }

  // Scoreboard: one compact line (a full table pushed the checklist onto a third page).
  out.push(para([
    new TextRun({ text: 'Scoreboard.  ', bold: true, size: SIZE, font: FONT, color: '1F3A68' }),
    new TextRun({ text: 'Normal question 1 pt (x30).  Boss question 2 pts (x30).  Spot the Mistake 2 pts (x4).  Joke Decoder 3 pts (x4).  Total possible: 110.   My score: ______', size: SIZE, font: FONT, color: INK }),
  ], { before: 60, after: 0 }));

  // Top 20 checklist in two columns
  out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 160, after: 60 }, children: [new TextRun({ text: 'The Top 20 Checklist', font: FONT })] }));
  out.push(para('These 20 topics cover the skills that show up most on the PSAT 8/9. Check a box when you can do BOTH practice questions without looking at the answers.', { run: { size: SIZE, font: FONT, color: INK }, after: 100 }));
  const top = TOPICS.filter((t) => t.top20);
  const half = Math.ceil(top.length / 2);
  const colW = CONTENT_W / 2;
  const cRows = [];
  for (let i = 0; i < half; i++) {
    const mk = (t) => t
      ? cell([para([
          new TextRun({ text: '☐  ', size: 22, font: 'Segoe UI Symbol' }),
          new TextRun({ text: pad(t.num) + '  ', bold: true, size: 19, color: DOMAINS[t.domain].dark, font: FONT }),
          new TextRun({ text: t.title, size: 19, font: FONT }),
        ], { after: 0 })], colW, { borders: { top: NONE, left: NONE, right: NONE, bottom: { style: BorderStyle.SINGLE, size: 2, color: 'DDDDDD' } }, padV: 30 })
      : cell([new Paragraph({ children: [] })], colW, { borders: { top: NONE, left: NONE, right: NONE, bottom: NONE } });
    cRows.push(new TableRow({ children: [mk(top[i]), mk(top[i + half])] }));
  }
  out.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [colW, colW], borders: NO_BORDERS, rows: cRows }));
  out.push(para('Topics without the Top 20 tag still show up, just less often. Do them after the Top 20.', { run: { size: 18, italics: true, color: GREY, font: FONT }, before: 120 }));
  out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}

// ===========================================================================
// PART 3: TOPIC BLOCKS
// ===========================================================================

// Multiple-choice answers laid out like the test: short answers in one row,
// medium answers in a 2 x 2 grid, long answers one per line.
function choicesBlock(choices) {
  const L = ['A', 'B', 'C', 'D'];
  const maxLen = Math.max(...choices.map(visibleLen));
  const perRow = maxLen <= 14 ? 4 : maxLen <= 38 ? 2 : 1;
  const indent = 360;
  const colW = Math.floor((CONTENT_W - indent) / perRow);
  const widths = Array(perRow).fill(colW);
  const rows = [];
  for (let i = 0; i < 4; i += perRow) {
    rows.push(new TableRow({
      cantSplit: true,
      children: widths.map((w, j) => cell([para([
        new TextRun({ text: L[i + j] + ')  ', bold: true, size: SIZE, font: FONT }),
        ...rich(choices[i + j], { size: SIZE, font: FONT, color: INK }),
      ], { after: 0, keepNext: true })], w, { padV: 20, padH: 60 })),
    }));
  }
  return new Table({
    width: { size: colW * perRow, type: WidthType.DXA },
    columnWidths: widths,
    borders: NO_BORDERS,
    indent: { size: indent, type: WidthType.DXA },
    rows,
  });
}

// A small data table (used by topic 23's two-way table).
function dataTable(tbl, d) {
  const n = tbl.header.length;
  const first = 1800, rest = 1100;
  const widths = [first, ...Array(n - 1).fill(rest)];
  const border = { style: BorderStyle.SINGLE, size: 4, color: RULE };
  const b = { top: border, bottom: border, left: border, right: border };
  const mkRow = (vals, header) => new TableRow({
    cantSplit: true,
    children: vals.map((v, i) => cell([para(v, { run: { size: 18, font: FONT, bold: header || i === 0 }, after: 0, align: i ? AlignmentType.CENTER : AlignmentType.LEFT })], widths[i], { borders: b, fill: header ? d.tint : undefined, padV: 20 })),
  });
  return [
    para(tbl.caption, { run: { size: 18, italics: true, color: GREY, font: FONT }, indent: 360, after: 30, keepNext: true }),
    new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths, indent: { size: 360, type: WidthType.DXA }, rows: [mkRow(tbl.header, true), ...tbl.rows.map((r) => mkRow(r, false))] }),
    para('', { after: 40, keepNext: true }),
  ];
}

// Label run for "The idea", "Example", etc.
const label = (text, color) => new TextRun({ text: text + '  ', bold: true, size: SIZE, color, font: FONT });

function topicBlock(t) {
  const d = DOMAINS[t.domain];
  const out = [];
  const K = { keepNext: true };                 // keep every line with the next one
  const body = { size: SIZE, font: FONT, color: INK };

  // Header bar: big number, title, and tags on the right
  const tagText = (t.top20 ? 'TOP 20   ' : '') + (t.bonus ? 'BONUS   ' : '') + d.name.toUpperCase();
  out.push(new Paragraph({
    children: [
      new TextRun({ text: pad(t.num) + '  ', bold: true, size: 30, color: d.dark, font: FONT }),
      new TextRun({ text: t.title, bold: true, size: 25, color: d.dark, font: FONT }),
      new TextRun({ text: '\t' + tagText, bold: true, size: 15, color: d.dark, font: FONT }),
    ],
    shading: { type: ShadingType.CLEAR, fill: d.tint, color: 'auto' },
    border: { left: { style: BorderStyle.SINGLE, size: 36, color: d.dark, space: 6 } },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W - 140 }],
    spacing: { before: 220, after: 80 },
    keepNext: true, keepLines: true,
  }));

  // The idea
  out.push(para([label('The idea.', d.dark), ...rich(t.idea, body)], K));

  // Key formulas, if the topic has them, in a light shaded strip
  for (const f of t.formulas || []) {
    out.push(para(f, { run: { size: SIZE, font: FONT, color: INK }, shade: 'F2F2F2', indent: 360, after: 30, keepNext: true }));
  }
  if (t.formulas) out.push(para('', { after: 0, line: 120, keepNext: true }));

  // Two-way table, if any (topic 23)
  if (t.table) out.push(...dataTable(t.table, d));

  // Worked example: problem, then numbered steps
  out.push(para([label('Example.', d.dark), ...rich(t.example.problem, { ...body, italics: true })], K));
  t.example.steps.forEach((s) => out.push(para(s, { run: body, indent: 360, after: 30, keepNext: true })));

  // The two practice questions
  for (const [qk, name, pts] of [['q1', 'Question 1  (Normal, 1 pt)', 1], ['q2', 'Question 2  (Boss level, 2 pts)', 2]]) {
    const q = t[qk];
    out.push(para([label(name, d.dark), ...(q.choices ? [] : [new TextRun({ text: '[type-in]', size: 16, color: GREY, font: FONT })])], { before: 80, after: 20, keepNext: true }));
    out.push(para(q.stem, { run: body, indent: 360, after: 40, keepNext: true }));
    if (q.choices) out.push(choicesBlock(q.choices));
    else out.push(para('Your answer: ______________', { run: { size: SIZE, font: FONT, color: GREY }, indent: 360, after: 20, keepNext: true }));
  }

  // Calculator tip, if any
  if (t.desmos) out.push(para([label('Desmos.', d.dark), ...rich(t.desmos, body)], { before: 60, after: 30, keepNext: true }));

  // Watch out: the most common trap. Last line of the block (no keepNext).
  out.push(para([label('Watch out.', '9A3412'), ...rich(t.watch, body)], {
    before: 80, after: 60,
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: '9A3412', space: 6 } },
    indent: 120,
  }));
  return out;
}

// Spot the Mistake box shown after each domain's topics.
function mistakeBlock(key) {
  const m = MISTAKES[key];
  const d = DOMAINS[key];
  const out = [];
  out.push(new Paragraph({
    children: [new TextRun({ text: 'Spot the Mistake', bold: true, size: 25, color: d.dark, font: FONT }), new TextRun({ text: '\t2 pts', bold: true, size: 15, color: d.dark, font: FONT })],
    border: { top: { style: BorderStyle.SINGLE, size: 8, color: d.dark, space: 4 }, bottom: { style: BorderStyle.SINGLE, size: 8, color: d.dark, space: 4 } },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W - 140 }],
    spacing: { before: 300, after: 80 }, keepNext: true,
  }));
  out.push(para([new TextRun({ text: m.who + ' was asked: ', bold: true, size: SIZE, font: FONT }), ...rich(m.task, { size: SIZE, font: FONT, color: INK })], { keepNext: true }));
  out.push(para(m.who + '\'s work:', { run: { size: SIZE, font: FONT, italics: true, color: GREY }, after: 30, keepNext: true }));
  m.work.forEach((w, i) => out.push(para(w, { run: { size: 22, font: FONT, color: INK }, indent: 720, shade: 'F7F3E8', after: 0, keepNext: true })));
  out.push(para('Which line is wrong? What is the correct answer? Prove it by checking with a number.', { run: { size: SIZE, font: FONT, color: INK }, before: 80, after: 20, keepNext: true }));
  out.push(para('______________________________________________________________________________', { run: { color: 'BBBBBB', size: SIZE }, before: 160, after: 0 }));
  return out;
}

// Joke Decoder shown after each domain's topics.
function jokeBlock(key) {
  const j = JOKES[key];
  const d = DOMAINS[key];
  const out = [];
  out.push(new Paragraph({
    children: [new TextRun({ text: 'Joke Decoder', bold: true, size: 25, color: d.dark, font: FONT }), new TextRun({ text: '\t3 pts', bold: true, size: 15, color: d.dark, font: FONT })],
    border: { top: { style: BorderStyle.SINGLE, size: 8, color: d.dark, space: 4 }, bottom: { style: BorderStyle.SINGLE, size: 8, color: d.dark, space: 4 } },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W - 140 }],
    spacing: { before: 300, after: 80 }, keepNext: true,
  }));
  out.push(para([new TextRun({ text: j.setup, bold: true, size: 22, font: FONT })], { keepNext: true }));
  out.push(para('Write your answer to each Boss question below. Find that number in the key and write its word. If the sentence makes no sense, one of your answers is wrong.', { run: { size: SIZE, font: FONT, color: INK }, after: 80, keepNext: true }));

  // Left table: the slots the student fills in. Right table: the decoder key.
  const border = { style: BorderStyle.SINGLE, size: 4, color: RULE };
  const b = { top: border, bottom: border, left: border, right: border };
  const sW = [1700, 1500, 2100];
  const slotRows = [
    new TableRow({ children: ['Question', 'Your answer', 'Word'].map((h, i) => cell([para(h, { run: { bold: true, size: 17, font: FONT }, after: 0 })], sW[i], { fill: d.tint, borders: b })) }),
    ...j.slots.map((s) => new TableRow({ cantSplit: true, children: [
      cell([para('Topic ' + s.ref + ', Q2', { run: { size: 18, font: FONT }, after: 0 })], sW[0], { borders: b, padV: 70 }),
      cell([para('', { after: 0 })], sW[1], { borders: b }),
      cell([para('', { after: 0 })], sW[2], { borders: b }),
    ] })),
  ];
  // Key: correct words mixed with decoys, sorted by number so order gives nothing away
  const toNum = (v) => Number(String(v).replace('−', '-'));
  const key2 = [...j.slots.map((s) => [s.value, s.word]), ...j.decoys].sort((a, c) => toNum(a[0]) - toNum(c[0]));
  const kW = [1000, 1700];
  const half = Math.ceil(key2.length / 2);
  const keyRows = [];
  for (let i = 0; i < half; i++) {
    const pair = [key2[i], key2[i + half]];
    keyRows.push(new TableRow({ cantSplit: true, children: pair.flatMap((p) => p
      ? [cell([para(p[0], { run: { bold: true, size: 18, font: FONT }, after: 0, align: AlignmentType.RIGHT })], kW[0], { borders: b, fill: 'F2F2F2' }),
         cell([para(p[1], { run: { size: 17, font: FONT }, after: 0 })], kW[1], { borders: b })]
      : [cell([para('', { after: 0 })], kW[0], { borders: b }), cell([para('', { after: 0 })], kW[1], { borders: b })]) }));
  }
  const gap = 240;
  const leftW = sW.reduce((a, c) => a + c, 0);
  const rightW = 2 * (kW[0] + kW[1]);
  const outerW = [leftW, gap, rightW];
  out.push(new Table({
    width: { size: outerW.reduce((a, c) => a + c, 0), type: WidthType.DXA },
    columnWidths: outerW,
    borders: NO_BORDERS,
    rows: [new TableRow({ cantSplit: true, children: [
      cell([new Table({ width: { size: leftW, type: WidthType.DXA }, columnWidths: sW, rows: slotRows }), para('', { after: 0 })], leftW, { padH: 0, vAlign: VerticalAlign.TOP }),
      cell([para('', { after: 0 })], gap, { padH: 0 }),
      cell([para('Decoder key', { run: { bold: true, size: 17, font: FONT, color: GREY }, after: 30 }), new Table({ width: { size: rightW, type: WidthType.DXA }, columnWidths: [...kW, ...kW], rows: keyRows }), para('', { after: 0 })], rightW, { padH: 0, vAlign: VerticalAlign.TOP }),
    ] })],
  }));
  out.push(para([new TextRun({ text: 'Punchline:  ', bold: true, size: SIZE, font: FONT }), new TextRun({ text: j.slots.map(() => '____________').join('  '), color: '999999', size: SIZE, font: FONT })], { before: 140, after: 60 }));
  return out;
}

function buildTopics() {
  const out = [];
  for (const key of ['A', 'M', 'D', 'G']) {
    const d = DOMAINS[key];
    // Domain heading (Heading 1, so it shows in Word's navigation pane)
    out.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: key !== 'A',
      children: [new TextRun({ text: d.name, font: FONT, color: d.dark }), new TextRun({ text: '   ' + d.share, size: 20, color: GREY, font: FONT, bold: false })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: d.dark, space: 4 } },
      keepNext: true,
    }));
    for (const t of TOPICS.filter((x) => x.domain === key)) out.push(...topicBlock(t));
    out.push(...mistakeBlock(key));
    out.push(...jokeBlock(key));
  }
  return out;
}

// ===========================================================================
// PART 4: ANSWER KEY (two-column section)
// ===========================================================================
function buildAnswerKey() {
  const out = [];
  const S = 18; // 9pt in the key so it fits in two columns
  const body = { size: S, font: FONT, color: INK };
  out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: 'Answer Key: Step by Step', font: FONT })] }));
  out.push(para('Got one wrong? Compare your work with the steps. Find the FIRST step where yours is different. That step is the thing to practice.', { run: { ...body, italics: true }, after: 100 }));

  for (const key of ['A', 'M', 'D', 'G']) {
    const d = DOMAINS[key];
    out.push(para([new TextRun({ text: d.name.toUpperCase(), bold: true, size: 20, color: 'FFFFFF', font: FONT })], { shade: d.dark, before: 160, after: 60, keepNext: true }));
    for (const t of TOPICS.filter((x) => x.domain === key)) {
      out.push(para([new TextRun({ text: pad(t.num) + '  ' + t.title, bold: true, size: 19, color: d.dark, font: FONT })], {
        before: 120, after: 30, keepNext: true,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: d.dark, space: 1 } },
      }));
      for (const qk of ['q1', 'q2']) {
        const q = t[qk];
        const ans = q.choices ? q.answer + '  (' + q.choices['ABCD'.indexOf(q.answer)] + ')' : q.answer;
        out.push(para([
          new TextRun({ text: (qk === 'q1' ? 'Q1' : 'Q2') + '  Answer: ', bold: true, size: S, font: FONT }),
          ...rich(ans, { size: S, font: FONT, bold: true, color: d.dark }),
        ], { before: 50, after: 20, keepNext: true }));
        q.solution.forEach((s, i) => {
          const isTrap = s.startsWith('Trap:') || s.startsWith('Check') || s.startsWith('Shortcut') || s.startsWith('Sense check') || s.startsWith('Desmos') || s.startsWith('Another way') || s.startsWith('Surprise');
          const lead = isTrap ? '' : 'Step ' + (i + 1) + '.  ';
          out.push(para([
            new TextRun({ text: lead, bold: true, size: S, font: FONT, color: GREY }),
            ...rich(s, { ...body, italics: isTrap }),
          ], { indent: 200, after: 20, keepNext: i < q.solution.length - 1 }));
        });
      }
    }
    // Spot the mistake fix + joke punchline for this domain
    const m = MISTAKES[key];
    out.push(para([new TextRun({ text: 'Spot the Mistake (' + m.who + ')', bold: true, size: 19, color: d.dark, font: FONT })], { before: 140, after: 30, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: d.dark, space: 1 } } }));
    m.fix.forEach((s, i) => out.push(para(s, { run: body, indent: 200, after: 20, keepNext: i < m.fix.length - 1 })));
    const j = JOKES[key];
    out.push(para([new TextRun({ text: 'Joke Decoder', bold: true, size: 19, color: d.dark, font: FONT })], { before: 120, after: 30, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: d.dark, space: 1 } } }));
    out.push(para([new TextRun({ text: 'Answers: ', bold: true, size: S, font: FONT }), new TextRun({ text: j.slots.map((s) => s.value + ' = ' + s.word).join(',  '), size: S, font: FONT, color: INK })], { indent: 200, after: 20, keepNext: true }));
    out.push(para([new TextRun({ text: j.setup + ' ', size: S, font: FONT, color: INK }), new TextRun({ text: j.punchline, bold: true, size: S, font: FONT, color: INK })], { indent: 200, after: 60 }));
  }
  return out;
}

// ===========================================================================
// ASSEMBLE THE DOCUMENT
// ===========================================================================
const footer = new Footer({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    children: [
      new TextRun({ text: 'PSAT 8/9 Math Guide', size: 16, color: GREY, font: FONT }),
      new TextRun({ children: ['\t', PageNumber.CURRENT], size: 16, color: GREY, font: FONT }),
    ],
  })],
});
const pageProps = { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, footer: 400 } } };

const doc = new Document({
  creator: 'Math Club',
  title: 'PSAT 8/9 Math Guide',
  description: 'Formula sheet, 30 topics with worked examples and practice questions, step-by-step answer key',
  styles: {
    default: { document: { run: { font: FONT, size: SIZE, color: INK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 34, bold: true, font: FONT, color: '1F3A68' },
        paragraph: { spacing: { before: 120, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: '1F3A68' },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [
    { properties: pageProps, footers: { default: footer }, children: [...buildFormulaSheet(), ...buildIntro(), ...buildTopics()] },
    { properties: { ...pageProps, type: SectionType.NEXT_PAGE, column: { space: 500, count: 2, separate: true } }, footers: { default: footer }, children: buildAnswerKey() },
  ],
});

const outPath = path.join(__dirname, 'PSAT_8-9_Math_Guide.docx');
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log('Wrote ' + outPath);
});
