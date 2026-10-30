// Helpers comunes — estilo Certus (Calibri, navy/oro)
const d = require('docx');
const {
  Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, LevelFormat, PageBreak, ImageRun, VerticalAlign
} = d;

const NAVY = '1F3864';
const NAVY_DK = '0A1628';
const GOLD = 'B08D57';
const GOLD_LT = 'F4ECDD';
const GRAY = 'F2F4F8';
const RED_LT = 'FBEAEA';
const RED = '9B2C2C';
const W = 9360; // ancho útil carta con márgenes de 1"

const thin = { style: BorderStyle.SINGLE, size: 4, color: 'BFC5D2' };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: none, bottom: none, left: none, right: none };

// Texto con **negritas** inline
function runs(text, opts = {}) {
  const out = [];
  String(text).split('\n').forEach((line, li) => {
    const parts = line.split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
    parts.forEach((p, pi) => {
      const bold = p.startsWith('**') && p.endsWith('**');
      out.push(new TextRun({ text: bold ? p.slice(2, -2) : p, bold: bold || opts.bold, color: opts.color, size: opts.size, italics: opts.italics, font: 'Calibri', break: (li > 0 && pi === 0) ? 1 : undefined }));
    });
  });
  return out;
}

function P(text, opts = {}) {
  return new Paragraph({
    children: runs(text, opts),
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: 276 },
    keepNext: opts.keepNext,
  });
}

function H1(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: NAVY, size: 26, font: 'Calibri' })],
    spacing: { before: 300, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 4 } },
    keepNext: true,
  });
}

function H2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: NAVY, size: 22, font: 'Calibri' })],
    spacing: { before: 200, after: 100 },
    keepNext: true,
  });
}

function B(text, level = 0) {
  return new Paragraph({ children: runs(text), numbering: { reference: 'bul', level }, spacing: { after: 70, line: 264 }, alignment: AlignmentType.LEFT });
}
function N(text, ref = 'num') {
  return new Paragraph({ children: runs(text), numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 264 }, alignment: AlignmentType.LEFT });
}
function CB(text) { // casilla de verificación
  return new Paragraph({ children: [new TextRun({ text: '☐  ', font: 'Segoe UI Symbol', size: 22, color: NAVY }), ...runs(text)], spacing: { after: 70 }, indent: { left: 360, hanging: 360 } });
}

function cell(content, width, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(c =>
    c instanceof Paragraph ? c : new Paragraph({ children: runs(c, { bold: o.bold, color: o.color, size: o.size || 19 }), alignment: o.align || AlignmentType.LEFT, spacing: { after: 40 } })
  );
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    borders: o.borders || borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    verticalAlign: o.valign || VerticalAlign.TOP,
    columnSpan: o.span,
  });
}

// Tabla con encabezado navy
function T(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const hr = new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { fill: NAVY, color: 'FFFFFF', bold: true })) });
  const body = rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? GRAY : undefined })) }));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [hr, ...body] });
}

// Tabla de dos columnas etiqueta/valor
function KV(rows, widths = [2800, W - 2800]) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map(([k, v]) => new TableRow({ children: [cell(k, widths[0], { bold: true, fill: GRAY }), cell(v, widths[1])] }))
  });
}

// Recuadro de aviso
function BOX(title, lines, kind = 'gold') {
  const fill = kind === 'red' ? RED_LT : kind === 'navy' ? 'E8ECF4' : GOLD_LT;
  const col = kind === 'red' ? RED : NAVY;
  const bar = { style: BorderStyle.SINGLE, size: 24, color: kind === 'red' ? RED : GOLD };
  const children = [];
  if (title) children.push(new Paragraph({ children: [new TextRun({ text: title, bold: true, color: col, size: 21, font: 'Calibri' })], spacing: { after: 80 } }));
  lines.forEach(l => children.push(l instanceof Paragraph ? l : new Paragraph({ children: runs(l, { size: 20 }), spacing: { after: 60 } })));
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      children, width: { size: W, type: WidthType.DXA },
      borders: { top: none, bottom: none, right: none, left: bar },
      shading: { fill, type: ShadingType.CLEAR, color: 'auto' },
      margins: { top: 120, bottom: 120, left: 200, right: 160 },
    })] })]
  });
}

function SP(after = 120) { return new Paragraph({ children: [], spacing: { after } }); }
function BR() { return new Paragraph({ children: [new PageBreak()] }); }

// Encabezado de documento del Kit (mismo patrón que D0–D14)
function kitHeader(code, title, scope) {
  const p = new Paragraph({
    children: [
      new TextRun({ text: 'CERTUS PLD', bold: true, color: GOLD, size: 22, font: 'Calibri' }),
      new TextRun({ text: '   │   ', color: 'FFFFFF', size: 22 }),
      new TextRun({ text: code + ' ', bold: true, color: 'FFFFFF', size: 26, font: 'Calibri' }),
      new TextRun({ text: title, bold: true, color: 'FFFFFF', size: 26, font: 'Calibri' }),
    ], spacing: { after: 60 }
  });
  const p2 = new Paragraph({ children: [
    new TextRun({ text: 'Empresa: [Empresa]', color: 'D9DEE8', size: 19 }),
    new TextRun({ text: '   │   ' + scope, color: 'D9DEE8', size: 19 }),
  ] });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [
    new TableCell({ children: [p, p2], width: { size: W, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: { style: BorderStyle.SINGLE, size: 18, color: GOLD } }, shading: { fill: NAVY, type: ShadingType.CLEAR, color: 'auto' }, margins: { top: 200, bottom: 200, left: 240, right: 240 } })
  ] })] });
}

const numbering = {
  config: [
    { reference: 'bul', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } }, run: { color: GOLD } } },
      { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } },
    ] },
    ...['num', 'num2', 'num3', 'num4', 'num5', 'num6', 'num7', 'num8'].map(r => ({ reference: r, levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } }, run: { bold: true, color: NAVY } } }
    ] })),
  ]
};

const styles = { default: { document: { run: { font: 'Calibri', size: 21 } } } };
const page = { size: { width: 12240, height: 15840 }, margin: { top: 1300, bottom: 1200, left: 1440, right: 1440 } };

module.exports = { d, NAVY, NAVY_DK, GOLD, GOLD_LT, GRAY, RED, W, runs, P, H1, H2, B, N, CB, cell, T, KV, BOX, SP, BR, kitHeader, numbering, styles, page, borders, noBorders };
