// Shared helpers for building the audit report with docx-js.
const fs = require("fs");
const path = require("path");
const {
  Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, ImageRun, VerticalAlign, ExternalHyperlink, LevelFormat,
} = require("docx");

const TEXT_W = 9412; // A4 width 11906 dxa minus 2 x 1247 dxa margins
const PX_FULL = 610; // ~ full text width in pixels at 96 dpi
const C = { ink: "0B0B0B", ink2: "52514E", navy: "1F3864", teal: "1F4E5F", hdr: "DCE6F0", grid: "BFBFBF", key: "E2EFDA", tot: "F2F2F2", note: "FFF2CC" };

const f0 = (x) => Math.round(x).toLocaleString("en-US");
const f1 = (x) => (Math.round(x * 10) / 10).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const f2 = (x) => Number(x).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const f3 = (x) => Number(x).toLocaleString("en-US", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
const pc = (x, d = 1) => (x * 100).toFixed(d) + "%";
const hk = (x) => "HK$" + f0(x);

// Inline mini-markup: **bold**, CO2 -> CO with subscript 2.
function runs(text, opts = {}) {
  const out = [];
  let bold = !!opts.bold;
  const parts = String(text).split(/(\*\*|CO2)/);
  for (const p of parts) {
    if (p === "**") { bold = !bold; continue; }
    if (p === "") continue;
    if (p === "CO2") {
      out.push(new TextRun({ text: "CO", bold, italics: opts.italics, size: opts.size, color: opts.color, font: opts.font }));
      out.push(new TextRun({ text: "2", subScript: true, bold, italics: opts.italics, size: opts.size, color: opts.color, font: opts.font }));
      continue;
    }
    out.push(new TextRun({ text: p, bold, italics: opts.italics, size: opts.size, color: opts.color, font: opts.font }));
  }
  return out;
}

const para = (text, opts = {}) => new Paragraph({
  children: runs(text, opts),
  alignment: opts.align || AlignmentType.JUSTIFIED,
  spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: 276 },
  keepNext: opts.keepNext,
});
const bulletPara = (text, level = 0, after = 60) => new Paragraph({ children: runs(text), numbering: { reference: "bullets", level }, spacing: { after, line: 264 } });

// Multi-line text becomes several paragraphs: blank lines are dropped and lines starting with "- " become bullets.
// Returns a single Paragraph for one-line text, otherwise an array (the document children are flattened).
function P(text, opts = {}) {
  const s = String(text);
  if (!s.includes("\n")) return para(s, opts);
  const lines = s.split("\n").filter((l) => l.trim() !== "");
  const isB = (l) => l && l.startsWith("- ");
  return lines.map((ln, i) => {
    const next = lines[i + 1];
    if (isB(ln)) return bulletPara(ln.slice(2), 0, next && !isB(next) ? 120 : 60);
    const o = { ...opts, before: i === 0 ? opts.before : 0 };
    if (isB(next)) o.after = 40;
    else if (next) o.after = 120;
    return para(ln, o);
  });
}
const H1 = (text, opts = {}) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(text), pageBreakBefore: opts.pageBreak !== false });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(text) });
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: runs(text) });
// A bullet whose text has "- " lines gets sub-bullets; other extra lines become indented continuation paragraphs.
function bullet(text, level = 0) {
  const s = String(text);
  if (!s.includes("\n")) return bulletPara(s, level);
  const lines = s.split("\n").filter((l) => l.trim() !== "");
  return lines.map((ln, i) => {
    if (i === 0) return bulletPara(ln, level, 40);
    if (ln.startsWith("- ")) return bulletPara(ln.slice(2), level + 1, 40);
    return new Paragraph({ children: runs(ln), indent: { left: 400 * (level + 1) }, spacing: { before: 40, after: 80, line: 264 } });
  });
}
const bullets = (arr) => arr.map((t) => bullet(t));
const numbered = (arr, ref = "numbers") => arr.map((t) => new Paragraph({ children: runs(t), numbering: { reference: ref, level: 0 }, spacing: { after: 60, line: 264 } }));

function caption(text) {
  return new Paragraph({ children: runs(text, { italics: true, size: 17, color: C.ink2 }), alignment: AlignmentType.CENTER, spacing: { before: 60, after: 200 } });
}
function tableCaption(text) {
  return new Paragraph({ children: runs(text, { bold: true, size: 18, color: C.navy }), spacing: { before: 160, after: 80 }, keepNext: true });
}

const border = { style: BorderStyle.SINGLE, size: 4, color: C.grid };
const BORDERS = { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };
const NOBORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NOBORDERS = { top: NOBORDER, bottom: NOBORDER, left: NOBORDER, right: NOBORDER, insideHorizontal: NOBORDER, insideVertical: NOBORDER };

function cell(content, width, opts = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((t) => (t instanceof Paragraph ? t : new Paragraph({
    children: runs(t, { bold: opts.bold, size: opts.size || 17, color: opts.color }),
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { after: 0, line: 240 },
  })));
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    verticalAlign: opts.valign || VerticalAlign.CENTER,
    columnSpan: opts.span,
  });
}

// rows: array of arrays of strings; widths in dxa summing to TEXT_W; align: array of "L"/"R"/"C"
function table(headers, rows, widths, opts = {}) {
  const sum = widths.reduce((a, b) => a + b, 0);
  if (Math.abs(sum - TEXT_W) > 2) throw new Error(`table widths ${sum} != ${TEXT_W}: ${headers.join("|")}`);
  const al = (i) => ({ R: AlignmentType.RIGHT, C: AlignmentType.CENTER }[(opts.align || [])[i]] || AlignmentType.LEFT);
  const trs = [];
  if (headers) {
    trs.push(new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { bold: true, fill: C.hdr, align: al(i), color: C.navy, size: opts.size })) }));
  }
  rows.forEach((r, ri) => {
    const isTotal = opts.totalRows && opts.totalRows.includes(ri);
    trs.push(new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, widths[i], { align: al(i), bold: isTotal, fill: isTotal ? C.tot : (opts.keyCol === i ? C.key : undefined), size: opts.size })) }));
  });
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths, rows: trs, borders: BORDERS });
}

// Two-column key/value table
function kvTable(pairs, widths = [3300, TEXT_W - 3300], opts = {}) {
  return table(opts.headers || null, pairs, widths, { size: opts.size, keyCol: opts.keyCol });
}

function img(file, widthPx, sizePx) {
  const [w, h] = sizePx;
  const height = Math.round((widthPx * h) / w);
  const ext = path.extname(file).toLowerCase().replace(".", "");
  return new ImageRun({ type: ext === "jpg" || ext === "jpeg" ? "jpg" : "png", data: fs.readFileSync(file), transformation: { width: widthPx, height } });
}
function figure(file, sizePx, cap, widthPx = PX_FULL) {
  return [new Paragraph({ children: [img(file, widthPx, sizePx)], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 }, keepNext: true }), caption(cap)];
}
// Row of photos in a borderless table: items [{file, size:[w,h], cap, h?}] - images scaled to a common height
function photoRow(items, heightPx = 230) {
  const colW = Math.floor(TEXT_W / items.length);
  const widths = items.map((_, i) => (i === items.length - 1 ? TEXT_W - colW * (items.length - 1) : colW));
  const maxW = Math.floor((colW / 1440) * 96) - 12;
  const cells = items.map((it, i) => {
    let wpx = Math.round((heightPx * it.size[0]) / it.size[1]);
    if (wpx > maxW) wpx = maxW;
    return new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      borders: { top: NOBORDER, bottom: NOBORDER, left: NOBORDER, right: NOBORDER },
      children: [
        new Paragraph({ children: [img(it.file, wpx, it.size)], alignment: AlignmentType.CENTER, spacing: { after: 40 }, keepNext: true, keepLines: true }),
        new Paragraph({ children: runs(it.cap, { italics: true, size: 16, color: C.ink2 }), alignment: AlignmentType.CENTER, spacing: { after: 80 }, keepNext: true, keepLines: true }),
      ],
    });
  });
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths, rows: [new TableRow({ cantSplit: true, children: cells })], borders: NOBORDERS });
}

// Photo rows plus their figure caption in a one-cell table whose row cannot split, so the caption
// always stays on the same page as its photos. Pass one row of items or an array of rows.
function photoFigure(rowsOrItems, heightPx, cap) {
  const rows = Array.isArray(rowsOrItems[0]) ? rowsOrItems : [rowsOrItems];
  const children = [];
  rows.forEach((items, i) => {
    if (i > 0) children.push(new Paragraph({ children: [], spacing: { after: 0, line: 120 } }));
    children.push(photoRow(items, heightPx));
  });
  children.push(caption(cap));
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: [TEXT_W], borders: NOBORDERS,
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: TEXT_W, type: WidthType.DXA },
      borders: { top: NOBORDER, bottom: NOBORDER, left: NOBORDER, right: NOBORDER },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children,
    })] })],
  });
}

function noteBox(lines, fill = C.note) {
  const paras = lines.map((t) => new Paragraph({ children: runs(t, { size: 18 }), spacing: { after: 60, line: 252 } }));
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: [TEXT_W], borders: BORDERS,
    rows: [new TableRow({ children: [new TableCell({ width: { size: TEXT_W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, margins: { top: 100, bottom: 60, left: 140, right: 140 }, children: paras })] })],
  });
}

function link(text, url) {
  return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink", size: 17 })], link: url });
}

const spacer = (after = 120) => new Paragraph({ children: [], spacing: { after } });

const NUMBERING = {
  config: [
    { reference: "bullets", levels: [
      { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 260 } } } },
    ] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
    { reference: "numbers2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
    { reference: "numbers3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
  ],
};

const STYLES = {
  default: { document: { run: { font: "Arial", size: 20, color: C.ink } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: "Arial", size: 30, bold: true, color: C.navy }, paragraph: { spacing: { before: 120, after: 160 }, keepNext: true, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: "Arial", size: 24, bold: true, color: C.teal }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: "Arial", size: 21, bold: true, color: C.ink }, paragraph: { spacing: { before: 180, after: 80 }, keepNext: true, outlineLevel: 2 } },
    { id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal",
      run: { font: "Arial", size: 20, bold: true, color: C.navy }, paragraph: { spacing: { before: 100, after: 20 } } },
    { id: "TOC2", name: "toc 2", basedOn: "Normal", next: "Normal",
      run: { font: "Arial", size: 19, color: C.ink }, paragraph: { spacing: { after: 10 }, indent: { left: 360 } } },
  ],
};

module.exports = { TEXT_W, PX_FULL, C, f0, f1, f2, f3, pc, hk, runs, P, para, H1, H2, H3, bullet, bullets, numbered, caption, tableCaption, table, kvTable, img, figure, photoRow, photoFigure, noteBox, link, spacer, NUMBERING, STYLES, cell, BORDERS };
