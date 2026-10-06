#!/usr/bin/env node
/**
 * build-docx.js - Applies a MANIFEST.json and produces the .docx in APA 7 format.
 *
 * This file contains NO content from ANY document. Back in the day every job
 * meant editing a generator of hundreds of lines with the title,
 * the cover page, the tables and the 5 references typed by hand (and the
 * references were emitted ALWAYS, even if the .md had others). Here everything
 * arrives from MANIFIEST.json, which scripts\md-a-manifiesto.py produces.
 *
 * APA 7 rules it applies (references\apa7-format.md is the source):
 *   - Times New Roman 12, double spacing, no extra space before/after.
 *   - 0.5" first-line indent in body paragraphs.
 *   - Text aligned to the LEFT (not justified to both margins).
 *   - Level 1 centered and bold; level 2 flush left and bold.
 *   - Page number in the TOP RIGHT CORNER (header, not footer).
 *   - The cover page has NO page number; visible numbering starts on
 *     page 2, which is what APA 7 requires.
 *   - Tables and figures: number in bold, title in italics, then the
 *     object. Horizontal borders only in tables. The note ("Nota. ...")
 *     goes BELOW the table or the image, never above it.
 *   - Table of contents / list of tables / list of figures: real TOC
 *     fields, Times 12, entries not bold.
 *   - References: 0.5" hanging indent, alphabetical order.
 *
 * Usage:
 *   node build-docx.js --manifiesto MANIFIEST.json --out salida.docx
 *                      [--log salida\_logs\02-docx.log]
 */

"use strict";

const fs = require("fs");
const path = require("path");

/**
 * Node resolves `require` from the SCRIPT's folder, not from the current
 * folder. Since `docx` is installed in <skill>\.work and this script lives in
 * <skill>\scripts, a normal `require("docx")` always fails.
 *
 * The module path is resolved explicitly instead of forcing the user to
 * copy the .js inside .work or to install docx globally.
 */
const RAIZ_SKILL = path.resolve(__dirname, "..");
const WORKDIR = process.env.APA7_WORKDIR
  ? path.resolve(process.env.APA7_WORKDIR)
  : path.join(RAIZ_SKILL, ".work");
const RUTA_DOCX = path.join(WORKDIR, "node_modules", "docx");

if (!fs.existsSync(RUTA_DOCX)) {
  console.error("FAILURE: the 'docx' package was not found in");
  console.error("       " + RUTA_DOCX);
  console.error("Run scripts/apa7.py install to install it in the skill's workdir.");
  process.exit(3);
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Bookmark, SimpleField, PageBreak, Footer, Header, PageNumber, AlignmentType,
  BorderStyle, WidthType, ShadingType, VerticalAlign,
  SectionType, PageOrientation, TableLayoutType, HeightRule, convertInchesToTwip,
  LineRuleType, TableOfContents,
} = require(RUTA_DOCX);

// ---------------------------------------------------------------------------
// Parameters
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i++) {
    if (argv[i].startsWith("--")) {
      const k = argv[i].slice(2);
      const v = (argv[i + 1] && !argv[i + 1].startsWith("--")) ? argv[++i] : "true";
      out[k] = v;
    }
  }
  return out;
}

const args = parseArgs(process.argv);
if (!args.manifiesto || !args.out) {
  console.error("Usage: node build-docx.js --manifiesto MANIFEST.json --out salida.docx [--log log.txt]");
  process.exit(2);
}

const logLines = [];
function log(msg) {
  const line = `  ${msg}`;
  logLines.push(line);
  console.log(line);
}

// readFileSync with "utf8" does NOT strip the BOM and JSON.parse aborts if there
// is one: the parser writes the manifest without a BOM, but if the user opens it
// and saves it from a Windows editor, it comes back with a BOM.
const M = JSON.parse(fs.readFileSync(args.manifiesto, "utf8").replace(/^\uFEFF/, ""));
const DOCX = args.out;

// Real pages of the entries, measured from a first export in apa7.py's double
// pass (--paginas-json). A TOC field's result is computed by the PDF engine, not
// by this script, so the only way for the .docx to already carry correct numbers
// is to cache what the first pass measured. Without this file the entries render
// without page numbers and Word fills them in when the document is opened.
const PAGINAS = (args["paginas-json"] && fs.existsSync(args["paginas-json"]))
  ? JSON.parse(fs.readFileSync(args["paginas-json"], "utf8").replace(/^\uFEFF/, ""))
  : { secciones: {}, tablas: {}, figuras: {} };

function paginaDe(grupo, clave) {
  const valor = (PAGINAS[grupo] || {})[String(clave)];
  return (valor === undefined || valor === null || valor === "") ? undefined : valor;
}

// --- Critical manifest validation (abort, do not guess) ---------------------
// Back in the day the document was delivered with blank table and figure
// titles: the index printed "(sin titulo)" and the body skipped the whole title,
// without warning anywhere. An APA table or figure without a title is a formal
// defect, not a preference, so the process is cut here.
const faltantes = [];
for (const t of M.tablas || []) {
  if (!t.titulo || !String(t.titulo).trim()) {
    faltantes.push(`Table ${t.indice}: no title`);
  }
}
for (const f of M.figuras || []) {
  if (!f.titulo || !String(f.titulo).trim()) {
    faltantes.push(`Figure ${f.indice}: no caption`);
  }
}
if (faltantes.length) {
  console.error("ERROR: the manifest has APA objects without a title. The .docx is not generated.");
  for (const m of faltantes) console.error(`  - ${m}`);
  console.error("");
  console.error("Fix them in the parser (not in this script): the .md, a");
  console.error("--titulos-tabla.json / --titulos-figura-json, or by asking the user.");
  console.error("Check M.diagnostico.preguntas and M.diagnostico.pendientes_bloqueantes.");
  process.exit(4);
}

// ---------------------------------------------------------------------------
// Formatting constants
// ---------------------------------------------------------------------------

const FUENTE = "Times New Roman";
const TAM = 24;                       // 12 pt = 24 half-points
const DOBLE = 480;                    // double spacing (twips: 240 = single)
const SANGRIA_1RA = 720;              // 0.5 inch
const ANCHO_CONTENIDO = 9360;         // 8.5" - 1" - 1" = 6.5" in twips
const TIPO_PAGINA = { width: 12240, height: 15840 }; // Letter
const MARGEN = 1440;                  // 1 inch

// Landscape page for wide tables. APA 7 allows it: a table that does not fit
// comfortably in portrait is shown on a separate landscape page, keeping
// 1-inch margins. The usable width goes from 6.5" to 9".
const ANCHO_CONTENIDO_H = TIPO_PAGINA.height - 2 * MARGEN; // 11" - 2" = 9"

// Criterion to decide whether a table needs a landscape page. It triggers on
// WIDTH, not on row count: a 200-row, 3-column table reads fine in portrait,
// while an 8-column one ends up crushed.
const UMBRAL_COLUMNAS_ANCHAS = 6;    // from 6 columns on
const UMBRAL_LINEAS_APRETADAS = 8;    // lines per cell that are no longer readable
const ANCHO_CARACTER_TWIP = 120;      // average character width at 12 pt
const MARGEN_CELDA_LR = 160;          // cell left and right margins

// The "Nota." label is written by this script, in Spanish, before the note's
// text. A note that already carries the label (the user pasted it, or the agent
// wrote "Nota. Elaboracion propia" because that is how it reads in a document)
// would otherwise be printed as "Nota. Nota. Elaboracion propia". Only a
// LEADING label is removed, and only once; a note that merely mentions the word
// ("Segun la nota del autor...") is left untouched.
function sinEtiquetaNota(texto) {
  return String(texto == null ? "" : texto)
    .replace(/^\s*(nota|note)\s*[.:]\s*/i, "");
}

// The whole legend of a caption is bold ("Tabla 1.", "Figura 2."), the title
// that follows is italics. SimpleField writes its cached result with a DEFAULT
// run, so the number always came out plain and only the word was bold. This
// subclass keeps the field (a real SEQ, which is what "TOC \c" collects) and
// writes the cached number with the same run properties as the label.
// The instruction is passed to super() WITHOUT the cached value on purpose:
// super() would already have pushed its own run for it.
class CampoSecuencia extends SimpleField {
  constructor(instruccion, valor) {
    super(instruccion);
    this.root.push(new TextRun({ text: String(valor), bold: true, font: FUENTE, size: TAM }));
  }
}

// docx@9 always emits a <w:tblGrid>, and when columnWidths is omitted it fills
// that grid with 100 twips per column. The per-cell widths (tcW) are written too,
// so which of the two wins depends on the consumer. Passing columnWidths makes
// the grid agree with the cells, so the layout comes out the same in Word,
// LibreOffice and any other renderer. The LAST column absorbs the remainder so
// the widths add up to exactly `anchoTotal`: ncols * floor(ancho/ncols) is short
// by one or two twips on most divisions, and the table would not match the width
// it declares.
function repartirColumnas(anchoTotal, n) {
  if (!(n > 0)) return [];
  const base = Math.floor(anchoTotal / n);
  const out = new Array(n).fill(base);
  out[n - 1] += anchoTotal - base * n;
  return out;
}

// Every table is built here. The missing-columnWidths bug was SILENT: the .docx
// was produced, the pipeline advanced, and the only symptom was the rendered
// PDF. A table without columnWidths now aborts the build instead of shipping.
function crearTabla(opts) {
  if (!opts.columnWidths || !opts.columnWidths.length) {
    throw new Error("crearTabla: 'columnWidths' is required (use repartirColumnas)");
  }
  return new Table(opts);
}

// Cover page in 3 real vertical zones: title at the top, members in the middle
// of the page and the institutional block anchored at the bottom. It is built as
// a borderless table of 1 column x 3 rows, with fixed height and vertical
// alignment, instead of with filler empty paragraphs: that way the zones do not
// move depending on how many members or fields the document brings.
//
// INVARIANT: the three rows must add up to EXACTLY ALTO_UTIL. If they add up to
// more, the last row overflows onto a new page and the document gains a page.
// That is why the third zone absorbs the rounding instead of being a fixed third.
const ALTO_UTIL = TIPO_PAGINA.height - 2 * MARGEN;      // 9" = 12960 twips = 648 pt
const ZONA_ALTO = Math.round(ALTO_UTIL / 3);            // 4320: y 72..288 pt, logo + title
const ZONA_CENTRO = Math.round(ALTO_UTIL / 3);          // 4320: y 288..504 pt, members
const ZONA_BAJO = ALTO_UTIL - ZONA_ALTO - ZONA_CENTRO;  // 4320: y 504..720 pt, block
const AIRE_ZONA_ALTA = 720;          // the title starts ~0.5" below the margin

const B = (color) => ({ style: BorderStyle.SINGLE, size: 4, color: color || "000000" });
const SIN_BORDE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

/** Text segments from the manifest -> TextRun[], honoring bold/italics. */
function runsDe(segmentos, extra = {}) {
  if (!segmentos || !segmentos.length) return [new TextRun({ text: "", ...extra })];
  return segmentos.map((s) => new TextRun({
    text: s.text,
    bold: !!s.bold || !!extra.bold,
    italics: !!s.italics || !!extra.italics,
    font: FUENTE,
    size: TAM,
    ...(extra.fontSize ? { size: extra.fontSize } : {}),
  }));
}

/** Inserts italics only in the confirmed range of a reference. */
function runsDeReferencia(ref) {
  const texto = ref.texto;
  const rango = ref.cursiva_confirmada || ref.cursiva_propuesta;
  if (!rango) {
    return [new TextRun({ text: texto, font: FUENTE, size: TAM })];
  }
  const i = texto.indexOf(rango);
  if (i < 0) {
    log(`    [warning] the italic range "${rango.slice(0, 40)}..." does not appear in the reference; it is left in plain text.`);
    return [new TextRun({ text: texto, font: FUENTE, size: TAM })];
  }
  const antes = texto.slice(0, i);
  const despues = texto.slice(i + rango.length);
  const partes = [];
  if (antes) partes.push(new TextRun({ text: antes, font: FUENTE, size: TAM }));
  partes.push(new TextRun({ text: rango, italics: true, font: FUENTE, size: TAM }));
  if (despues) partes.push(new TextRun({ text: despues, font: FUENTE, size: TAM }));
  return partes;
}

// ---------------------------------------------------------------------------
// Header with page number (top right corner, APA 7)
// ---------------------------------------------------------------------------

const headerConPagina = new Header({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { line: 240, before: 0, after: 0 },
    children: [new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: TAM })],
  })],
});

const headerVacio = new Header({ children: [new Paragraph({ children: [] })] });
const footerVacio = new Footer({ children: [new Paragraph({ children: [] })] });

// ---------------------------------------------------------------------------
// Cover page
// ---------------------------------------------------------------------------

function parrafoCentrado(children, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: DOBLE, before: 0, after: 0 },
    children,
    ...opts,
  });
}

// A line of the cover page, centered.
function lineaPortada(children, spacing) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: spacing || { line: DOBLE, before: 0, after: 0 },
    children,
  });
}

/**
 * Cover page in 3 vertical zones: logo + title at the top, members at the vertical
 * center of the page and the remaining institutional data anchored at the bottom.
 *
 * Back in the day this was done by stacking paragraphs and padding with empty
 * lines, which left the members glued to the title and the institutional block
 * in the middle of the page.
 *
 * It is done with ONE table of 1 column x 3 rows, with no visible borders, with a
 * fixed height per row and vertical alignment. It is a table and not three stacked
 * on purpose: two or more consecutive tables get merged by Word and LibreOffice
 * treats them inconsistently, and the set overflows onto an extra page.
 */
function construirPortada() {
  const p = M.portada || {};
  const centro = (text, bold = false) => [new TextRun({ text, bold, font: FUENTE, size: TAM })];

  // --- Top zone: optional logo and title -----------------------------------
  const zonaAlta = [];
  if (p.logo) {
    // The logo is an optional resource supplied by the user. An institution's
    // name is never invented.
    const ext = path.extname(p.logo).toLowerCase();
    const tipo = ext === ".jpg" || ext === ".jpeg" ? "jpg" : (ext === ".gif" ? "gif" : "png");
    zonaAlta.push(lineaPortada([new ImageRun({
      type: tipo,
      data: fs.readFileSync(p.logo),
      transformation: { width: 160, height: 160 },
        })], { line: 240, before: 0, after: 0, lineRule: LineRuleType.AUTO }));
   // single: an image has no line spacing
    if (p.logo_leyenda) zonaAlta.push(lineaPortada(centro(p.logo_leyenda)));
  }
  // With no logo NOTHING is inserted: neither a placeholder, nor a frame, nor the
  // word "Logo" that used to sneak into the PDF.
  if (p.titulo) zonaAlta.push(lineaPortada(centro(p.titulo, true)));

  // --- Central zone: members ------------------------------------------------
  // references\institutional-cover.md: the members go in a SINGLE paragraph
  // centered, separated by comma and with "y" before the last one, regardless of
  // how many there are. Back in the day one paragraph per person was emitted.
  const zonaCentro = [];
  if (p.autores && p.autores.length) {
    const nombres = p.autores.map((a) => String(a).trim()).filter(Boolean);
    if (nombres.length === 1) {
      zonaCentro.push(lineaPortada(centro(nombres[0])));
    } else if (nombres.length === 2) {
      zonaCentro.push(lineaPortada(centro(`${nombres[0]} y ${nombres[1]}`)));
    } else {
      zonaCentro.push(lineaPortada(centro(`${nombres.slice(0, -1).join(", ")} y ${nombres[nombres.length - 1]}`)));
    }
  }

  // --- Bottom zone: institutional data --------------------------------------
  // The order depends on what the document brings: if a field is absent, no
  // blank line is left. The missing ones are simply not emitted.
  const zonaBaja = [];
  if (p.docente) {
    const d = [p.docente, p.docente_titulo].filter(Boolean).join(", ");
    zonaBaja.push(lineaPortada(centro(d)));
  }
  if (p.materia_nrc) zonaBaja.push(lineaPortada(centro(p.materia_nrc)));
  if (p.facultad) zonaBaja.push(lineaPortada(centro(p.facultad)));
  if (p.vicerrectoria) zonaBaja.push(lineaPortada(centro(p.vicerrectoria)));
  if (p.sede) zonaBaja.push(lineaPortada(centro(p.sede)));
  if (p.fecha) zonaBaja.push(lineaPortada(centro(p.fecha)));

  const filaZona = (children, alto, alineacion, margenSuperior = 0) => new TableRow({
    // ATLEAST and not EXACT: if the content does not fit (a 3-line title, 6
    // members) the row grows instead of cutting the text.
    height: { value: alto, rule: HeightRule.ATLEAST },
    children: [new TableCell({
      children: children.length ? children : [lineaPortada([new TextRun({ text: "", size: TAM })])],
      width: { size: ANCHO_CONTENIDO, type: WidthType.DXA },
      verticalAlign: alineacion,
      margins: { top: margenSuperior, bottom: 0, left: 0, right: 0 },
      borders: {
        top: SIN_BORDE, bottom: SIN_BORDE, left: SIN_BORDE, right: SIN_BORDE,
        insideHorizontal: SIN_BORDE, insideVertical: SIN_BORDE,
      },
    })],
  });

  return [crearTabla({
    width: { size: ANCHO_CONTENIDO, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    // The cover grid: a single column as wide as the usable text width, so the
    // grid docx emits agrees with the cell width set in filaZona.
    columnWidths: [ANCHO_CONTENIDO],
    // The top zone's spacing is SUBTRACTED from its height: the rows have to add
    // up to exactly the usable height (12960 twips). Adding more overflows the
    // last row onto a new page.
    rows: [
      filaZona(zonaAlta, ZONA_ALTO - AIRE_ZONA_ALTA, VerticalAlign.TOP, AIRE_ZONA_ALTA),
      filaZona(zonaCentro, ZONA_CENTRO, VerticalAlign.CENTER),
      filaZona(zonaBaja, ZONA_BAJO, VerticalAlign.BOTTOM),
    ],
  })];
}

// ---------------------------------------------------------------------------
// Table of contents, list of tables and list of figures (real Word fields)
// ---------------------------------------------------------------------------
//
// These are REAL TOC fields. A hand-written PAGEREF list is not a TOC: Word
// cannot update it and it does not survive editing. A TOC field, on the other
// hand, is only computed by the rendering engine (LibreOffice on export, Word on
// open), so a single pass cannot know the page numbers. apa7.py therefore runs
// two passes: it exports, measures the real pages, and rebuilds with
// --paginas-json, whose numbers become the field's cached result. If the file is
// absent the entries still print, just without numbers, and Word fills them in
// when the document is opened (features.updateFields).
//
// The field switch that selects entries differs per index:
//   content  -> \u  (paragraphs with an applied outline level)
//   tables   -> \c "Tabla"   (captions with a SEQ Tabla field)
//   figures  -> \c "Figura"  (captions with a SEQ Figura field)
// That is why the headings carry `outlineLevel` and the captions a real SEQ.

// One cached entry per line: title, level (-> TOC1..TOC5 style), page, target.
function entradasDeSecciones(opts) {
  // Repeated titles are disambiguated ONLY if the manifest asks for it, so that
  // the TOC is navigable without altering the document's text.
  const vistos = new Map();
  const salida = [];
  for (const s of M.secciones) {
    let texto = s.texto;
    const n = (vistos.get(texto) || 0) + 1;
    vistos.set(texto, n);
    if (n > 1 && opts.dedupe_toc) {
      texto = `${texto} (${s.hid})`;
      log(`    TOC: repeated title -> "${texto}"`);
    }
    salida.push({
      title: texto,
      level: s.nivel,
      page: paginaDe("secciones", s.hid),
      href: `apa_sec_${s.hid}`,
    });
  }
  return salida;
}

function entradasDeLista(grupo, etiqueta) {
  // Only the figures that actually make it into the body. Listing the omitted
  // ones too leaves an entry whose bookmark is never created.
  return (M[grupo] || [])
    .filter((o) => (grupo === "figuras" ? o.existe : true))
    .map((o) => ({
      title: `${etiqueta} ${o.indice}. ${o.titulo}`,
      level: 1,
      page: paginaDe(grupo, o.indice),
      href: grupo === "tablas" ? `apa_tbl_${o.indice}` : `apa_fig_${o.indice}`,
    }));
}

function crearTOC(alias, entradas, extra = {}) {
  return new TableOfContents(alias, {
    hyperlink: true,
    beginDirty: true,
    cachedEntries: entradas,
    ...extra,
  });
}

function construirIndices() {
  const hijos = [];
  const opts = M.opciones || {};

  // --- Table of contents: ALWAYS --------------------------------------------
  hijos.push(parrafoCentrado([new TextRun({ text: "Tabla de contenido", bold: true, font: FUENTE, size: TAM })]));
  hijos.push(crearTOC("Contenido", entradasDeSecciones(opts), {
    useAppliedParagraphOutlineLevel: true,
  }));

  // --- List of tables: only if there ARE tables -----------------------------
  if (opts.indice_tablas && M.tablas && M.tablas.length) {
    hijos.push(new Paragraph({ children: [new PageBreak()] }));
    hijos.push(parrafoCentrado([new TextRun({ text: "Indice de tablas", bold: true, font: FUENTE, size: TAM })]));
    hijos.push(crearTOC("Tablas", entradasDeLista("tablas", "Tabla"), {
      captionLabelIncludingNumbers: "Tabla",
    }));
  } else if (M.tablas && M.tablas.length) {
    log("    List of tables OMITTED by explicit manifest decision.");
  }

  // --- List of figures: only if there ARE figures ---------------------------
  if (opts.indice_figuras && M.figuras && M.figuras.length) {
    const figsValidas = M.figuras.filter((f) => f.existe);
    const omitidas = M.figuras.length - figsValidas.length;
    if (omitidas > 0) {
      log(`    List of figures: ${omitidas} line(s) are omitted for a missing figure.`);
    }
    if (figsValidas.length) {
      hijos.push(new Paragraph({ children: [new PageBreak()] }));
      hijos.push(parrafoCentrado([new TextRun({ text: "Indice de figuras", bold: true, font: FUENTE, size: TAM })]));
      hijos.push(crearTOC("Figuras", entradasDeLista("figuras", "Figura"), {
        captionLabelIncludingNumbers: "Figura",
      }));
    }
  } else if (M.figuras && M.figuras.length) {
    log("    List of figures OMITTED by explicit manifest decision.");
  }

  // The body ALWAYS starts on a new page. It is not left to be pushed by the
  // first level 1 heading: in the body that heading goes first, so its "if
  // something has already been written" guard is never met and the text used to
  // hang below the list of tables, on the same page.
  if (hijos.length) {
    hijos.push(new Paragraph({ children: [new PageBreak()] }));
  }

  return hijos;
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

// Estimates how many lines the text of a cell takes if the table is laid out in
// portrait. With `layout: FIXED` each column measures ancho_total / ncols, so the
// number of characters per line comes from there, not from the real content.
function lineasDeCelda(texto, anchoColumna) {
  const utiles = anchoColumna - MARGEN_CELDA_LR;
  if (utiles <= 0) return Number.POSITIVE_INFINITY;
  const porLinea = Math.floor(utiles / ANCHO_CARACTER_TWIP);
  if (porLinea < 1) return Number.POSITIVE_INFINITY;
  return texto.split("\n").reduce((total, linea) => total + Math.max(1, Math.ceil(linea.length / porLinea)), 0);
}

// Decides whether a table goes on a landscape page and records the reason, which
// is printed in the log and saved in the manifest so it can be audited.
function evaluarAncho(t) {
  const ncols = Math.max(...t.filas.map((f) => f.length));
  const columnaVertical = ANCHO_CONTENIDO / ncols;
  const columnaApaisada = ANCHO_CONTENIDO_H / ncols;

  let peorVertical = 0;
  let peorApaisada = 0;
  for (const fila of t.filas) {
    for (const celda of fila) {
      const texto = String(celda || "");
      peorVertical = Math.max(peorVertical, lineasDeCelda(texto, columnaVertical));
      peorApaisada = Math.max(peorApaisada, lineasDeCelda(texto, columnaApaisada));
    }
  }

  // The decision is made with the table in portrait. If it is already cramped
  // there but reads well in landscape, landscape is the solution, not just a luxury.
  let motivo = "";
  if (ncols >= UMBRAL_COLUMNAS_ANCHAS) {
    motivo = `${ncols} columns (threshold ${UMBRAL_COLUMNAS_ANCHAS})`;
  } else if (peorVertical >= UMBRAL_LINEAS_APRETADAS && peorApaisada < UMBRAL_LINEAS_APRETADAS) {
    motivo = `cell of up to ${peorVertical} lines in portrait, ${peorApaisada} in landscape`;
  }

  return {
    ancha: motivo !== "",
    ncols,
    motivo,
    peorVertical: Number.isFinite(peorVertical) ? peorVertical : -1,
    peorApaisada: Number.isFinite(peorApaisada) ? peorApaisada : -1,
  };
}

function construirTabla(t, ancho = ANCHO_CONTENIDO) {
  const hijos = [];
  const ncols = Math.max(...t.filas.map((f) => f.length));
  // Computed ONCE here and reused for both the cell widths (tcW) and the table
  // grid (columnWidths): if they were computed separately they could disagree.
  const anchos = repartirColumnas(ancho, ncols);

  // Number in bold above the table; title in italics below (APA 7).
  // The bookmark goes on the number: it is the hyperlink target of the cached
  // list-of-tables entry (and of the field Word rebuilds on open), so without
  // it the entry links to nothing.
  hijos.push(new Paragraph({
    spacing: { line: DOBLE, before: 0, after: 0 },
    keepNext: true,
    children: [new Bookmark({
      id: `apa_tbl_${t.indice}`,
      children: [
        new TextRun({ text: "Tabla ", bold: true, font: FUENTE, size: TAM }),
        // A real SEQ field, not the literal number: "TOC \c \"Tabla\"" only
        // collects captions that carry "SEQ Tabla". The cached value keeps the
        // number visible even before Word recalculates the field, and it is
        // bold like the rest of the legend.
        new CampoSecuencia(" SEQ Tabla \\* ARABIC ", String(t.indice)),
      ],
    })],
  }));
  if (t.titulo) {
    hijos.push(new Paragraph({
      spacing: { line: DOBLE, before: 0, after: 0 },
      keepNext: true,
      children: [new TextRun({ text: t.titulo, italics: true, font: FUENTE, size: TAM })],
    }));
  }

  const filas = t.filas;
  const filasDoc = filas.map((fila, i) => {
    const esCabecera = i === 0;
    const esUltima = i === filas.length - 1;
    const celdas = [];
    for (let c = 0; c < ncols; c++) {
      const texto = (fila[c] || "").replace(/\u21b5/g, "\n");
      const par = new Paragraph({
        spacing: { line: 240, before: 20, after: 20 },
        alignment: AlignmentType.LEFT,
        children: texto.split("\n").map((l, k) => new TextRun({
          text: l,
          bold: esCabecera,
          font: FUENTE,
          size: TAM,
          break: k > 0 ? 1 : 0,
        })),
      });
      celdas.push(new TableCell({
        children: [par],
        width: { size: anchos[c], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        // APA 7: no vertical lines. Only the table's top border,
        // the line under the header row and the bottom border.
        borders: {
          top: esCabecera ? B() : SIN_BORDE,
          bottom: (esCabecera || esUltima) ? B() : SIN_BORDE,
          left: SIN_BORDE,
          right: SIN_BORDE,
        },
      }));
    }
    return new TableRow({ children: celdas, tableHeader: esCabecera, cantSplit: false });
  });

  // Bottom border only on the last row: APA 7 does not use vertical lines.
  // (resolved when creating each cell, not by mutating the generated XML)

  hijos.push(crearTabla({
    rows: filasDoc,
    width: { size: ancho, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: anchos,
  }));

  if (t.nota) {
    hijos.push(new Paragraph({
      spacing: { line: 240, before: 0, after: 0 },
      children: [
        new TextRun({ text: "Nota. ", italics: true, font: FUENTE, size: TAM }),
        new TextRun({ text: sinEtiquetaNota(t.nota), font: FUENTE, size: TAM }),
      ],
    }));
  }
  return hijos;
}

// ---------------------------------------------------------------------------
// Figures
// ---------------------------------------------------------------------------

function construirFigura(f) {
  const hijos = [];
  if (!f.existe) {
    log(`    Figure ${f.indice} OMITTED: the file "${f.ruta}" does not exist.`);
    return hijos;
  }

  // The bookmark goes on the number, which is what the list of figures points at.
  hijos.push(new Paragraph({
    spacing: { line: DOBLE, before: 0, after: 0 },
    keepNext: true,
    children: [new Bookmark({
      id: `apa_fig_${f.indice}`,
      children: [
        new TextRun({ text: "Figura ", bold: true, font: FUENTE, size: TAM }),
        new CampoSecuencia(" SEQ Figura \\* ARABIC ", String(f.indice)),
      ],
    })],
  }));
  if (f.titulo) {
    hijos.push(new Paragraph({
      spacing: { line: DOBLE, before: 0, after: 0 },
      keepNext: true,
      children: [new TextRun({ text: f.titulo, italics: true, font: FUENTE, size: TAM })],
    }));
  }

// docx works in pixels at 96 DPI: 1 inch = 96 px.

// Real image size by reading its header. `new ImageRun(...)` is not used as a
// probe because it does NOT expose the dimensions: it returned 1x1 and every
// figure came out square (6.50 x 6.50 in) even when the original was 14.58 x
// 5.00. This way the real width and height are read, without adding dependencies.
function dimensionesImagen(ruta) {
  const b = fs.readFileSync(ruta);
  if (b.length > 24 && b.readUInt32BE(0) === 0x89504e47) {
    return { ancho: b.readUInt32BE(16), alto: b.readUInt32BE(20) };        // PNG
  }
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i += 1; continue; }
      const marcador = b[i + 1];
      // SOF0..SOF15, skipping the non-load ones (DHT/DQT/DRI...).
      if (marcador >= 0xc0 && marcador <= 0xcf &&
          ![0xc4, 0xc8, 0xcc].includes(marcador)) {
        return { ancho: b.readUInt16BE(i + 7), alto: b.readUInt16BE(i + 5) };
      }
      i += 2 + b.readUInt16BE(i + 2);
    }
    return null;
  }
  if (b.length > 10 && b.slice(0, 3).toString("latin1") === "GIF") {
    return { ancho: b.readUInt16LE(6), alto: b.readUInt16LE(8) };
  }
  if (b.length > 26 && b.slice(0, 2).toString("latin1") === "BM") {
    return { ancho: b.readInt32LE(18), alto: Math.abs(b.readInt32LE(22)) };
  }
  return null;
}

  const anchoMaxIn = f.ancho_in || 6.5;
  const dim = dimensionesImagen(f.ruta_absoluta);
  if (!dim || !dim.ancho || !dim.alto) {
    throw new Error(
      `Could not read the size of the image "${f.ruta}". The header of PNG, ` +
      "JPEG, GIF and BMP can be read; if it is another format, convert it first.");
  }
  const ext = path.extname(f.ruta).toLowerCase();
  const tipo = ext === ".jpg" || ext === ".jpeg" ? "jpg" : (ext === ".gif" ? "gif" : "png");

  // The width comes from the manifest (the original's box) and the height from
  // the file, so it does not get distorted; if it still ends up too tall, it is scaled as a whole.
  let anchoPx = Math.round(anchoMaxIn * 96);
  let altoPx = Math.round(anchoPx * (dim.alto / dim.ancho));
  if (f.alto_in) {
    const topePx = Math.round(f.alto_in * 96);
    if (altoPx > topePx) {
      const k = topePx / altoPx;
      anchoPx = Math.round(anchoPx * k);
      altoPx = topePx;
    }
  }
  if (altoPx > 648) {                       // 6.75 in: half of the usable sheet
    const k = 648 / altoPx;
    anchoPx = Math.round(anchoPx * k);
    altoPx = 648;
  }
  // An image is never enlarged beyond its own size in pixels: stretching
  // a 260 px icon to 6.5 inches only produces a blur. The original's box is
  // respected, unless it asks for something bigger than the image provides.
  if (anchoPx > dim.ancho) {
    const k = dim.ancho / anchoPx;
    anchoPx = dim.ancho;
    altoPx = Math.round(altoPx * k);
  }

  // keepNext on the IMAGE, not on the note: the note goes BELOW the image
  // (APA 7), so it is the image that has to be pulled to the next page to keep
  // the two together. With keepNext on the note instead, a note at the top of a
  // page would drag a note away from the figure it belongs to, which is the
  // same defect in the opposite direction.
  hijos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 0, after: 0, lineRule: LineRuleType.AUTO },
    keepNext: !!f.nota,
    children: [new ImageRun({
      type: tipo,
      data: fs.readFileSync(f.ruta_absoluta),
      transformation: { width: anchoPx, height: altoPx },
    })],
  }));

  if (f.nota) {
    hijos.push(new Paragraph({
      spacing: { line: 240, before: 0, after: 0, lineRule: LineRuleType.AUTO },
      children: [
        new TextRun({ text: "Nota. ", italics: true, font: FUENTE, size: TAM }),
        new TextRun({ text: sinEtiquetaNota(f.nota), font: FUENTE, size: TAM }),
      ],
    }));
  }

  return hijos;
}

// ---------------------------------------------------------------------------
// Body: walks the manifest's block stream
// ---------------------------------------------------------------------------

// Returns the body split into segments, instead of a flat list. A segment
// is `{paisaje, hijos, ocupado}` and each segment becomes a section of the
// .docx. Wide tables go in landscape segments and the text goes back to portrait
// afterwards.
//
// `hayContenidoPrevio` says whether something was already written in the section
// before this body (they are the indexes). It is not used to fill the segment,
// but to decide whether a level 1 heading needs a page break: without it, the
// first title would stick to the end of the index.
function construirCuerpoSegmentado(hayContenidoPrevio) {
  const tramos = [];
  const porHid = new Map(M.secciones.map((s) => [s.hid, s]));
  const decisiones = [];

  let actual = { paisaje: false, hijos: [], ocupado: !!hayContenidoPrevio };
  tramos.push(actual);

  // Emit any block in the current segment.
  const emitir = (x) => {
    actual.hijos.push(x);
    actual.ocupado = true;
  };

  // Opens a segment. If the previous one is already landscape it is reused: two
  // wide tables in a row share a landscape page instead of alternating twice.
  const abrirTramo = (paisaje) => {
    const ultimo = tramos[tramos.length - 1];
    if (ultimo.paisaje === paisaje) {
      actual = ultimo;
      return ultimo;
    }
    actual = { paisaje, hijos: [], ocupado: false };
    tramos.push(actual);
    return actual;
  };

  for (const b of M.bloques) {
    switch (b.tipo) {
      case "portada":
      case "portada_dato":
        // The cover page was already built separately, with its own format.
        break;

      case "h": {
        const s = porHid.get(b.hid) || { nivel: b.nivel, texto: b.texto };
        const nivel = s.nivel;
        const nivel1 = nivel === 1;
        // APA 7: n1 centered bold; n2 flush left bold; n3 flush left
        // bold+italics; n4 and n5 with first-line indent, bold (n5
        // also italic) and final period, with the text continuing on the SAME
        // line. Back in the day everything that was not level 1 came out bold
        // and left aligned, so levels 3, 4 and 5 were indistinguishable.
        const nivel3 = nivel === 3;
        const nivelEnLinea = nivel === 4 || nivel === 5;
        const estilo = { bold: true, italics: nivel3 || nivel === 5 };
        // The children are built BEFORE creating the Paragraph: the library does
        // not expose `children` as a mutable array, so pushing afterwards fails.
        const hijosCab = [new Bookmark({
          id: `apa_sec_${b.hid}`,
          children: runsDe(b.segmentos, estilo),
        })];
        if (nivelEnLinea) {
          // Title inline: the final period closes the title and the body of the
          // paragraph continues right after it in the same paragraph.
          hijosCab.push(new TextRun({ text: ". ", font: FUENTE, size: TAM, ...estilo }));
          if (b.en_linea_segmentos) hijosCab.push(...runsDe(b.en_linea_segmentos));
        }
        // APA 7: level 1 starts on a new page. It is marked with
        // `pageBreakBefore` on the heading itself and NOT with a stray
        // paragraph containing a PageBreak: that empty paragraph, when the
        // previous page is already full, does not fit in it, jumps to the next
        // one and leaves a blank page before the title.
        const p = new Paragraph({
          alignment: nivel1 ? AlignmentType.CENTER : AlignmentType.LEFT,
          spacing: { line: DOBLE, before: 0, after: 0 },
          keepNext: !nivelEnLinea,
          pageBreakBefore: nivel1 && actual.ocupado,
          indent: nivelEnLinea ? { firstLine: SANGRIA_1RA } : undefined,
          // The applied outline level is what the content TOC's "\u" collects
          // and what LibreOffice turns into PDF bookmarks (doc.get_toc()), which
          // is how the double pass learns each section's real page.
          outlineLevel: nivel - 1,
          children: hijosCab,
        });
        emitir(p);
        break;
      }

      case "p":
        emitir(new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { line: DOBLE, before: 0, after: 0 },
          indent: { firstLine: SANGRIA_1RA },
          children: runsDe(b.segmentos),
        }));
        break;

      case "lista":
        emitir(new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { line: DOBLE, before: 0, after: 0 },
          indent: { firstLine: SANGRIA_1RA },
          children: runsDe(b.segmentos),
        }));
        break;

      case "tabla": {
        const t = M.tablas[b.indice];
        if (!t) break;
        const ancho = evaluarAncho(t);
        decisiones.push({ indice: t.indice, ...ancho });
        if (ancho.ancha) {
          // The table goes on its own landscape page, with its title and its note.
          abrirTramo(true);
          actual.hijos.push(...construirTabla(t, ANCHO_CONTENIDO_H));
          actual.ocupado = true;
          abrirTramo(false);
        } else {
          actual.hijos.push(...construirTabla(t, ANCHO_CONTENIDO));
          actual.ocupado = true;
        }
        break;
      }

      case "figura": {
        const f = M.figuras[b.indice];
        if (f) {
          actual.hijos.push(...construirFigura(f));
          actual.ocupado = true;
        }
        break;
      }

      // The parser already attached this paragraph to its table or figure
      // (caption or a "Nota." line written in the .md): it is printed by
      // construirTabla()/construirFigura(), never as body text.
      case "nota_tabla":
      case "nota_figura":
        break;

      default:
        log(`    [warning] unknown block type: "${b.tipo}". It is skipped.`);
    }
  }

  // The segments that ended up empty (typically the portrait one before and after
  // a wide table) are not used because they would create blank pages.
  const conContenido = tramos.filter((t) => t.hijos.length > 0);

  for (const d of decisiones) {
    if (d.ancha) {
      log(`  Table ${d.indice}: landscape page -> ${d.motivo}.`);
    } else {
      log(`  Table ${d.indice}: portrait (${d.ncols} columns, cell of up to ${d.peorVertical} lines).`);
    }
  }
  const apaisadas = decisiones.filter((d) => d.ancha).length;
  log(`  Orientation: ${apaisadas} landscape table(s) out of ${decisiones.length}.`);

  return conContenido;
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

function construirReferencias() {
  const hijos = [];
  const refs = (M.referencias || []).slice();
  // APA 7 requires alphabetical order. The .md may not come sorted.
  refs.sort((a, b) => a.texto.localeCompare(b.texto, "es", { sensitivity: "base" }));

  log(`  References: ${refs.length} (sorted alphabetically in the .docx).`);
  for (const r of refs) {
    if (r.cursiva_propuesta) {
      log(`    italics to review: "${r.cursiva_propuesta.slice(0, 60)}"`);
    }
    hijos.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: DOBLE, before: 0, after: 0 },
      // Hanging indent: 0.5" to the left and -0.5" first line.
      indent: { left: SANGRIA_1RA, hanging: SANGRIA_1RA },
      children: runsDeReferencia(r),
    }));
  }
  return hijos;
}

// ---------------------------------------------------------------------------
// Assembly
// ---------------------------------------------------------------------------

async function main() {
  console.log("=== PHASE 2: building the .docx ===");
  log(`Manifest   : ${args.manifiesto}`);
  log(`Output     : ${DOCX}`);
  log(`Source .md : ${M.fuente ? M.fuente.md : "(unknown)"}`);
  const p = M.portada || {};
  log(`Cover      : title=${p.titulo ? "yes" : "NO"} authors=${(p.autores || []).length} instructor=${p.docente ? "yes" : "NO"} date=${p.fecha ? "yes" : "NO"}`);
  if (p.campos_faltantes && p.campos_faltantes.length) {
    log(`Cover      : MISSING FIELDS -> ${p.campos_faltantes.join(", ")}`);
    log("            (the cover page is generated anyway; warn the user about what is missing)");
  }
  log(`Tables     : ${M.tablas.length}   Figures: ${M.figuras.length}   References: ${M.referencias.length}`);
  log(`Indexes    : tables=${M.opciones.indice_tablas ? "yes" : "no"} figures=${M.opciones.indice_figuras ? "yes" : "no"}`);
  log("");

  const indices = construirIndices();
  // The indexes stay in the SAME section as the body's first portrait segment.
  // If they were separated, the section break would add a blank page between the
  // index and the text, which did not exist before.
  const tramos = construirCuerpoSegmentado(indices.length > 0);
  const primerTramo = tramos[0];
  if (primerTramo && !primerTramo.paisaje) {
    primerTramo.hijos = [...indices, ...primerTramo.hijos];
  }
  const refs = construirReferencias();

  const propsVerticales = {
    page: {
      size: { width: TIPO_PAGINA.width, height: TIPO_PAGINA.height, orientation: PageOrientation.PORTRAIT },
      margin: { top: MARGEN, right: MARGEN, bottom: MARGEN, left: MARGEN, header: 720, footer: 720 },
    },
  };
  // `docx` swaps width and height on its own when it sees
  // orientation=LANDSCAPE. It has to be given the measurements IN PORTRAIT
  // (12240 x 15840) and the orientation marked: passing them already inverted
  // produces a `w:pgSz w:w="12240" w:h="15840" w:orient="landscape"`, which
  // Word and LibreOffice read as portrait and the table comes out cramped.
  const propsApaisadas = {
    page: {
      size: { width: TIPO_PAGINA.width, height: TIPO_PAGINA.height, orientation: PageOrientation.LANDSCAPE },
      margin: { top: MARGEN, right: MARGEN, bottom: MARGEN, left: MARGEN, header: 720, footer: 720 },
    },
  };

  const secciones = [
    // Section 1: cover page alone, WITHOUT page number.
    {
      properties: { ...propsVerticales, type: SectionType.NEXT_PAGE },
      headers: { default: headerVacio },
      footers: { default: footerVacio },
      children: construirPortada(),
    },
  ];

  // If the first segment is landscape, the indexes need their own portrait
  // section; if not, they are already inside that segment and are not duplicated.
  if (indices.length && (!primerTramo || primerTramo.paisaje)) {
    secciones.push({
      properties: { ...propsVerticales, type: SectionType.NEXT_PAGE },
      headers: { default: headerConPagina },
      footers: { default: footerVacio },
      children: indices,
    });
  }

  // One section per body segment. Each segment opens a new page, so a landscape
  // table stays on its own page and the text afterwards goes back to portrait.
  // All of them repeat header and footer so the numbering is not lost.
  for (const tramo of tramos) {
    secciones.push({
      properties: { ...(tramo.paisaje ? propsApaisadas : propsVerticales), type: SectionType.NEXT_PAGE },
      headers: { default: headerConPagina },
      footers: { default: footerVacio },
      children: tramo.hijos,
    });
  }

  // Final section: references, always in portrait.
  //
  // The "Referencias Bibliograficas" heading does NOT come from here: it is a
  // level 1 heading of the .md and already ended up at the end of the last body
  // segment. With a page break it was left alone on a page and the entries began
  // on the next one. If the previous segment is already portrait a CONTINUOUS
  // section is used, which does not break the page: the heading drops along with
  // the first entry. Only if the document ends on a landscape table is a new page
  // needed, because it has to go back to portrait.
  if (refs.length) {
    const ultimoTramo = tramos[tramos.length - 1];
    const haceFaltaSalto = !ultimoTramo || ultimoTramo.paisaje;
    if (haceFaltaSalto) {
      if (ultimoTramo) ultimoTramo.hijos = [...ultimoTramo.hijos, ...refs];
      else {
        secciones.push({
          properties: { ...propsVerticales, type: SectionType.NEXT_PAGE },
          headers: { default: headerConPagina },
          footers: { default: footerVacio },
          children: refs,
        });
      }
    } else {
      secciones.push({
        properties: { ...propsVerticales, type: SectionType.CONTINUOUS },
        headers: { default: headerConPagina },
        footers: { default: footerVacio },
        children: refs,
      });
    }
  }

  const doc = new Document({
    creator: "generate-apa-document",
    // Asks the consumer (Word) to recalculate every field when the file is
    // opened. The indexes ship with a cached result (the pages measured in the
    // second pass), so this is a convenience for Word users, not what makes the
    // PDF correct: LibreOffice exports the cached result as-is.
    features: { updateFields: true },
    title: p.titulo || "Documento APA 7",
    description: `Generado desde ${M.fuente ? M.fuente.md : "fuente .md"}`,
    styles: {
      default: {
        document: { run: { font: FUENTE, size: TAM }, paragraph: { spacing: { line: DOBLE, before: 0, after: 0 } } },
      },
      // The cached TOC entries are emitted with the built-in TOC1..TOC5 styles.
      // Defining them here keeps the indexes in APA type (Times 12, double
      // spaced, entries NOT bold even when the section title is: see
      // references/apa7-format.md) instead of whatever the renderer falls
      // back to when the style is missing.
      paragraphStyles: Array.from({ length: 5 }, (_, i) => ({
        id: `TOC${i + 1}`,
        name: `TOC ${i + 1}`,
        quickFormat: true,
        run: { font: FUENTE, size: TAM },
        paragraph: {
          spacing: { line: DOBLE, before: 0, after: 0 },
          indent: { left: i * SANGRIA_1RA },
        },
      })),
      // TOC entries are hyperlinks to the headings. The default hyperlink style
      // is blue and underlined, which APA does not use, so the index keeps black
      // text and no underline.
      characterStyles: [{
        id: "IndexLink",
        name: "Index Link",
        run: { font: FUENTE, size: TAM, color: "000000" },
      }],
    },
    sections: secciones,
  });

  const buffer = await Packer.toBuffer(doc);
  const dir = path.dirname(path.resolve(DOCX));
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(DOCX, buffer);

  log("");
  log(`.docx written: ${DOCX} (${buffer.length} bytes)`);
  if (Object.keys(PAGINAS.secciones || {}).length || Object.keys(PAGINAS.tablas || {}).length
      || Object.keys(PAGINAS.figuras || {}).length) {
    log("Indexes built as real TOC fields with cached page numbers (--paginas-json).");
    log("Word recalculates them on open; LibreOffice/Word can update them in place.");
  } else {
    log("Indexes built as real TOC fields; page numbers are left for the renderer.");
    log("apa7.py resolves them by exporting once and rebuilding with --paginas-json.");
  }

  if (args.log) {
    const ld = path.dirname(path.resolve(args.log));
    if (ld && !fs.existsSync(ld)) fs.mkdirSync(ld, { recursive: true });
    fs.writeFileSync(args.log, logLines.join("\n") + "\n", "utf8");
    log(`Log written : ${args.log}`);
  }
}

main().catch((e) => {
  console.error("FAILURE in build-docx.js:", e && e.stack ? e.stack : e);
  process.exit(1);
});
