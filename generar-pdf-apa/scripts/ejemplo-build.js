const fs = require('fs');
const path = require('path');
const d = require('./node_modules/docx');

const MD_PATH = 'C:/Users/User/Downloads/Docling/salida_markdown/Informe_Comparativo_SDLC/hybrid_auto/Informe_Comparativo_SDLC.md';
const OUT_DIR = 'C:/Users/User/Downloads/Docling/salida_markdown/Informe_Comparativo_SDLC/hybrid_auto/Informe_Comparativo_SDLC_APA';
const OUT_DOCX = path.join(OUT_DIR, 'Informe_Comparativo_SDLC_APA.docx');

const md = fs.readFileSync(MD_PATH, 'utf-8');
const lines = md.split(/\r?\n/);

function norm(s) {
  return s.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

// ---------- parse markdown ----------
const tables = [];   // {rows:[[cell,...],...]}
let inRefs = false;
let seenHeading = false;
const blocks = [];   // {t:'h1'|'h2'|'p', text} | {t:'table', idx}

// ---------- heading resolution ----------
const H1_PATTERNS = [
  [/^1 introduccion/, 'Introducción'],
  [/^2 analisis de microsoft/, 'Análisis de Microsoft SDL'],
  [/^3 analisis de owasp/, 'Análisis de OWASP SAMM'],
  [/^4 analisis de bsimm/, 'Análisis de BSIMM'],
  [/^5 analisis de building/, 'Análisis de Building Security In (BSI)'],
  [/^6 analisis de iso/, 'Análisis de ISO/IEC 27002:2022 y Desarrollo Seguro'],
  [/^7 cuadro comparativo/, 'Cuadro Comparativo de Metodologías'],
  [/^8 analisis de aplicabilidad/, 'Análisis de aplicabilidad y selección del marco'],
  [/^9 actividades en grupo/, 'Actividades en grupo'],
  [/^referencias bibliograficas/, 'Referencias Bibliográficas'],
];
const H2_PATTERNS = [
  [/^actividad 4/, 'ACTIVIDAD 4: Plataforma de Expediente Médico Digital (MediCloud)'],
  [/^actividad 5/, 'ACTIVIDAD 5: FinTech SecureApp'],
  [/^1 segun/, 'Según el Marco BSIMM'],
  [/^2 propone/, 'Propone 3 Actividades BSIMM prioritarias'],
  [/^argumentacion/, 'Argumentación Técnica y de Riesgo (El "Porqué" crítico)'],
  [/^actividad sm1 1/, 'Actividad SM1.1 o AM1.2: Establecer un Grupo de Seguridad de Software (SSG) y definir roles y objetivos'],
  [/^actividad cr1 4/, 'Actividad CR1.4: Integrar herramientas de análisis estático de código (SAST) en el pipeline de CI/CD'],
  [/^actividad t1 1/, 'Actividad T1.1: Iniciar un plan de concientización y entrenamiento técnico en seguridad para desarrolladores'],
  [/^medidas tecnicas/, 'Medidas Técnicas Recomendadas'],
  [/^objetivos del informe/, 'Objetivos del Informe'],
  [/^estructura detallada/, 'Estructura Detallada y Prácticas Clave'],
  [/^fortalezas/, 'Fortalezas, Limitaciones y Aplicabilidad'],
  [/^las 5 funciones/, 'Las 5 Funciones de Negocio y sus Prácticas'],
  [/^ventajas/, 'Ventajas, Limitaciones y Aplicabilidad'],
  [/^los 4 dominios/, 'Los 4 Dominios y el Análisis Empírico (Spider Chart)'],
  [/^pilares/, 'Pilares Fundamentales y los 7 Touchpoints'],
  [/^la estructura de 4 temas/, 'La Estructura de 4 Temas y el Enfoque en Software'],
];

function resolveHeading(rawNorm, text) {
  for (const [re, out] of H1_PATTERNS) {
    if (re.test(rawNorm)) return { level: 1, text: out };
  }
  for (const [re, out] of H2_PATTERNS) {
    if (re.test(rawNorm)) return { level: 2, text: out };
  }
  let out = text.replace(/^[•\-\d]+\s*/, '').replace(/^\.\s*/, '');
  out = out.replace(/:\s*$/, '').replace(/\s{2,}/g, ' ').trim();
  return { level: 2, text: out };
}

// ---------- parse markdown into blocks ----------
for (const raw of lines) {
  const line = raw.trimEnd();
  const t = line.trimStart();
  if (!t) continue;

  const n = norm(t);

  // references section: everything after the referencias heading
  if (inRefs) continue;
  if (n === 'referencias bibliograficas' || n.startsWith('10 referencias')) {
    inRefs = true;
    blocks.push({ t: 'h1', text: 'Referencias Bibliográficas' });
    continue;
  }

  if (t.startsWith('# ')) continue; // document title -> cover

  if (t.startsWith('## ')) {
    seenHeading = true;
    const { level, text } = resolveHeading(n, t.slice(3).trim());
    blocks.push({ t: level === 1 ? 'h1' : 'h2', text });
    continue;
  }
  if (!seenHeading) continue; // skip front-matter block (cover fields duplicated here)

  if (t.startsWith('<table>')) {
    const rows = [];
    const trs = t.match(/<tr>([\s\S]*?)<\/tr>/g) || [];
    for (const tr of trs) {
      const tds = tr.match(/<td>([\s\S]*?)<\/td>/g) || [];
      const cells = tds.map(td => {
        let c = td.replace(/^<td>/, '').replace(/<\/td>$/, '');
        c = c.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
        c = c.replace(/\s+/g, ' ').trim();
        c = c.replace(/\s+DOCX$/i, '');
        return c;
      });
      if (cells.length) rows.push(cells);
    }
    if (rows.length) {
      const idx = tables.length;
      tables.push({ rows });
      blocks.push({ t: 'table', idx });
    }
    continue;
  }

  // body paragraph / list
  blocks.push({ t: 'p', text: t });
}

// ---------- heading index (para TOC con PAGEREF + bookmarks) ----------
let hid = 0;
const headings = [];
for (const b of blocks) {
  if (b.t === 'h1' || b.t === 'h2') {
    b.hid = ++hid;
    headings.push({ hid, level: b.t === 'h1' ? 1 : 2, text: b.text });
  }
}

// ---------- build content children ----------
const TABLE_TITLES = [
  'Matriz comparativa de metodologías S-SDLC',
  'Matriz argumentativa de controles ISO/IEC 27002 – FintechX',
  'Matriz argumentativa de prácticas de Microsoft SDL – Alcaldía Digital',
  'Evaluación de prácticas OWASP SAMM – EcoPay',
  'Casos de abuso – MediCloud (Ejercicio 1)',
  'Bugs y flaws – MediCloud (Ejercicio 2)',
  'Estrategia de touchpoints BSI – MediCloud (Ejercicio 3)',
];

const run = (text, opts = {}) => new d.TextRun(Object.assign({ text }, opts));
const para = (opts, children) => new d.Paragraph(Object.assign({}, opts, { children }));

function bodyParagraph(text) {
  // normalize doubled list numbering artifacts: "1. 1." -> "1. "
  text = text.replace(/^\s*(\d+)\.\s*\d+\.\s/, '$1. ');
  // collapse repeated leading bullet markers ("•  •", "– •", "• ••") to a single bullet
  if (/^\s*[•*\u2013\u2014-]\s*[•*\u2013\u2014-]/.test(text)) {
    text = '• ' + text.replace(/^\s*[•*\u2013\u2014-\s]+/, '');
    // collapse doubled numbering after the bullet ("• 1. 1. text" -> "• 1. text")
    text = text.replace(/^(•\s*)?(\d+)\.\s*\d+\.\s/, '$1$2. ');
  }
  if (/^•/.test(text)) {
    return para({ indent: { left: 720 }, spacing: { before: 0, after: 0, line: 480 }, alignment: d.AlignmentType.LEFT },
      [run(text)]);
  }
  if (/^\d+\.\s/.test(text)) {
    return para({ indent: { left: 720 }, spacing: { before: 0, after: 0, line: 480 }, alignment: d.AlignmentType.LEFT },
      [run(text)]);
  }
  return para({ indent: { firstLine: 720 }, spacing: { before: 0, after: 0, line: 480 }, alignment: d.AlignmentType.LEFT },
    [run(text)]);
}

function captionBlocks(tableIdx) {
  const title = TABLE_TITLES[tableIdx] || `Matriz ${tableIdx + 1}`;
  const num = String(tableIdx + 1);
  return [
    para({ spacing: { before: 240, after: 240, line: 480 }, alignment: d.AlignmentType.LEFT },
      [
        new d.Bookmark({ id: '_Tabla' + num, children: [
          new d.TextRun({ bold: true, children: ['Tabla ' + num] }),
        ] }),
        new d.TextRun({ break: 1 }),
        run(title, { italics: true }),
      ]),
  ];
}

function noteBlock() {
  return para({ spacing: { before: 120, after: 240, line: 480 }, alignment: d.AlignmentType.LEFT },
    [run('Nota. ', { italics: true }), run('Elaboración propia.')]);
}

const border = { style: d.BorderStyle.SINGLE, size: 4, color: '000000' };
const noBorder = { style: d.BorderStyle.NONE, size: 0, color: 'auto' };
function makeTable(tableIdx) {
  const { rows } = tables[tableIdx];
  const head = rows[0];
  const usable = tableIdx === 0 ? 12960 : 9360;   // ancho util: horizontal (paisaje) vs vertical (carta)
  const cols = head.length || 1;
  const firstCol = tableIdx === 0 ? 2160 : 1560;
  const rest = Math.max(0, usable - firstCol);
  const columnWidths = cols === 1 ? [usable] : (cols === 2 ? [firstCol, rest] : [firstCol, ...Array(cols - 1).fill(Math.round(rest / (cols - 1)))]);
  const rowsOut = rows.map((cells, ri) => {
    const cellsOut = cells.map(cellText => {
      return new d.TableCell({
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        shading: ri === 0 ? { fill: 'D9E2F3' } : undefined,
        children: [para({ alignment: d.AlignmentType.LEFT },
          [run(cellText, { bold: ri === 0, size: 20 })]
        )],
      });
    });
    // pad missing cells
    while (cellsOut.length < head.length) {
      cellsOut.push(new d.TableCell({ children: [para({}, [run('')])] }));
    }
    return new d.TableRow({ tableHeader: ri === 0, children: cellsOut });
  });
  return new d.Table({
    rows: rowsOut,
    width: { size: usable, type: d.WidthType.DXA },
    columnWidths,
    layout: d.TableLayoutType.FIXED,
    borders: { top: border, bottom: border, left: noBorder, right: noBorder, insideHorizontal: border, insideVertical: noBorder },
  });
}

// ---------- section assembly ----------
const PAGE_PORTRAIT = { width: 12240, height: 15840 };
const PAGE_LANDSCAPE = { width: 15840, height: 12240 };
const MARGIN = 1440;

function headerFor() {
  return new d.Header({
    children: [new d.Paragraph({
      alignment: d.AlignmentType.RIGHT,
      children: [new d.TextRun({ children: [d.PageNumber.CURRENT], size: 24 })],
    })],
  });
}

// ---- portada: layout posicional (referencia: PDF ARQUITECTURA_X_SEGMED_APA7_corregido.pdf) ----
// Título arriba (y≈100pt), autores al centro (y≈293pt), resto anclado abajo (y≈514pt+),
// pitch de línea del bloque inferior ≈27.6pt. Todas las distancias en puntos (1pt = 20 twips).
// LOGO_PATH es opcional: si está vacío NO se inserta logo (valor por defecto).
const COVER_LOGO_PATH = '';              // ruta del logo si el usuario lo aporta ('' = sin logo)
const COVER_LOGO_MAX_W_IN = 2;           // ancho máximo del logo en pulgadas (96 dpi → 1in = 96px)
const COVER_TITLE_BEFORE_PT = 28;        // empuja el título a ~y100 (margen 72 + before)
const COVER_AUTHORS_BEFORE_PT = 165;     // desde el final del título hasta autores (~293)
const COVER_BOTTOM_BEFORE_PT = 193;      // desde autores hasta facultad (~514)
const COVER_PITCH_PT = 0.1;              // before extra para pitch ≈27.6pt entre líneas del bloque

// Lee dimensiones reales (px) de PNG o JPEG leyendo la cabecera (sin dependencias).
function imageSizePx(filePath) {
  const buf = fs.readFileSync(filePath);
  if (buf[0] === 0x89 && buf[1] === 0x50) {            // PNG: IHDR en bytes 16-24
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
  }
  if (buf[0] === 0xFF && buf[1] === 0xD8) {            // JPEG: recorrer marcadores hasta SOF
    let off = 2;
    while (off < buf.length) {
      if (buf[off] !== 0xFF) { off += 1; continue; }
      const marker = buf[off + 1];
      if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
        const h = buf.readUInt16BE(off + 5);
        const w = buf.readUInt16BE(off + 7);
        return { width: w, height: h };
      }
      const len = buf.readUInt16BE(off + 2);
      off += 2 + len;
    }
  }
  return { width: 1, height: 1 };
}

// autores en un solo párrafo centrado: coma + "y" antes del último (estándar portada)
const authorsPara = (names) => {
  const text = names.length > 1
    ? names.slice(0, -1).join(', ') + ' y ' + names[names.length - 1]
    : (names[0] || '');
  return [[[text, {}]]];
};

// Párrafo de portada centrado con before explícito (twips) para posicionar el bloque.
function coverPara(beforePt, children, extra = {}) {
  return para(Object.assign({ alignment: d.AlignmentType.CENTER, spacing: { line: 480, before: Math.round(beforePt * 20), after: 0 } }, extra), children);
}

function coverChildren() {
  const children = [];

  // Logo opcional: solo si se aporta (COVER_LOGO_PATH). Centrado arriba del título.
  if (COVER_LOGO_PATH && fs.existsSync(COVER_LOGO_PATH)) {
    const { width, height } = imageSizePx(COVER_LOGO_PATH);
    const wPx = Math.round(COVER_LOGO_MAX_W_IN * 96);
    const hPx = Math.max(1, Math.round(wPx * (height / width)));
    children.push(coverPara(4, [
      new d.ImageRun({ type: 'png', data: fs.readFileSync(COVER_LOGO_PATH), transformation: { width: wPx, height: hPx } }),
    ]));
  }

  // ZONA ARRIBA: título.
  children.push(coverPara(COVER_TITLE_BEFORE_PT, [run('Informe Comparativo de Metodologías S-SDLC', { bold: true })]));

  // ZONA CENTRO: autores (único elemento al medio de la página).
  children.push(coverPara(COVER_AUTHORS_BEFORE_PT, [run(authorsPara(['Nombre Apellido'])[0][0][0], {})]));

  // ZONA ABAJO: bloque institucional anclado abajo con pitch ≈27.6pt entre líneas.
  children.push(coverPara(COVER_BOTTOM_BEFORE_PT, [run('Facultad de Ingeniería de Sistemas, Corporación Universitaria Minuto de Dios', {})]));
  children.push(coverPara(COVER_PITCH_PT, [run('Vicerrectoría Regional Centro Sur, Sede Ibagué, Tolima', {})]));
  children.push(coverPara(COVER_PITCH_PT, [run('Seguridad de la información NRC 95405', {})]));
  children.push(coverPara(COVER_PITCH_PT, [run('Romulo Betancourt Hortua', {})]));
  children.push(coverPara(COVER_PITCH_PT, [run('15 de septiembre de 2026', {})]));

  return children;
}

// split body blocks into portrait / landscape / portrait around Table idx 0
const secBody1 = [], secLandscape = [], secBody2 = [];
for (const b of blocks) {
  if (b.t === 'table' && b.idx === 0) { secLandscape.push(b); continue; }
  if (secLandscape.length === 0) secBody1.push(b); else secBody2.push(b);
}

const bodyChildren = (list) => list.map(b => {
  switch (b.t) {
    case 'h1': return para({ heading: d.HeadingLevel.HEADING_1, pageBreakBefore: b.text === 'Referencias Bibliográficas', spacing: { before: 240, after: 240, line: 480 }, alignment: d.AlignmentType.CENTER }, [new d.Bookmark({ id: '_H' + b.hid, children: [new d.TextRun({ text: b.text, bold: true, color: '000000', font: 'Times New Roman', size: 24 })] })]);
    case 'h2': return para({ heading: d.HeadingLevel.HEADING_2, spacing: { before: 240, after: 120, line: 480 }, alignment: d.AlignmentType.LEFT }, [new d.Bookmark({ id: '_H' + b.hid, children: [new d.TextRun({ text: b.text, bold: true, color: '000000', font: 'Times New Roman', size: 24 })] })]);
    case 'table': return [...captionBlocks(b.idx), makeTable(b.idx), noteBlock()];
    case 'p': return bodyParagraph(b.text);
  }
}).flat();

// Build references with italics ranges via segments
function buildReferences() {
  const defs = [
    { segs: [['Colsubsidio, Gerencia Corporativa. (2024). ', {}], ['Políticas de seguridad de la información basadas en BSIMM (Building Security In Maturity Model).', { italics: true }], [' https://cms.colsubsidio.com/sites/default/files/Documentos/colsubsidio/politicas-de-seguridad-de-la-informacion-gerencia-corporativa-de-tecnologia.pdf', {}]] },
    { segs: [['Fundación OWASP. (2021). ', {}], ['Cómo iniciar un programa de AppSec con el OWASP Top 10 y SAMM.', { italics: true }], [' https://owasp.org/Top10/2021/es/A00_2021-How_to_start_an_AppSec_program_with_the_OWASP_Top_10/', {}]] },
    { segs: [['ISO/IEC. (2022). ', {}], ['Estructura y controles de la norma ISO/IEC 27002:2022 y desarrollo seguro.', { italics: true }], [' https://repositorio.utn.edu.ec/bitstream/123456789/15944/2/PG 1823 TRABAJO DE GRADO.pdf', {}]] },
    { segs: [['McGraw, G. (2026). ', {}], ['Seguridad de software: Building Security In. ', {}], ['Revista Más Seguridad', { italics: true }], [' (164). https://www.revistamasseguridad.com.mx/wp-content/uploads/2026/08/MAS-SEGURIDAD-MAGAZINE-164.pdf', {}]] },
    { segs: [['Microsoft Corporation. (2025). ', {}], ['Microsoft Security Development Lifecycle (SDL).', { italics: true }], [' https://learn.microsoft.com/es-es/compliance/assurance/assurance-microsoft-securitydevelopment-lifecycle', {}]] },
  ];
  return [
    ...defs.map(r => para({ indent: { left: 720, hanging: 720 }, spacing: { before: 0, after: 0, line: 480 }, alignment: d.AlignmentType.LEFT },
      r.segs.map(([t, o]) => run(t, o))
    )),
  ];
}

const TAB_INDEX_STOPS = [{ type: d.TabStopType.RIGHT, position: 9350, leader: d.LeaderType.DOT }];
function indexTablesChildren() {
  return TABLE_TITLES.map((title, i) => {
    const n = String(i + 1);
    return para({ spacing: { before: 0, after: 0, line: 480 }, tabStops: TAB_INDEX_STOPS, alignment: d.AlignmentType.LEFT },
      [
        new d.TextRun({ children: ['Tabla ' + n, '  ', title] }),
        new d.TextRun({ children: [new d.Tab()] }),
        new d.SimpleField('PAGEREF _Tabla' + n + ' \\h'),
      ]);
  });
}

function tocEntry(entry) {
  return para({ style: entry.level === 1 ? 'TOC1' : 'TOC2', spacing: { before: 0, after: 0, line: 480 }, tabStops: TAB_INDEX_STOPS, alignment: d.AlignmentType.LEFT },
    [
      run(entry.text),
      new d.TextRun({ children: [new d.Tab()] }),
      new d.SimpleField('PAGEREF _H' + entry.hid + ' \\h'),
    ]);
}
function tocEntries() {
  return headings.map(tocEntry);
}

// TOC section children (índice de tablas solo si el documento tiene tablas)
function tocChildren() {
  const idxTables = tables.length ? [
    para({ alignment: d.AlignmentType.CENTER, pageBreakBefore: true, spacing: { before: 0, after: 240, line: 480 } }, [run('Índice de tablas', { bold: true, size: 24, color: '000000' })]),
    ...indexTablesChildren(),
  ] : [];
  return [
    para({ alignment: d.AlignmentType.CENTER, spacing: { before: 0, after: 240, line: 480 } }, [run('Tabla de contenido', { bold: true, size: 24, color: '000000' })]),
    ...tocEntries(),
    ...idxTables,
  ];
}

const mkSection = (pageSize, children) => ({
  properties: {
    page: { size: pageSize, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } },
  },
  headers: { default: headerFor() },
  children,
});

const doc = new d.Document({
  creator: 'Nombre Apellido',
  description: 'Informe comparativo de metodologías S-SDLC',
  features: { updateFields: true },
  styles: {
    default: {
      document: { run: { font: 'Times New Roman', size: 24 } },
    },
    paragraphStyles: [
      { id: 'TOC1', name: 'TOC 1', basedOn: 'Normal', run: { font: 'Times New Roman', size: 24, bold: false, color: '000000' }, paragraph: { spacing: { line: 480 } } },
      { id: 'TOC2', name: 'TOC 2', basedOn: 'Normal', run: { font: 'Times New Roman', size: 24, bold: false, color: '000000' }, paragraph: { spacing: { line: 480 } } },
    ],
  },
  sections: [
    mkSection(PAGE_PORTRAIT, coverChildren()),
    mkSection(PAGE_PORTRAIT, tocChildren()),
    mkSection(PAGE_PORTRAIT, bodyChildren(secBody1)),
    mkSection(PAGE_LANDSCAPE, bodyChildren(secLandscape)),
    mkSection(PAGE_PORTRAIT, [...bodyChildren(secBody2), ...buildReferences()]),
  ],
});

fs.mkdirSync(OUT_DIR, { recursive: true });
d.Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT_DOCX, buf);
  console.log('OK ->', OUT_DOCX, buf.length, 'bytes');
}).catch(err => { console.error(err); process.exit(1); });