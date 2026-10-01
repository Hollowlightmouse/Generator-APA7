#!/usr/bin/env node
/**
 * build-docx.js - Aplica un MANIFIEST.json y produce el .docx en formato APA 7.
 *
 * Este archivo NO contiene ningun contenido de ningun documento. Antes cada
 * trabajo obligaba a editar un generador de cientos de lineas con el titulo,
 * la portada, las tablas y las 5 referencias escritas a mano (y las
 * referencias se emitian SIEMPRE, aunque el .md tuviera otras). Aqui todo
 * llega desde MANIFIEST.json, que produce scripts\md-a-manifiesto.py.
 *
 * Reglas APA 7 que aplica (references\normas-apa7.md es la fuente):
 *   - Times New Roman 12, doble espacio, sin espacio extra antes/despues.
 *   - Sangria de primera linea 0,5" en parrafos de cuerpo.
 *   - Texto justificado a la IZQUIERDA (no justificado a ambos margenes).
 *   - Nivel 1 centrado y en negrita; nivel 2 al margen izquierdo y en negrita.
 *   - Numero de pagina en la ESQUINA SUPERIOR DERECHA (encabezado, no pie).
 *   - La portada NO lleva numero de pagina; la numeracion visible arranca en
 *     la pagina 2, que es lo que exige APA 7.
 *   - Tablas y figuras: numero en negrita, titulo en cursiva, despues el
 *     objeto. Bordes solo horizontales en tabla.
 *   - Referencias: sangria francesa de 0,5", orden alfabetico.
 *
 * Uso:
 *   node build-docx.js --manifiesto MANIFIEST.json --out salida.docx
 *                      [--log salida\_logs\02-docx.log]
 */

"use strict";

const fs = require("fs");
const path = require("path");

/**
 * Node resuelve `require` desde la carpeta del SCRIPT, no desde la carpeta
 * actual. Como `docx` se instala en <skill>\.work y este script vive en
 * <skill>\scripts, un `require("docx")` normal falla siempre.
 *
 * Se resuelve explicitamente la ruta del modulo en vez de obligar al usuario a
 * copiar el .js dentro de .work o a instalar docx de forma global.
 */
const RAIZ_SKILL = path.resolve(__dirname, "..");
const WORKDIR = process.env.APA7_WORKDIR
  ? path.resolve(process.env.APA7_WORKDIR)
  : path.join(RAIZ_SKILL, ".work");
const RUTA_DOCX = path.join(WORKDIR, "node_modules", "docx");

if (!fs.existsSync(RUTA_DOCX)) {
  console.error("FALLO: no se encuentra el paquete 'docx' en");
  console.error("       " + RUTA_DOCX);
  console.error("Ejecute scripts\\instalar-entorno.ps1 para instalarlo en el workdir de la skill.");
  process.exit(3);
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Bookmark, SimpleField, PageBreak, Footer, Header, PageNumber, AlignmentType,
  TabStopType, LeaderType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  SectionType, PageOrientation, TableLayoutType, HeightRule, convertInchesToTwip,
} = require(RUTA_DOCX);

// ---------------------------------------------------------------------------
// Parametros
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
  console.error("Uso: node build-docx.js --manifiesto MANIFEST.json --out salida.docx [--log log.txt]");
  process.exit(2);
}

const logLines = [];
function log(msg) {
  const line = `  ${msg}`;
  logLines.push(line);
  console.log(line);
}

// readFileSync con "utf8" NO quita el BOM y JSON.parse aborta si lo hay:
// el parser escribe el manifiesto sin BOM, pero si el usuario lo abre y lo
// guarda desde un editor de Windows, vuelve con BOM.
const M = JSON.parse(fs.readFileSync(args.manifiesto, "utf8").replace(/^\uFEFF/, ""));
const DOCX = args.out;

// --- Validacion critica del manifiesto (abortar, no adivinar) ---------------
// Antes el documento se entregaba con los titulos de tablas y figuras en
// blanco: el indice imprimia "(sin titulo)" y el cuerpo se saltaba el titulo
// entero, sin avisar en ninguna parte. Una tabla o figura APA sin titulo es un
// defecto de forma, no una preferencia, asi que aqui se corta el proceso.
const faltantes = [];
for (const t of M.tablas || []) {
  if (!t.titulo || !String(t.titulo).trim()) {
    faltantes.push(`Tabla ${t.indice}: sin titulo`);
  }
}
for (const f of M.figuras || []) {
  if (!f.titulo || !String(f.titulo).trim()) {
    faltantes.push(`Figura ${f.indice}: sin leyenda`);
  }
}
if (faltantes.length) {
  console.error("ERROR: el manifiesto tiene objetos de APA sin titulo. No se genera el .docx.");
  for (const m of faltantes) console.error(`  - ${m}`);
  console.error("");
  console.error("Resuelvelas en el parser (no en este script): el .md, un");
  console.error("--titulos-tabla.json / --titulos-figura-json, o preguntando al usuario.");
  console.error("Consulta M.diagnostico.preguntas y M.diagnostico.pendientes_bloqueantes.");
  process.exit(4);
}

// ---------------------------------------------------------------------------
// Constantes de formato
// ---------------------------------------------------------------------------

const FUENTE = "Times New Roman";
const TAM = 24;                       // 12 pt = 24 half-points
const DOBLE = 480;                    // doble espacio (twips: 240 = simple)
const SANGRIA_1RA = 720;              // 0,5 pulgada
const ANCHO_CONTENIDO = 9360;         // 8,5" - 1" - 1" = 6,5" en twips
const TIPO_PAGINA = { width: 12240, height: 15840 }; // Carta
const MARGEN = 1440;                  // 1 pulgada

// Pagina apaisada para tablas anchas. APA 7 lo permite: una tabla que no cabe
// comoda en vertical se presenta en una pagina horizontal aparte, manteniendo
// margenes de 1 pulgada. El ancho util pasa de 6,5" a 9".
const ANCHO_CONTENIDO_H = TIPO_PAGINA.height - 2 * MARGEN; // 11" - 2" = 9"

// Criterio para decidir si una tabla necesita pagina apaisada. Se dispara por
// ANCHURA, no por numero de filas: una tabla de 200 filas y 3 columnas se lee
// bien en vertical, mientras que una de 8 columnas queda aplastada.
const UMBRAL_COLUMNAS_ANCHAS = 6;    // a partir de 6 columnas
const UMBRAL_LINEAS_APRETADAS = 8;    // lineas por celda que ya no se leen
const ANCHO_CARACTER_TWIP = 120;      // ancho medio de un caracter a 12 pt
const MARGEN_CELDA_LR = 160;          // margenes izquierdo y derecho de celda

// Portada en 3 zonas verticales reales: titulo arriba, integrantes en el centro
// de la pagina y bloque institucional anclado abajo. Se construye como una tabla
// SIN bordes de 1 columna x 3 filas, con alto fijo y alineacion vertical, en vez
// de con parrafos vacios de relleno: asi las zonas no se mueven segun quantos
// integrantes o campos traiga el documento.
//
// INVARIANTE: las tres filas tienen que sumar EXACTAMENTE ALTO_UTIL. Si suman mas,
// la ultima fila se desborda a una pagina nueva y el documento gana una pagina.
// Por eso la tercera zona absorbe el redondeo en vez de valer un tercio fijo.
const ALTO_UTIL = TIPO_PAGINA.height - 2 * MARGEN;      // 9" = 12960 twips = 648 pt
const ZONA_ALTO = Math.round(ALTO_UTIL / 3);            // 4320: y 72..288 pt, logo + titulo
const ZONA_CENTRO = Math.round(ALTO_UTIL / 3);          // 4320: y 288..504 pt, integrantes
const ZONA_BAJO = ALTO_UTIL - ZONA_ALTO - ZONA_CENTRO;  // 4320: y 504..720 pt, bloque
const AIRE_ZONA_ALTA = 720;          // el titulo arranca ~0,5" por debajo del margen

const B = (color) => ({ style: BorderStyle.SINGLE, size: 4, color: color || "000000" });
const SIN_BORDE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

/** Segmentos de texto del manifiesto -> TextRun[], respetando negrita/cursiva. */
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

/** Inserta cursiva solo en el rango confirmado de una referencia. */
function runsDeReferencia(ref) {
  const texto = ref.texto;
  const rango = ref.cursiva_confirmada || ref.cursiva_propuesta;
  if (!rango) {
    return [new TextRun({ text: texto, font: FUENTE, size: TAM })];
  }
  const i = texto.indexOf(rango);
  if (i < 0) {
    log(`    [aviso] el rango en cursiva "${rango.slice(0, 40)}..." no aparece en la referencia; se deja en plano.`);
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
// Encabezado con numero de pagina (esquina superior derecha, APA 7)
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
// Portada
// ---------------------------------------------------------------------------

function parrafoCentrado(children, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: DOBLE, before: 0, after: 0 },
    children,
    ...opts,
  });
}

// Linea de la portada, centrada.
function lineaPortada(children, spacing) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: spacing || { line: DOBLE, before: 0, after: 0 },
    children,
  });
}

/**
 * Portada en 3 zonas verticales: logo + titulo arriba, integrantes en el centro
 * vertical de la pagina y el resto de los datos institucionales anclados abajo.
 *
 * Antes esto se hacia apilando parrafos y rellenando con lineas vacias, lo que
 * dejaba a los integrantes pegados al titulo y el bloque institucional a media
 * pagina.
 *
 * Se hace con UNA tabla de 1 columna x 3 filas, sin bordes visibles, con alto
 * fijo por fila y alineacion vertical. Es una tabla y no tres apiladas a proposito:
 * dos o mas tablas consecutivas se fusionan en Word y LibreOffice las trata de
 * forma inconsistente, y el conjunto se desborda a una pagina extra.
 */
function construirPortada() {
  const p = M.portada || {};
  const centro = (text, bold = false) => [new TextRun({ text, bold, font: FUENTE, size: TAM })];

  // --- Zona alta: logo opcional y titulo ---------------------------------
  const zonaAlta = [];
  if (p.logo) {
    // El logo es un recurso opcional que aporta el usuario. Nunca se inventa el
    // nombre de una institucion.
    const ext = path.extname(p.logo).toLowerCase();
    const tipo = ext === ".jpg" || ext === ".jpeg" ? "jpg" : (ext === ".gif" ? "gif" : "png");
    zonaAlta.push(lineaPortada([new ImageRun({
      type: tipo,
      data: fs.readFileSync(p.logo),
      transformation: { width: 160, height: 160 },
    })], { line: 240, before: 0, after: 0 }));   // simple: una imagen no se interlinea
    if (p.logo_leyenda) zonaAlta.push(lineaPortada(centro(p.logo_leyenda)));
  }
  // Sin logo no se inserta NADA: ni placeholder, ni marco, ni la palabra "Logo"
  // que antes se colaba en el PDF.
  if (p.titulo) zonaAlta.push(lineaPortada(centro(p.titulo, true)));

  // --- Zona central: integrantes -----------------------------------------
  // references/portada-institucional.md: los integrantes van en UN SOLO parrafo
  // centrado, separados por coma y con "y" antes del ultimo, sin importar
  // cuantos sean. Antes se emitia un parrafo por persona.
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

  // --- Zona baja: datos institucionales ----------------------------------
  // El orden depende de lo que traiga el documento: si un campo no esta, no se
  // deja linea en blanco. Los ausentes simplemente no se emiten.
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
    // ATLEAST y no EXACT: si el contenido no cabe (titulo de 3 lineas, 6
    // integrantes) la fila crece en vez de recortar el texto.
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

  return [new Table({
    width: { size: ANCHO_CONTENIDO, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    // El aire de la zona alta se DESCUENTA de su altura: las filas tienen que
    // sumar justo el alto util (12960 twips). Sumar por encima desborda la
    // ultima fila a una pagina nueva.
    rows: [
      filaZona(zonaAlta, ZONA_ALTO - AIRE_ZONA_ALTA, VerticalAlign.TOP, AIRE_ZONA_ALTA),
      filaZona(zonaCentro, ZONA_CENTRO, VerticalAlign.CENTER),
      filaZona(zonaBaja, ZONA_BAJO, VerticalAlign.BOTTOM),
    ],
  })];
}

// ---------------------------------------------------------------------------
// Tabla de contenido con campos PAGEREF
// ---------------------------------------------------------------------------

function lineaTOC(texto, bookmark, nivel = 1) {
  return new Paragraph({
    spacing: { line: DOBLE, before: 0, after: 0 },
    indent: { left: (nivel - 1) * SANGRIA_1RA },
    tabStops: [{ type: TabStopType.RIGHT, position: ANCHO_CONTENIDO, leader: LeaderType.DOT }],
    children: [
      new TextRun({ text: texto, font: FUENTE, size: TAM, bold: nivel === 1 }),
      new TextRun({ text: "\t", font: FUENTE, size: TAM }),
      new SimpleField(`PAGEREF ${bookmark} \\h`),
    ],
  });
}

function construirIndices() {
  const hijos = [];
  const opts = M.opciones || {};

  // --- Tabla de contenido: SIEMPRE -------------------------------------
  hijos.push(parrafoCentrado([new TextRun({ text: "Tabla de contenido", bold: true, font: FUENTE, size: TAM })]));

  // Los titulos repetidos se desambiguan SOLO si el manifiesto lo pide, para
  // que la TOC sea navegable sin alterar el texto del documento.
  const vistos = new Map();
  for (const s of M.secciones) {
    let texto = s.texto;
    const n = (vistos.get(texto) || 0) + 1;
    vistos.set(texto, n);
    if (n > 1 && opts.dedupe_toc) {
      texto = `${texto} (${s.hid})`;
      log(`    TOC: titulo repetido -> "${texto}"`);
    }
    hijos.push(lineaTOC(texto, `apa_sec_${s.hid}`, s.nivel));
  }

  // --- Indice de tablas: solo si HAY tablas -----------------------------
  if (opts.indice_tablas && M.tablas && M.tablas.length) {
    hijos.push(new Paragraph({ children: [new PageBreak()] }));
    hijos.push(parrafoCentrado([new TextRun({ text: "Indice de tablas", bold: true, font: FUENTE, size: TAM })]));
    for (const t of M.tablas) {
          const et = `${t.titulo}`;
      hijos.push(lineaTOC(`Tabla ${t.indice}. ${et}`, `apa_tbl_${t.indice}`, 1));
    }
  } else if (M.tablas && M.tablas.length) {
    log("    Indice de tablas OMITIDO por decision explicita del manifiesto.");
  }

  // --- Indice de figuras: solo si HAY figuras ---------------------------
  if (opts.indice_figuras && M.figuras && M.figuras.length) {
    // Solo las figuras que de verdad entran en el cuerpo. Listar tambien las
    // omitidas deja una linea con "Error: no se encontro el origen de la
    // referencia", porque su marcador nunca se crea.
    const figsValidas = M.figuras.filter((f) => f.existe);
    const omitidas = M.figuras.length - figsValidas.length;
    if (omitidas > 0) {
      log(`    Indice de figuras: se omiten ${omitidas} linea(s) por figura ausente.`);
    }
    if (figsValidas.length) {
      hijos.push(new Paragraph({ children: [new PageBreak()] }));
      hijos.push(parrafoCentrado([new TextRun({ text: "Indice de figuras", bold: true, font: FUENTE, size: TAM })]));
      for (const f of figsValidas) {
            const et = `${f.titulo}`;
        hijos.push(lineaTOC(`Figura ${f.indice}. ${et}`, `apa_fig_${f.indice}`, 1));
      }
    }
  } else if (M.figuras && M.figuras.length) {
    log("    Indice de figuras OMITIDO por decision explicita del manifiesto.");
  }

  // El cuerpo arranca SIEMPRE en pagina nueva. No se deja que lo empuje el
  // primer encabezado de nivel 1: en el cuerpo ese heading va el primero, asi
  // que su guardian "si ya hay algo escrito" nunca se cumple y el texto se
  // colgaba debajo del indice de tablas, en la misma pagina.
  if (hijos.length) {
    hijos.push(new Paragraph({ children: [new PageBreak()] }));
  }

  return hijos;
}

// ---------------------------------------------------------------------------
// Tablas
// ---------------------------------------------------------------------------

// Estima cuantas lineas ocupa el texto de una celda si la tabla se maqueta en
// vertical. Con `layout: FIXED` cada columna mide ancho_total / ncols, asi que
// el numero de caracteres por linea sale de ahi, no del contenido real.
function lineasDeCelda(texto, anchoColumna) {
  const utiles = anchoColumna - MARGEN_CELDA_LR;
  if (utiles <= 0) return Number.POSITIVE_INFINITY;
  const porLinea = Math.floor(utiles / ANCHO_CARACTER_TWIP);
  if (porLinea < 1) return Number.POSITIVE_INFINITY;
  return texto.split("\n").reduce((total, linea) => total + Math.max(1, Math.ceil(linea.length / porLinea)), 0);
}

// Decide si una tabla va en pagina apaisada y deja constancia del motivo, que
// se imprime en el log y se guarda en el manifiesto para poder auditarlo.
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

  // La decision se toma con la tabla en vertical. Si ahi ya esta apretada pero
  // en apaisado se lee bien, el apaisado es la solucion y no solo un luxejo.
  let motivo = "";
  if (ncols >= UMBRAL_COLUMNAS_ANCHAS) {
    motivo = `${ncols} columnas (umbral ${UMBRAL_COLUMNAS_ANCHAS})`;
  } else if (peorVertical >= UMBRAL_LINEAS_APRETADAS && peorApaisada < UMBRAL_LINEAS_APRETADAS) {
    motivo = `celda de hasta ${peorVertical} lineas en vertical, ${peorApaisada} en apaisado`;
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

  // Numero en negrita encima de la tabla; titulo en cursiva debajo (APA 7).
  // El marcador va en el numero: es lo que el indice de tablas apunta con
  // PAGEREF, y si no existe, el indice sale con "Error: no se encontro el
  // origen de la referencia" en cada linea.
  hijos.push(new Paragraph({
    spacing: { line: DOBLE, before: 0, after: 0 },
    keepNext: true,
    children: [new Bookmark({
      id: `apa_tbl_${t.indice}`,
      children: [new TextRun({ text: `Tabla ${t.indice}`, bold: true, font: FUENTE, size: TAM })],
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
        width: { size: Math.floor(ancho / ncols), type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        // APA 7: sin lineas verticales. Solo el borde superior de la tabla,
        // la linea bajo la fila de encabezado y el borde inferior.
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

  // Borde inferior solo en la ultima fila: APA 7 no usa lineas verticales.
  // (se resuelve al crear cada celda, no mutando el XML generado)

  hijos.push(new Table({
    rows: filasDoc,
    width: { size: ancho, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
  }));

  if (t.nota) {
    hijos.push(new Paragraph({
      spacing: { line: 240, before: 0, after: 0 },
      children: [
        new TextRun({ text: "Nota. ", italics: true, font: FUENTE, size: TAM }),
        new TextRun({ text: t.nota, font: FUENTE, size: TAM }),
      ],
    }));
  }
  return hijos;
}

// ---------------------------------------------------------------------------
// Figuras
// ---------------------------------------------------------------------------

function construirFigura(f) {
  const hijos = [];
  if (!f.existe) {
    log(`    Figura ${f.indice} OMITIDA: no existe el archivo "${f.ruta}".`);
    return hijos;
  }

  // El marcador va en el numero, que es lo que apunta el indice de figuras.
  hijos.push(new Paragraph({
    spacing: { line: DOBLE, before: 0, after: 0 },
    keepNext: true,
    children: [new Bookmark({
      id: `apa_fig_${f.indice}`,
      children: [new TextRun({ text: `Figura ${f.indice}`, bold: true, font: FUENTE, size: TAM })],
    })],
  }));
  if (f.titulo) {
    hijos.push(new Paragraph({
      spacing: { line: DOBLE, before: 0, after: 0 },
      keepNext: true,
      children: [new TextRun({ text: f.titulo, italics: true, font: FUENTE, size: TAM })],
    }));
  }

// docx trabaja en pixeles a 96 DPI: 1 pulgada = 96 px.

// Tamano real de la imagen leyendo su cabecera. No se usa `new ImageRun(...)`
// como sonda porque NO expone las dimensiones: devolvia 1x1 y todas las figuras
// salian cuadradas (6,50 x 6,50 in) aunque el original fuera 14,58 x 5,00.
// Asi se lee el ancho y el alto de verdad, sin anadir dependencias.
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
      // SOF0..SOF15, saltando los no-carga (DHT/DQT/DRI...).
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
      `No se pudo leer el tamaño de la imagen "${f.ruta}". Se sabe leer la ` +
      "cabecera de PNG, JPEG, GIF y BMP; si es otro formato, conviértela antes.");
  }
  const ext = path.extname(f.ruta).toLowerCase();
  const tipo = ext === ".jpg" || ext === ".jpeg" ? "jpg" : (ext === ".gif" ? "gif" : "png");

  // El ancho sale del manifiesto (caja del original) y el alto del archivo, para
  // no deformar. Si aun asi se pasa de alto, se escala en bloque.
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
  if (altoPx > 648) {                       // 6,75 in: la mitad de la hoja util
    const k = 648 / altoPx;
    anchoPx = Math.round(anchoPx * k);
    altoPx = 648;
  }
  // Una imagen nunca se agranda mas alla de su propio tamano en pixeles: estirar
  // un icono de 260 px hasta 6,5 pulgadas solo produce un borron. Se respeta la
  // caja del original, salvo que pida mas grande de lo que la imagen da.
  if (anchoPx > dim.ancho) {
    const k = dim.ancho / anchoPx;
    anchoPx = dim.ancho;
    altoPx = Math.round(altoPx * k);
  }

  // La nota general de la FIGURA va ENCIMA de la imagen (APA 7 la distingue de
  // la de tabla, que va debajo). Antes se empujaba despues del ImageRun, que es
  // la posicion de la nota especifica y no de la general.
  //
  // keepNext es obligatorio: sin el, una nota al final de una pagina se queda
  // sola y su imagen salta a la siguiente, dejando la nota de la figura N en la
  // pagina anterior. Es exactamente el caso que produce un documento incoherente.
  if (f.nota) {
    hijos.push(new Paragraph({
      spacing: { line: 240, before: 0, after: 0 },
      keepNext: true,
      children: [
        new TextRun({ text: "Nota. ", italics: true, font: FUENTE, size: TAM }),
        new TextRun({ text: f.nota, font: FUENTE, size: TAM }),
      ],
    }));
  }

  hijos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 0, after: 0 },
    children: [new ImageRun({
      type: tipo,
      data: fs.readFileSync(f.ruta_absoluta),
      transformation: { width: anchoPx, height: altoPx },
    })],
  }));

  return hijos;
}

// ---------------------------------------------------------------------------
// Cuerpo: recorre el flujo de bloques del manifiesto
// ---------------------------------------------------------------------------

// Devuelve el cuerpo partido en tramos, en vez de una lista plana. Un tramo
// es `{paisaje, hijos, ocupado}` y cada tramo se convierte en una seccion del
// .docx. Las tablas anchas van en tramos apaisados y el texto vuelve a vertical
// despues.
//
// `hayContenidoPrevio` dice si algo ya se escribio en la seccion antes de este
// cuerpo (son los indices). No se usa para rellenar el tramo, sino para decidir
// si un encabezado de nivel 1 necesita salto de pagina: sin el, el primer titulo
// se pegaria al final del indice.
function construirCuerpoSegmentado(hayContenidoPrevio) {
  const tramos = [];
  const porHid = new Map(M.secciones.map((s) => [s.hid, s]));
  const decisiones = [];

  let actual = { paisaje: false, hijos: [], ocupado: !!hayContenidoPrevio };
  tramos.push(actual);

  // Emitir un bloque cualquiera en el tramo actual.
  const emitir = (x) => {
    actual.hijos.push(x);
    actual.ocupado = true;
  };

  // Abre un tramo. Si el anterior ya es apaisado, se reutiliza: dos tablas
  // anchas seguidas comparten pagina horizontal en vez de alternar dos veces.
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
        // La portada ya se construyo aparte, con su propio formato.
        break;

      case "h": {
        const s = porHid.get(b.hid) || { nivel: b.nivel, texto: b.texto };
        const nivel = s.nivel;
        const nivel1 = nivel === 1;
        // APA 7: n1 centrado negrita; n2 margen izq. negrita; n3 margen izq.
        // negrita+cursiva; n4 y n5 con sangria de primera linea, negrita (n5
        // tambien cursiva) y punto final, con el texto siguiendo en la MISMA
        // linea. Antes todo lo que no fuera nivel 1 salia en negrita y alineado
        // a la izquierda, asi que los niveles 3, 4 y 5 no se distinguian.
        const nivel3 = nivel === 3;
        const nivelEnLinea = nivel === 4 || nivel === 5;
        const estilo = { bold: true, italics: nivel3 || nivel === 5 };
        // Los hijos se montan ANTES de crear el Paragraph: la libreria no
        // expone `children` como array mutable, asi que pushear despues falla.
        const hijosCab = [new Bookmark({
          id: `apa_sec_${b.hid}`,
          children: runsDe(b.segmentos, estilo),
        })];
        if (nivelEnLinea) {
          // Titulo en linea: el punto final cierra el titulo y el cuerpo del
          // parrafo continua a continuacion en el mismo parrafo.
          hijosCab.push(new TextRun({ text: ". ", font: FUENTE, size: TAM, ...estilo }));
          if (b.en_linea_segmentos) hijosCab.push(...runsDe(b.en_linea_segmentos));
        }
        // APA 7: nivel 1 empieza en pagina nueva. Se marca con
        // `pageBreakBefore` en el propio encabezado y NO con un parrafo
        // suelto que contenga un PageBreak: ese parrafo vacio, cuando la pagina
        // anterior ya esta llena, no cabe en ella, salta a la siguiente y deja
        // una pagina en blanco antes del titulo.
        const p = new Paragraph({
          alignment: nivel1 ? AlignmentType.CENTER : AlignmentType.LEFT,
          spacing: { line: DOBLE, before: 0, after: 0 },
          keepNext: !nivelEnLinea,
          pageBreakBefore: nivel1 && actual.ocupado,
          indent: nivelEnLinea ? { firstLine: SANGRIA_1RA } : undefined,
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
          // La tabla va en su propia pagina apaisada, con su titulo y su nota.
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

      case "nota_tabla":
        break;

      default:
        log(`    [aviso] tipo de bloque desconocido: "${b.tipo}". Se omite.`);
    }
  }

  // Los tramos que se quedaron vacios (tipicamente el vertical anterior y
  // posterior a una tabla ancha) no se usan porque darian paginas en blanco.
  const conContenido = tramos.filter((t) => t.hijos.length > 0);

  for (const d of decisiones) {
    if (d.ancha) {
      log(`  Tabla ${d.indice}: pagina apaisada -> ${d.motivo}.`);
    } else {
      log(`  Tabla ${d.indice}: vertical (${d.ncols} columnas, celda de hasta ${d.peorVertical} lineas).`);
    }
  }
  const apaisadas = decisiones.filter((d) => d.ancha).length;
  log(`  Orientacion: ${apaisadas} tabla(s) apaisada(s) de ${decisiones.length}.`);

  return conContenido;
}

// ---------------------------------------------------------------------------
// Referencias
// ---------------------------------------------------------------------------

function construirReferencias() {
  const hijos = [];
  const refs = (M.referencias || []).slice();
  // APA 7 exige orden alfabetico. El .md puede no venir ordenado.
  refs.sort((a, b) => a.texto.localeCompare(b.texto, "es", { sensitivity: "base" }));

  log(`  Referencias: ${refs.length} (ordenadas alfabeticamente en el .docx).`);
  for (const r of refs) {
    if (r.cursiva_propuesta) {
      log(`    cursiva a revisar: "${r.cursiva_propuesta.slice(0, 60)}"`);
    }
    hijos.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: DOBLE, before: 0, after: 0 },
      // Sangria francesa: 0,5" a la izquierda y -0,5" de primera linea.
      indent: { left: SANGRIA_1RA, hanging: SANGRIA_1RA },
      children: runsDeReferencia(r),
    }));
  }
  return hijos;
}

// ---------------------------------------------------------------------------
// Ensamblado
// ---------------------------------------------------------------------------

async function main() {
  console.log("=== FASE 2: construccion del .docx ===");
  log(`Manifiesto : ${args.manifiesto}`);
  log(`Salida     : ${DOCX}`);
  log(`Fuente .md : ${M.fuente ? M.fuente.md : "(desconocido)"}`);
  const p = M.portada || {};
  log(`Portada    : titulo=${p.titulo ? "si" : "NO"} autores=${(p.autores || []).length} docente=${p.docente ? "si" : "NO"} fecha=${p.fecha ? "si" : "NO"}`);
  if (p.campos_faltantes && p.campos_faltantes.length) {
    log(`Portada    : CAMPOS FALTANTES -> ${p.campos_faltantes.join(", ")}`);
    log("            (la portada se genera igual; avisar al usuario de lo que falta)");
  }
  log(`Tablas     : ${M.tablas.length}   Figuras: ${M.figuras.length}   Referencias: ${M.referencias.length}`);
  log(`Indices    : tablas=${M.opciones.indice_tablas ? "si" : "no"} figuras=${M.opciones.indice_figuras ? "si" : "no"}`);
  log("");

  const indices = construirIndices();
  // Los indices se quedan en la MISMA seccion que el primer tramo vertical del
  // cuerpo. Si se separaran, el salto de seccion anadiria una pagina en blanco
  // entre el indice y el texto, que antes no existia.
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
  // `docx` intercambia ancho y alto por su cuenta cuando ve
  // orientation=LANDSCAPE. Hay que darle las medidas EN VERTICAL (12240 x
  // 15840) y marcar la orientacion: pasarlas ya invertidas produce un
  // `w:pgSz w:w="12240" w:h="15840" w:orient="landscape"`, que Word y
  // LibreOffice leen como vertical y la tabla sale apretada.
  const propsApaisadas = {
    page: {
      size: { width: TIPO_PAGINA.width, height: TIPO_PAGINA.height, orientation: PageOrientation.LANDSCAPE },
      margin: { top: MARGEN, right: MARGEN, bottom: MARGEN, left: MARGEN, header: 720, footer: 720 },
    },
  };

  const secciones = [
    // Seccion 1: portada sola, SIN numero de pagina.
    {
      properties: { ...propsVerticales, type: SectionType.NEXT_PAGE },
      headers: { default: headerVacio },
      footers: { default: footerVacio },
      children: construirPortada(),
    },
  ];

  // Si el primer tramo es apaisado, los indices necesitan seccion vertical
  // propia; si no, ya van dentro de ese primer tramo y no se duplican.
  if (indices.length && (!primerTramo || primerTramo.paisaje)) {
    secciones.push({
      properties: { ...propsVerticales, type: SectionType.NEXT_PAGE },
      headers: { default: headerConPagina },
      footers: { default: footerVacio },
      children: indices,
    });
  }

  // Una seccion por tramo del cuerpo. Cada tramo abre pagina nueva, asi que una
  // tabla apaisada queda en su propia pagina y el texto posterior vuelve a
  // vertical. Todas repiten cabecera y pie para no perder la numeracion.
  for (const tramo of tramos) {
    secciones.push({
      properties: { ...(tramo.paisaje ? propsApaisadas : propsVerticales), type: SectionType.NEXT_PAGE },
      headers: { default: headerConPagina },
      footers: { default: footerVacio },
      children: tramo.hijos,
    });
  }

  // Seccion final: referencias, siempre en vertical.
  //
  // El encabezado "Referencias Bibliograficas" NO viene de aqui: es un heading
  // de nivel 1 del .md y ya quedo al final del ultimo tramo del cuerpo. Con un
  // salto de pagina se quedaba solo en una pagina y las entradas empezaban en la
  // siguiente. Si el tramo anterior ya es vertical se usa una seccion CONTINUA,
  // que no rompe pagina: el encabezado baja con la primera entrada. Solo si el
  // documento termina en una tabla apaisada hace falta una pagina nueva, porque
  // hay que volver a vertical.
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
    creator: "generar-pdf-apa",
    title: p.titulo || "Documento APA 7",
    description: `Generado desde ${M.fuente ? M.fuente.md : "fuente .md"}`,
    styles: {
      default: {
        document: { run: { font: FUENTE, size: TAM }, paragraph: { spacing: { line: DOBLE, before: 0, after: 0 } } },
      },
    },
    sections: secciones,
  });

  const buffer = await Packer.toBuffer(doc);
  const dir = path.dirname(path.resolve(DOCX));
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(DOCX, buffer);

  log("");
  log(`.docx escrito: ${DOCX} (${buffer.length} bytes)`);
  log("Campos PAGEREF insertados. LibreOffice los resuelve al exportar;");
  log("para fijarlos tambien en el .docx hace falta la macro opt-in.");

  if (args.log) {
    const ld = path.dirname(path.resolve(args.log));
    if (ld && !fs.existsSync(ld)) fs.mkdirSync(ld, { recursive: true });
    fs.writeFileSync(args.log, logLines.join("\n") + "\n", "utf8");
    log(`Log escrito : ${args.log}`);
  }
}

main().catch((e) => {
  console.error("FALLO en build-docx.js:", e && e.stack ? e.stack : e);
  process.exit(1);
});
