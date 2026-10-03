#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verificar-pdf.py - Checks the generated PDF against its MANIFEST.json.

This script is STEP 7 of the pipeline. It does not judge the academic quality of
the document (that is up to the agent reading the text), it checks whatever can
be checked automatically and REPRODUCIBLY:

  1. The PDF opens and has pages.
  2. The page size is Letter (612 x 792 pt) with 1 inch margins.
  3. The cover page does NOT carry a page number.
  4. Page 2 DOES carry the number, and in the top right corner.
  5. Every section of the manifest appears as a heading in the PDF.
  6. Every table appears, with its number and its content.
  7. Every figure appears.
  8. Every reference appears in the references section.
  9. The numbers in the table of contents MATCH the real page of each
     heading. This is the check that really matters: an unresolved PAGEREF
     leaves a 0 or a dash, and a page mismatch invalidates the TOC.
 10. There are no broken bookmarks ("Error! Bookmark not defined", "0", "??").
 11. Only the declared typeface is used.
 12. Double spacing in the body paragraphs.

Usage:
    python verificar-pdf.py --pdf salida.pdf --manifiesto MANIFEST.json
                            [--log salida/_logs/04-verificacion.log]
                            [--json salida/_logs/04-verificacion.json]

Returns 0 if there are no FAILURES, 1 if there is any.
"""

import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import datetime

try:
    import pymupdf
except ImportError:  # pymupdf < 1.24.3 published the same module as 'fitz'
    import fitz as pymupdf

# The font rule lives in lib/ so it can be tested without pymupdf. This script
# is otherwise standalone, hence the explicit path insert.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import fuentes as fuentes_mod  # noqa: E402

PT_PULGADA = 72.0
PAGINA_CARTA = (612.0, 792.0)

# Vertical bands of the cover page, in points from the top edge of the sheet.
# Must match the ZONA_* constants of build-docx.js.
PORTADA_BANDA_ALTA = (72.0, 288.0)
PORTADA_BANDA_CENTRO = (288.0, 504.0)
PORTADA_BANDA_BAJA = (504.0, 720.0)


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def norm(text):
    """Lowercase, no accents, no punctuation, collapsed spaces."""
    t = unicodedata.normalize("NFD", (text or "").lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def y_de_texto(bloques_pagina, aguja, prefijo=24):
    """y of the first block that contains the beginning of `aguja`, or None.

    It is used to anchor the zone check to the REAL texts from the manifest.
    Counting blocks by position does not work: a heading that spans two lines
    produces two blocks and shifts the guessed "the second block is the
    members".
    """
    a = norm(aguja)[:prefijo]
    if len(a) < 6:
        return None
    for b in sorted(bloques_pagina, key=lambda x: x[1]):
        if a and a in norm(b[4]):
            return b[1]
    return None


def cargar_texto(pdf):
    """Returns (list_of_texts_per_page, list_of_blocks_per_page)."""
    doc = pymupdf.open(pdf)
    paginas = []
    bloques = []
    for i in range(doc.page_count):
        paginas.append(doc[i].get_text())
        bloques.append(doc[i].get_text("blocks"))
    return doc, paginas, bloques


def paginas_con(paginas, aguja_norm, ignorar_portada=True, desde=0):
    """0-based indices of the pages whose normalized text contains the needle."""
    out = []
    for i, t in enumerate(paginas):
        if i < desde:
            continue
        if ignorar_portada and i == 0:
            continue
        if aguja_norm and aguja_norm in norm(t):
            out.append(i)
    return out


def lineas_toc(paginas):
    """
    Extracts the table of contents lines: 'Heading ......... 12'.
    It relies on the dot leader, which is what a tab with dots generates.
    """
    filas = []
    for i, t in enumerate(paginas[:4]):     # the TOC lives at the beginning
        for linea in t.splitlines():
            l = linea.strip()
            m = re.match(r"^(.*?)\s*\.{4,}\s*(\d{1,4})\s*$", l)
            if m:
                filas.append({"titulo": m.group(1).strip(), "pagina": int(m.group(2)), "pagina_pdf": i})
    return filas


RE_ETIQUETA_INDICE = re.compile(r"^(tabla|figura)\s+\d+")


def pagina_fin_de_indices(paginas, lineas):
    """
    Last page occupied by the indices (TOC, table index, figure index).

    It is essential to know it: the heading of each section ALSO appears
    inside the table of contents itself. Searching "la pagina donde esta
    Introduccion" without excluding the TOC returns page 2 (the TOC's one)
    instead of the real page, and produces 29 false mismatches.
    """
    if not lineas:
        return 0
    return max(f["pagina_pdf"] for f in lineas)


def texto_en_orden_visual(doc, desde, hasta):
    """
    Text of pages [desde, hasta) in real READING order.

    `get_text()` returns the blocks in the PDF's internal flow order, not the
    visual one: when there are italic fragments inside each reference, the
    order gets scrambled and an "alphabetical order" check gives a false
    negative. Sorting by (y, x) reconstructs the correct reading order.
    """
    partes = []
    for i in range(desde, min(hasta, doc.page_count)):
        bloques = [b for b in doc[i].get_text("blocks") if len(b) >= 5 and isinstance(b[4], str)]
        bloques.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))
        partes.append("\n".join(b[4] for b in bloques))
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

class Resultado:
    def __init__(self):
        self.items = []

    def anota(self, nombre, ok, detalle="", critico=True):
        self.items.append({
            "nombre": nombre,
            "ok": bool(ok),
            "detalle": detalle,
            "critico": critico,
        })
        return ok

    @property
    def fallas(self):
        return [i for i in self.items if not i["ok"] and i["critico"]]

    @property
    def avisos(self):
        return [i for i in self.items if not i["ok"] and not i["critico"]]


def verificar(args):
    R = Resultado()
    with open(args.manifiesto, encoding="utf-8") as fh:
        M = json.load(fh)

    doc, paginas, bloques = cargar_texto(args.pdf)
    total = doc.page_count
    texto_total = "\n".join(paginas)
    norm_total = norm(texto_total)

    # Page from which the body starts: the indices repeat themselves, so
    # searching for a heading before here gives a false positive.
    _toc = lineas_toc(paginas)
    fin_indices = pagina_fin_de_indices(paginas, _toc)
    CUERPO = fin_indices + 1

    # --- 1. Opening -----------------------------------------------------
    R.anota("The PDF opens and has pages", total > 0,
            "%d pages" % total)

    # --- 2. Page size and margins --------------------------------------
    r0 = doc[0].rect if total else None
    if r0:
        w, h = round(r0.width, 1), round(r0.height, 1)
        R.anota("Letter-size page (612x792 pt)", abs(w - PAGINA_CARTA[0]) < 2 and abs(h - PAGINA_CARTA[1]) < 2,
                "actual size %.1f x %.1f pt" % (w, h))
    else:
        R.anota("Letter-size page (612x792 pt)", False, "no pages", critico=False)

    # Left margin: text must not start before 1 inch (with a tolerance margin
    # of 2 pt due to the rounding of the text engine).
    xs = []
    for pg in bloques[1:] if len(bloques) > 1 else bloques:
        for b in pg:
            if len(b) >= 5 and isinstance(b[4], str) and b[4].strip():
                xs.append(b[0])
    if xs:
        minx = min(xs)
        R.anota("Left margin >= 1 inch", minx >= PT_PULGADA - 2,
                "min text x = %.1f pt (1 inch = 72 pt)" % minx)
    else:
        R.anota("Left margin >= 1 inch", False, "no text blocks found", critico=False)

    # --- 3 and 4. Cover page numbering ---------------------------------
    # The cover page must not show any loose number in the top right corner;
    # page 2 must, and it must match the real 2.
    def numero_en_esquina(i):
        if i >= total:
            return None
        zona = doc[i].get_text("text", clip=pymupdf.Rect(0, 0, r0.width, 60))
        m = re.search(r"\b(\d{1,4})\b", zona)
        return int(m.group(1)) if m else None

    n1 = numero_en_esquina(0)
    R.anota("The cover page does NOT carry a page number", n1 is None,
            "number detected in the top right corner: %s" % (n1 if n1 is not None else "none"))

    if total > 1:
        n2 = numero_en_esquina(1)
        R.anota("Page 2 shows the number 2", n2 == 2,
                "number detected: %s" % (n2 if n2 is not None else "none"))
    else:
        R.anota("Page 2 shows the number 2", False, "the PDF has only 1 page", critico=False)

    # --- 3b. Cover page in 3 zones (WARNINGS, never block) ---------------
    # These four are NOT failures: the cover page is built with whatever data
    # exists. If the logo or the instructor's title is missing we warn so the agent
    # can ask about it, but the document is still produced. Blocking would be counterproductive.
    portada = M.get("portada") or {}

    if not total:
        R.anota("The cover page respects the 3 zones", False, "the PDF has no pages", critico=False)
    else:
        pg0 = doc[0]
        # Blocks in the header strip (y < 72 pt) are ignored, since only the
        # page number can appear there.
        pb = [b for b in bloques[0]
              if len(b) >= 5 and isinstance(b[4], str) and b[4].strip() and b[1] >= PT_PULGADA - 2]
        pb.sort(key=lambda b: b[1])
        alto = PORTADA_BANDA_ALTA[1]        # 288: end of the high band
        bajo = PORTADA_BANDA_BAJA[0]        # 504: end of the middle band
        centro = PORTADA_BANDA_CENTRO[0]    # 288: start of the middle band

        # Each zone is anchored to its own text from the manifest, not to the
        # order of the blocks: so a two-line heading does not skew the measurement.
        autores_txt = [a for a in (portada.get("autores") or []) if str(a).strip()]
        campos_bajos = [c for c in ("docente", "materia_nrc", "facultad",
                                    "vicerrectoria", "sede", "fecha")
                        if portada.get(c)]
        y_titulo = y_de_texto(pb, portada.get("titulo") or "")
        y_autores = y_de_texto(pb, autores_txt[0]) if autores_txt else None

        problemas = []
        if not pb:
            R.anota("The cover page respects the 3 zones", False,
                    "no text found on the cover page", critico=False)
        else:
            if y_titulo is not None and y_titulo >= alto:
                problemas.append("the title starts at y=%.0f pt, it should be above y=%.0f" % (y_titulo, alto))
            if autores_txt and y_autores is None:
                problemas.append("the members were not found on the cover page")
            elif y_autores is not None and not (centro <= y_autores < bajo):
                problemas.append("the members are at y=%.0f pt, they should fall between %.0f and %.0f"
                                 % (y_autores, centro, bajo))
            # The low band is only checked if the manifest provides data to
            # anchor to. On a cover page reduced to a title there is no block to
            # measure, and warning about that would be a false positive.
            y_bajo = None
            if campos_bajos:
                y_bajo = y_de_texto(pb, portada.get(campos_bajos[0]))
                if y_bajo is None:
                    # The field exists in the manifest but does not appear on
                    # page 1: most likely the cover page overflowed and the
                    # institutional block is on page 2.
                    problemas.append("'%s' does not appear on the cover page: the cover page overflowed onto another page"
                                     % campos_bajos[0])
                elif y_bajo <= bajo:
                    problemas.append("the institutional block starts at y=%.0f pt, it should be below y=%.0f"
                                     % (y_bajo, bajo))
            fin = pb[-1][3]
            if campos_bajos and fin <= bajo:
                problemas.append("the institutional block ends at y=%.0f pt, it should reach the low band" % fin)
            if campos_bajos:
                zona_baja_txt = ("y=%.0f..%.0f" % (y_bajo, fin)) if y_bajo is not None else "not found on page 1"
            else:
                zona_baja_txt = "not applicable (no data)"
            detalle = "title y=%s | members y=%s | institutional block %s" % (
                "%.0f" % y_titulo if y_titulo is not None else "n/a",
                "%.0f" % y_autores if y_autores is not None else "n/a",
                zona_baja_txt)
            R.anota("The cover page respects the 3 zones", not problemas,
                    "; ".join(problemas) if problemas else detalle,
                    critico=False)

        # The members go in a SINGLE paragraph: if there are two or more, no
        # name may start a new line, except the last one if the previous one
        # ends exactly at the line break.
        if len(autores_txt) > 1:
            lineas_p0 = [" ".join(l.split()) for l in pg0.get_text().splitlines() if l.strip()]
            i_ult = next((i for i, l in enumerate(lineas_p0) if l.endswith(" ".join(str(autores_txt[-1]).split()))), None)
            i_pre = next((i for i, l in enumerate(lineas_p0) if l.endswith(" ".join(str(autores_txt[-2]).split()))), None)
            if i_ult is None:
                R.anota("The members are in a single paragraph", False,
                        "the last name (%s) was not found on the cover page" % autores_txt[-1], critico=False)
            elif i_pre is not None and i_pre - i_ult > 1:
                R.anota("The members are in a single paragraph", False,
                        "there is a line between the second-to-last and the last member: it looks like one paragraph per person",
                        critico=False)
            else:
                R.anota("The members are in a single paragraph", True,
                        "%d members in a single paragraph" % len(autores_txt))

        # Logo: consistency between what the manifest asks for and what was drawn.
        logo = portada.get("logo")
        if logo:
            if not os.path.exists(logo):
                R.anota("The cover page logo exists", False,
                        "the manifest points to a nonexistent file: %s" % logo, critico=False)
            else:
                n_img = len(pg0.get_images())
                R.anota("The cover page logo exists", n_img > 0,
                        "file present" if n_img else "the file exists but the cover page shows no images",
                        critico=False)
        else:
            R.anota("The cover page logo was asked about", False,
                    "there is no logo in the manifest; it is optional, but it must always be asked about before omitting it",
                    critico=False)

        # The instructor's title or profession is always asked about.
        if portada.get("docente") and not portada.get("docente_titulo"):
            R.anota("The instructor's title was asked about", False,
                    "docente_titulo empty: their title or profession must be asked about", critico=False)

    # --- 5. Sections ----------------------------------------------------
    faltantes = []
    for s in M.get("secciones", []):
        if not paginas_con(paginas, norm(s["texto"]), desde=CUERPO):
            faltantes.append(s["texto"][:60])
    R.anota("All manifest sections appear in the PDF", not faltantes,
            "%d of %d missing (searching from page %d)%s"
            % (len(faltantes), len(M.get("secciones", [])), CUERPO + 1,
               (": " + "; ".join(faltantes[:5])) if faltantes else ""))

    # --- 6. Tables -------------------------------------------------------
    problemas_tablas = []
    for t in M.get("tablas", []):
        if ("tabla %d" % t["indice"]) not in norm_total:
            problemas_tablas.append("'Tabla %d' does not appear" % t["indice"])
            continue
        # content: at least the first cell of the first data row
        celdas = [c for fila in t["filas"] for c in fila if c.strip()]
        if celdas:
            muestra = norm(celdas[0])[:40]
            if muestra and muestra not in norm_total:
                problemas_tablas.append("Tabla %d: the content '%s' does not appear" % (t["indice"], muestra))
    R.anota("All tables appear with their number and content", not problemas_tablas,
            "%d table(s); %s" % (len(M.get("tablas", [])),
                                 "; ".join(problemas_tablas[:4]) if problemas_tablas else "all present"))

    # --- 7. Figures ------------------------------------------------------
    problemas_fig = []
    figs_omitidas = []
    figs_esperadas = []
    for f in M.get("figuras", []):
        if not f.get("existe"):
            figs_omitidas.append(f["indice"])
            continue
        figs_esperadas.append(f)
        if ("figura %d" % f["indice"]) not in norm_total:
            problemas_fig.append("'Figura %d' does not appear" % f["indice"])
        if f.get("titulo") and norm(f["titulo"])[:40] not in norm_total:
            problemas_fig.append("Figura %d: the caption does not appear" % f["indice"])
    R.anota("All figures appear with their caption",
            not problemas_fig if figs_esperadas else True,
            "%d figure(s) inserted; %s" % (
                len(figs_esperadas),
                "; ".join(problemas_fig[:4]) if problemas_fig else "all present"),
            critico=bool(figs_esperadas))
    # A figure without a file is not a PDF failure: the pipeline omits it and
    # warns about it, so it is reported separately without failing the check.
    if figs_omitidas:
        R.anota("Figures without a file in the source", False,
                "%d: %s" % (len(figs_omitidas),
                            ", ".join("Figura %d" % i for i in figs_omitidas)
                            + " were not inserted; the .md points them to a nonexistent "
                              "file"),
                critico=False)

    # --- 8. References --------------------------------------------------
    refs = M.get("referencias", [])
    faltan_refs = []
    for r in refs:
        # compare using a distinctive fragment: the first 60 characters
        frag = norm(r["texto"])[:60]
        if frag and frag not in norm_total:
            faltan_refs.append(r["texto"][:55])
    R.anota("All the .md references are in the PDF", not faltan_refs,
            "%d reference(s); %d missing%s" % (len(refs), len(faltan_refs),
                                               (": " + "; ".join(faltan_refs[:3])) if faltan_refs else ""))

    # references index sorted alphabetically
    if len(refs) > 1:
        # The references section page is the BODY one, not the TOC's (which also
        # contains the word "Referencias").
        pag_ref = None
        for i in range(CUERPO, total):
            if "referencias" in norm(paginas[i]):
                pag_ref = i
                break
        if pag_ref is not None:
            pag = norm(texto_en_orden_visual(doc, pag_ref, total))
            pos = [pag.find(norm(r["texto"])[:30]) for r in refs]
            # Careful: -1 means "not found", and -1 is truthy in Python.
            todos = all(p >= 0 for p in pos)
            if todos:
                # The order must be judged on the PDF's VISUAL order, not on the
                # manifest's: comparing `pos` against `sorted(pos)` in manifest
                # order would always report OUT OF ORDER when the .md did not
                # come sorted, which is the normal case.
                visual = [refs[k]["texto"] for k in sorted(range(len(pos)), key=lambda i: pos[i])]
                desorden = [visual[i][:40] for i in range(1, len(visual))
                            if visual[i].lower() < visual[i - 1].lower()]
                R.anota("The references are in alphabetical order", not desorden,
                        ("correct visual order: %s" % " | ".join(t[:22] for t in visual))
                        if not desorden
                        else "%d out of order: %s" % (len(desorden), "; ".join(desorden[:3])),
                        critico=False)
            else:
                R.anota("The references are in alphabetical order", False,
                        "not verifiable: %d of %d references were not found in the section"
                        % (sum(1 for p in pos if p < 0), len(refs)),
                        critico=False)
        else:
            R.anota("The references are in alphabetical order", False,
                    "the references section page was not found", critico=False)

    # --- 9. Table of contents consistency -------------------------------
    # Empty-title entries are continuations of a TOC line that was split in two
    # (long headings wrap). They are not entries.
    toc = [f for f in _toc if f["titulo"].strip()]
    if not toc and (M.get("opciones", {}).get("indice_tablas") or M.get("secciones")):
        R.anota("The table of contents has entries", False,
                "no line with a dot leader was found", critico=False)
    else:
        R.anota("The table of contents has entries", bool(toc),
                "%d entries with a page number" % len(toc), critico=False)

    # REPEATED headings (e. g. "Fortalezas, Limitaciones y Aplicabilidad"
    # four times) cannot be checked by matching one to one: the k-th occurrence
    # of the TOC has to be matched with the k-th one of the body.
    reclamadas = {}
    for fila in toc:
        reclamadas.setdefault(norm(fila["titulo"]), []).append(fila["pagina"])

    desajustes = []
    for titulo_norm, pedidas in reclamadas.items():
        # A table or figure index entry reads "Tabla 1. (sin titulo)", but in
        # the body only "Tabla 1" is printed: the full title does not exist
        # there and searching for it would give a false negative. For those
        # entries only the label is checked.
        etiqueta = RE_ETIQUETA_INDICE.match(titulo_norm)
        busca = etiqueta.group(0).strip() if etiqueta else titulo_norm
        # If the entry was generated with a missing-title placeholder (for
        # example "(sin titulo)"), it must not be tolerated: the process should
        # have aborted before generating the DOCX. It is flagged as a critical
        # failure to detect that case.
        if "(sin titulo)" in titulo_norm or "(sin leyenda)" in titulo_norm:
            desajustes.append("'%s' contains '(sin titulo/leyenda)' in the TOC" % titulo_norm[:40])
            continue
        encontradas = [p + 1 for p in paginas_con(paginas, busca, desde=CUERPO)]
        if not encontradas:
            desajustes.append("'%s': not found in the body" % titulo_norm[:40])
            continue
        if len(encontradas) < len(pedidas):
            desajustes.append("'%s': the TOC lists it %d times and the body has it %d"
                              % (titulo_norm[:40], len(pedidas), len(encontradas)))
            continue
        for k, pedida in enumerate(pedidas):
            if encontradas[k] == pedida:
                continue
            # A heading may show up in passing in the body (a mention inside a
            # paragraph) BEFORE the real section starts. In that case the
            # expected page is still correct and the position-based match must
            # not flag a failure: it is enough that the page stated by the TOC
            # is among the pages where the heading appears.
            if pedida in encontradas[k:]:
                continue
            desajustes.append("'%s' (occurrence %d): the TOC says %d but it is on page %d"
                              % (titulo_norm[:40], k + 1, pedida, encontradas[k]))
    R.anota("The TOC numbers match the real pages", not desajustes,
            ("%d mismatch(es): %s" % (len(desajustes), "; ".join(desajustes[:4])))
            if desajustes else "%d entries checked (%d distinct headings)"
            % (len(toc), len(reclamadas)))

    # --- 10. Broken bookmarks -------------------------------------------
    rotas = []
    for patron in (r"error!\s*bookmark", r"error!\s*reference", r"\?\?\s*$"):
        m = re.search(patron, texto_total, re.I | re.M)
        if m:
            rotas.append(patron)
    # a lone 0 or dash after the dot leader is also a broken PAGEREF
    for linea in texto_total.splitlines():
        if re.search(r"\.{4,}\s*(0|-|\?)\s*$", linea.strip()):
            rotas.append("TOC with empty/0 page number")
            break
    R.anota("There are no broken bookmarks or page references", not rotas,
            "; ".join(sorted(set(rotas))) if rotas else "none", critico=False)

    # --- 11. Typography --------------------------------------------------
    # PDFs embed the fonts with a subset prefix ("BAAAAA+"), so it must be
    # removed before comparing the name.
    #
    # Times New Roman is not installed with the operating system, so LibreOffice
    # substitutes a metrically compatible font wherever it is missing -- which on
    # Linux and macOS is the normal case, not an error. Those substitutes have
    # the same advance widths, so the layout is identical and they are accepted;
    # they are still reported so the reader knows what the machine really used.
    # Anything else stays a failure, because with different metrics the line
    # advance measured in check 12 no longer means anything.
    fuentes = set()
    for i in range(total):
        for f in doc[i].get_fonts(full=False):
            fuentes.add(f[3] if len(f) > 3 else str(f))

    declaradas, sustituciones, rechazadas = fuentes_mod.particiona(sorted(fuentes))
    detalle = "fonts: %s" % (", ".join(fuentes) if fuentes else "none")
    if sustituciones:
        detalle += ("  -> metric-compatible substitute for Times New Roman: %s"
                    % ", ".join(sustituciones))
    if rechazadas:
        detalle += "  -> NOT accepted: " + "; ".join(
            "%s (%s)" % (nombre, motivo) for nombre, motivo in rechazadas)

    R.anota("Only the declared typeface is used", not rechazadas, detalle)

    # --- 12. Line spacing ------------------------------------------------
    # pymupdf returns each LINE as a separate block, not the paragraph, so
    # "box height / number of lines" cannot be measured.
    #
    # The line ADVANCE is measured instead: the difference between the top edge of
    # two consecutive lines of the same column. Times New Roman 12 advances ~13.8 pt
    # at single spacing and ~27.6 pt at double spacing.
    #   Careful: measuring the gap between boxes (y0_next - y1_previous) does not
    #   work: at double spacing that gap is also ~13.8 pt (double spacing pushes the
    #   box, it does not make it taller), and it would give a false negative.
    UMBRAL_DOUBLE = 20.0
    avances = []
    for i in range(CUERPO, min(total, CUERPO + 15)):
        bs = [b for b in bloques[i] if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
        bs.sort(key=lambda b: b[1])
        for a, b in zip(bs, bs[1:]):
            if abs(a[0] - b[0]) < 3 and 0 < (b[1] - a[1]) < 45:
                avances.append(b[1] - a[1])
    if avances:
        avances.sort()
        mediana = avances[len(avances) // 2]
        R.anota("Double line spacing in the body", mediana >= UMBRAL_DOUBLE,
                "median line advance = %.1f pt (double ~27,6 | single ~13,8)"
                % mediana, critico=False)
    else:
        R.anota("Double line spacing in the body", False,
                "consecutive lines could not be paired to measure", critico=False)

    # --- 13. Table and figure notes -------------------------------------
    # Two levels of checking, because they are not the same thing:
    #   a) that the manifest declares a note, and
    #   b) that that note really appears in the PDF.
    # Only (a) would be trivial (the parser fills it in by default) and would let a
    # document delivered without notes pass.
    Mobj = {}
    try:
        with open(args.manifiesto, "r", encoding="utf-8") as fh:
            Mobj = json.load(fh)
    except Exception:
        Mobj = {}

    # Global reading flow excluding the index pages: in the table of tables each
    # entry reads "Tabla N. titulo" and would generate false positives.
    paginas_indice = set()
    for i in range(doc.page_count):
        t = paginas[i]
        if re.search(r"\bIndice de (tablas|figuras)\b", t, re.I) and re.search(r"\.{4,}", t):
            paginas_indice.add(i)

    flujo = []
    for i in range(doc.page_count):
        if i in paginas_indice:
            continue
        for b in bloques[i]:
            if len(b) >= 5 and isinstance(b[4], str) and b[4].strip():
                flujo.append((i + 1, round(b[1], 1), b[4].strip()))
    flujo.sort(key=lambda x: (x[0], x[1]))

    sin_nota = []
    sin_nota_en_pdf = []
    pos = 0
    for t in Mobj.get("tablas", []):
        etiqueta = "Tabla %d" % t["indice"]
        nota = t.get("nota")
        if not nota:
            sin_nota.append(etiqueta)
            continue
        # The note must appear AFTER the label of its table in the flow and
        # BEFORE the next table: that way it stays below the table (APA 7).
        ini = None
        for k in range(pos, len(flujo)):
            if re.match(r"^%s\b" % re.escape(etiqueta), flujo[k][2]):
                ini = k
                break
        if ini is None:
            sin_nota_en_pdf.append("%s: the label was not found" % etiqueta)
            continue
        needle = norm(nota)[:40]
        enc = None
        for k in range(ini, min(len(flujo), ini + 60)):
            if re.match(r"^Nota\.\s", flujo[k][2]) and norm(flujo[k][2]).find(needle) >= 0:
                enc = k
                break
        if enc is None:
            sin_nota_en_pdf.append("%s: the note does not appear below the table" % etiqueta)
        else:
            pos = enc + 1
    R.anota("Tables carry a note below (in the PDF, not only in the manifest)",
            not (sin_nota or sin_nota_en_pdf),
            "; ".join(sin_nota + sin_nota_en_pdf) if (sin_nota or sin_nota_en_pdf)
            else "%d table(s) with note verified below the table" % len(Mobj.get("tablas", [])))

    # Figures are checked geometrically further down (note above the image).
    # Here we only warn if the manifest declares no note.
    fig_sin_nota = [f["indice"] for f in Mobj.get("figuras", []) if not f.get("nota")]
    if fig_sin_nota:
        R.anota("The manifest declares a note for each figure", False,
                "no note: Figura %s" % ", ".join(str(i) for i in fig_sin_nota))
    else:
        R.anota("The manifest declares a note for each figure", True,
                "%d figure(s)" % len(Mobj.get("figuras", [])))

    # --- 14. Figure note position (above the image) ----------
    # Searching the note by text does NOT work: several figures usually share the
    # same default note ("Elaboracion propia"), so a text search always returns
    # the first one and flags all the others as errors. The GEOMETRY must be
    # compared: the general note of the figure is the "Nota." text block that
    # lies ABOVE (lower y) the image, on the same page.
    sin_nota_encima = []
    paginas_con_foto = 0
    for i in range(CUERPO, total):
        rects_imagen = []
        for img in doc[i].get_images(full=True):
            rects_imagen.extend(doc[i].get_image_rects(img[0]))
        if not rects_imagen:
            continue
        paginas_con_foto += 1
        # Text blocks on the page that start with "Nota.". Careful: the RAW text
        # must be inspected, not norm(), because norm() strips the punctuation and
        # "Nota." would become "nota" (and the prefix with the dot would never
        # match).
        bs = [b for b in bloques[i] if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
        notas = [b for b in bs if re.match(r"^Nota\.\s", b[4].strip())]
        for r in rects_imagen:
            # The bottom edge of the note block must lie above (or exactly at
            # the top edge of) the image.
            if not [b for b in notas if b[3] <= r.y0 + 4]:
                sin_nota_encima.append("p. %d: image at y0=%.0f with no note above"
                                       % (i + 1, r.y0))

    if paginas_con_foto == 0:
        R.anota("Figure notes go above the image", True,
                "the document has no images: not applicable", critico=False)
    elif sin_nota_encima:
        R.anota("Figure notes go above the image", False,
                "%d image(s) with no note above: %s"
                % (len(sin_nota_encima), "; ".join(sin_nota_encima[:3])), critico=False)
    else:
        R.anota("Figure notes go above the image", True,
                "%d page(s) with an image, all with a note above" % paginas_con_foto)

    return R, total


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Verifies the generated PDF against its manifest")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--manifiesto", required=True)
    ap.add_argument("--log")
    ap.add_argument("--json")
    args = ap.parse_args()

    lineas = []

    def w(s=""):
        print(s)
        lineas.append(s)

    if not os.path.isfile(args.pdf):
        w("FAIL: the PDF does not exist: %s" % args.pdf)
        return 1

    w("=== PHASE 4: PDF verification ===")
    w("[%s] PDF        : %s" % (datetime.now().strftime("%H:%M:%S"), args.pdf))
    w("[%s] Manifest   : %s" % (datetime.now().strftime("%H:%M:%S"), args.manifiesto))
    w("")

    R, total = verificar(args)

    for it in R.items:
        w("  %-6s %-52s %s" % ("OK" if it["ok"] else ("FAIL" if it["critico"] else "WARN"),
                                it["nombre"], it["detalle"]))
    w("")

    criticas = len(R.fallas)
    no_criticas = len(R.avisos)
    w("Checks: %d   OK: %d   FAILURES: %d   WARNINGS: %d"
      % (len(R.items), len(R.items) - criticas - no_criticas, criticas, no_criticas))
    w("Pages: %d" % total)
    w("")
    if criticas:
        w("RESULT: FAIL (%d critical check(s))" % criticas)
    else:
        w("RESULT: OK")

    if args.log:
        d = os.path.dirname(os.path.abspath(args.log))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(args.log, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lineas) + "\n")
        w("Log written: %s" % args.log)

    if args.json:
        d = os.path.dirname(os.path.abspath(args.json))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({
                "pdf": os.path.abspath(args.pdf),
                "paginas": total,
                "resultado": "FAIL" if criticas else "OK",
                "fallas_criticas": criticas,
                "avisos": no_criticas,
                "comprobaciones": R.items,
            }, fh, ensure_ascii=False, indent=2)
        w("JSON report: %s" % args.json)

    return 1 if criticas else 0


if __name__ == "__main__":
    sys.exit(main())
