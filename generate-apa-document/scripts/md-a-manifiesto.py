#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md-a-manifiesto.py - Converts a .md (and its layout JSONs) into MANIFEST.json

This script is the BRAIN of the pipeline. Before, every document forced you to
rewrite a .js generator of hundreds of lines with everything hardcoded (cover,
titles, tables, references). Now the .md is analyzed ONCE, a declarative
manifest is produced, and build-docx.js only applies it.

What it fixes compared to the previous pipeline:
  * References ALWAYS come from the .md. Before they were ignored completely and
    5 fixed literal references from the sample document were emitted.
  * The correction of "glued words" (deglue) is really implemented. Before it was
    only documented and never executed.
  * Tables, figures and references are inventoried, and it is DECLARED when
    something is missing, so that build-docx.js applies the conditional index policy.
  * U+2022 bullets and Docling's "o " artifact are detected, as well as the
    duplicated numbering "1. 1.".

Standard library only: it requires no pip install of anything.

Typical usage:
    python md-a-manifiesto.py --md doc.md --out salida/MANIFEST.json \
        --log salida/_logs/01-analisis.log \
        [--portada portada.json] [--ancho-pagina-landscape]

Agent answers (JSON {index: text}, 1-based index):
    --titulos-tabla-json   title of each table, if the .md does not bring it
    --titulos-figura-json  caption of each figure, if the .md does not bring it
    --notas-tabla-json      note of each table; default "Elaboracion propia"
    --notas-figura-json     note of each figure; default "Elaboracion propia"

    They are passed only for the missing indexes. They serve to answer the
    blocking questions that the parser itself leaves in
    MANIFEST.json -> diagnostico.preguntas: if the user chooses that the title be
    drafted from the context, it is written here and the parser is run again.
    build-docx.js does NOT accept these files.

Analysis options:
    --sin-deglue            do not apply the "glued words" correction
    --detectar-niveles-en-linea
                            treat "**Titulo.** text" as level 4 and
                            "***Titulo.*** text" as level 5 (APA 7 inline).
                            Opt-in: bold at the start of a paragraph is also
                            normal text, and enabling it without control would
                            create fake titles.

Output: MANIFEST.json + a readable analysis report. The pending blocking
questions stay in diagnostico.pendientes_bloqueantes; missing titles or captions
make build-docx.js abort instead of inventing them.
"""

import argparse
import hashlib
import html
import json
import os
import re
import sys
import unicodedata
from datetime import datetime, timezone

VERSION_MANIFIESTO = 1

# ---------------------------------------------------------------------------
# Bullets and symbols that the conversions leave glued to the start of a line.
BULLET_CHARS = "\u2022\u25cf\u25aa\u25e6\u2043\u2219\u00b7\u2023\u2027"

# ---------------------------------------------------------------------------
# Text utilities
# ---------------------------------------------------------------------------


def norm_key(text):
    """Lowercase without accents, to compare titles in a tolerant way."""
    t = unicodedata.normalize("NFD", text.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def read_text(path):
    """Reads text detecting UTF-8/BOM. The .md files from MinerU and Docling are UTF-8."""
    with open(path, "rb") as fh:
        raw = fh.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        for enc in ("cp1252", "latin-1"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                continue
    return raw.decode("utf-8", errors="replace")


def read_json(path):
    """Loads an input JSON tolerating a BOM.

    It is better to use read_text instead of open(..., encoding='utf-8'): on Windows,
    Set-Content -Encoding UTF8 and most editors save with a BOM, and
    json.load aborts with 'Unexpected UTF-8 BOM' if the file carries it.
    """
    return json.loads(read_text(path))


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def unescape_md(text):
    """Removes the markdown escaping: 'A00\\_2021' -> 'A00_2021'."""
    return re.sub(r"\\([_\\*`~\[\]()#+\-.!])", r"\1", text)


def strip_leading_markers(text):
    """
    Removes conversion bullets at the start of the line.

    Real cases measured in Docling documents:
      '\u2022 Objetivos Especificos:'  -> 'Objetivos Especificos:'
      '\u2022 1. Entrenamiento (Training): ...' -> '1. Entrenamiento ...'
      'o Descripcion Tecnica: ...'       -> 'Descripcion Tecnica: ...'
    The 'o ' pattern is only removed if it is NOT a legitimate word: we require that
    after the marker comes an uppercase letter or a number, and that the rest is long.
    """
    original = text
    # 1) Unicode bullets (possibly repeated: '\u2022 \u2022 text')
    text = re.sub(r"^[\s" + BULLET_CHARS + r"]+[\s]*", "", text)
    # 2) Repeated ascii markers: '- - text', '* * text', '\u2013 \u2022 text'
    text = re.sub(r"^[\s\-\*\u2013\u2014]{2,}", "", text)
    # 3) Docling's 'o ' artifact (second-level bullet converted to a letter)
    m = re.match(r"^[oO][\s\t]+(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]|\d)", text)
    if m and len(text) > 3:
        text = text[m.end():]
    return text if text.strip() else original


def dedupe_numbering(text):
    """
    Normalizes list numbering duplicated by an extraction artifact.

    '1. 1. Analizar ...' -> '1. Analizar ...'
    '\u2022 1. 1. texto'  -> '\u2022 1. texto'   (the prefix was removed before)
    The first number is kept and the repeated one is discarded.
    """
    text = re.sub(r"^(\d+)\.\s+\d+\.\s+", r"\1. ", text)
    text = re.sub(r"^(\d+)\)\s+\d+\)\s+", r"\1) ", text)
    return text


def fix_colon_spacing(text):
    """'ACTIVIDAD 4 : Title' -> 'ACTIVIDAD 4: Title'  (space before ':')."""
    return re.sub(r"[ \t]+([:;])", r"\1", text)


# ---------------------------------------------------------------------------
# deglue: split words glued together by the conversion
# ---------------------------------------------------------------------------

GLUE_RE = re.compile(r"([a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0])([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc\u00c0\u00c8][a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0]{0,})")
# 'ACTIVIDAD1:' / 'CR1.4' -> number glued to an UPPERCASE word.
# The (?![-\u2013]\d) guard avoids splitting hyphenated identifiers: 'CVE2021-44228'
# must stay whole, not become 'CVE 2021-44228'.
GLUE_NUM_RE = re.compile(r"\b([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{3,})(\d+)\b(?![-\u2013]\d)")
# 'SegunelMarco' is already covered by GLUE_RE; ':Diferenciacion' covers the deglue of ':'
COLON_RE = re.compile(r"(:)(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc])")
# 'Digital(Portal)' -> 'Digital (Portal)': the space before the parenthesis is missing.
# >=4 characters inside are required so abbreviations like '(Ed.', '(n. 2)' are left alone.
PAREN_RE = re.compile(r"([a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0])\(([^()]{4,})\)")
# 'SCAySBOM' -> 'SCA y SBOM': acronym, one-letter conjunction, another acronym.
# This is the most common acronym gluing in reports and GLUE_RE does not see it, because
# it requires lowercase followed by uppercase and here there is only one lowercase in between.
GLUE_SIGLA_Y_RE = re.compile(
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{3,})"
    r"([a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0])"
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{2,})")
# 'OWASPTop10' -> 'OWASP Top10': uppercase acronym glued to a capitalized word.
GLUE_ACR_RE = re.compile(
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{3,})"
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc][a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0]+)")
# 'Top10' -> 'Top 10', 'Junio2021' -> 'Junio 2021': word glued to a number.
# This rule also splits legitimate mixed alphanumerics (IPv6, SHA256, Base64):
# nothing in the pipeline keeps a list of them, and a term list would have to be
# fed per document. The price is accepted on purpose: the residue report shows
# every change so the term can be corrected in the source .md, which is where it
# belongs.
GLUE_PALABRA_NUM_RE = re.compile(
    r"([A-Za-z\u00c0-\u00ff]{3,})(\d{1,4})(?![0-9A-Za-z-])")
# '10de' -> '10 de', '2Capa' -> '2 Capa', '3Mes5' -> '3 Mes 5': number glued to
# the following word. Only capitalized words or Spanish function words are
# accepted, so as not to split '3er' nor 'v2beta'.
GLUE_NUM_PALABRA_RE = re.compile(
    r"(\d{1,4})(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc][a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0]"
    r"|(?:de|del|la|el|los|las|un|una|y|en|con|para|por)\b)")


# Structural protection: before any rule runs, URLs, DOIs, e-mails and file
# names with an extension are set aside behind a sentinel and put back at the
# end, byte for byte. A URL query string ('?v=dQw4w9WgXcQ') must never be
# touched: splitting it corrupts the link. These never show up in the change
# report because there was no change.
RE_MASCARA = re.compile(
    r"https?://[^\s<>()\"']+"
    r"|www\.[^\s<>()\"']+"
    r"|\b10\.\d{4,9}/[^\s<>()\"']+"
    r"|\b[\w.+-]+@[\w-]+\.[\w.-]+\b"
    r"|\b[\w./\\-]+\.(?:md|txt|json|png|jpg|jpeg|gif|bmp|pdf|docx|xlsx|pptx"
    r"|csv|tsv|zip|js|py|html|css|xml|yml|yaml|sql|sh|bat)\b",
    re.I,
)

# Terms the document marks as untouchable (one per line, via
# --terminos-protegidos). They are matched as a full token, case-sensitively,
# and masked like the structural patterns. They are set from main().
TERMINOS_PROTEGIDOS = []


def _mascara(text):
    """(text_with_sentinels, saved_fragments) for URLs/DOIs/e-mails/files/terms."""
    spans = [(m.start(), m.end()) for m in RE_MASCARA.finditer(text)]
    for term in TERMINOS_PROTEGIDOS:
        if not term:
            continue
        patron = re.compile(r"(?<![0-9A-Za-z_])" + re.escape(term) + r"(?![0-9A-Za-z_])")
        for m in patron.finditer(text):
            spans.append((m.start(), m.end()))
    if not spans:
        return text, []
    spans.sort()
    fusion = []
    for s, e in spans:
        if fusion and s <= fusion[-1][1]:
            fusion[-1] = (fusion[-1][0], max(fusion[-1][1], e))
        else:
            fusion.append((s, e))
    guardados = []
    partes = []
    ult = 0
    for s, e in fusion:
        partes.append(text[ult:s])
        guardados.append(text[s:e])
        # The sentinel uses NUL bytes on purpose: NUL is in none of the rule
        # character classes, so no rule can match or alter it.
        partes.append("\x00%d\x00" % (len(guardados) - 1))
        ult = e
    partes.append(text[ult:])
    return "".join(partes), guardados


def _desmascara(text, guardados):
    if not guardados:
        return text
    return re.sub(r"\x00(\d+)\x00", lambda m: guardados[int(m.group(1))], text)


def _ambiguo(token):
    """True when a split could be a legitimate term rather than a glued error.

    Mixed alphanumerics ('SHA256', 'IPv6', 'Base64') and camelCase
    ('JavaScript', 'GitHub') land here: they are the cases that must be
    reviewed before accepting the correction, because they may well be the
    real spelling of a product or a variable.
    """
    tiene_letra = bool(re.search(r"[A-Za-z\u00c0-\u00ff]", token))
    tiene_num = bool(re.search(r"\d", token))
    camel = bool(re.search(
        r"[a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0][A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]", token))
    return (tiene_letra and tiene_num) or camel


def deglue(text, apply=True):
    """
    Inserts spaces in words glued together by the PDF-to-text conversion.

    Real cases from the reference document:
      'SegunelMarcoBSIMM'                  -> 'Segun el Marco BSIMM'
      'modeladodeamenazas'                 -> 'modelado de amenazas'
      'ACTIVIDAD4'                         -> 'ACTIVIDAD 4'
      '2:Diferenciacion'                   -> '2: Diferenciacion'

    Protected before the rules run and restored unchanged afterwards: URLs,
    DOIs, e-mails, file names with an extension, and every term listed in the
    protected-terms file. They never appear in the change report.

    Pure UPPERCASE acronyms ('OWASP', 'CVSS') are not split: the rule requires
    lowercase followed by uppercase.

    Every change is reported with a flag: `True` when it is AMBIGUOUS (mixed
    alphanumeric or camelCase, i.e. it could be a product name or a variable),
    `False` when it is a clear correction. The pipeline does NOT decide: the
    agent reads the report and either keeps the correction, adds the term to
    the protected list and re-runs, or asks the user.

    Returns (corrected_text, list_of_changes), each change being
    (kind, before, after, ambiguous).
    """
    cambios = []
    if not apply or not text:
        return text, cambios

    # Set the protected fragments aside so no rule can touch them.
    text, guardados = _mascara(text)

    def sub_glue(m):
        a, b = m.group(1), m.group(2)
        cambios.append(("glue", m.group(0), a + " " + b, _ambiguo(m.group(0))))
        return a + " " + b

    # The order matters. GLUE_RE goes first because it sees 'SegunelMarco'; but it splits
    # 'SCAySBOM' into 'SCAy SBOM' and afterwards there is nothing left that an acronym
    # rule could fix, so the acronym rules are applied BEFORE GLUE_RE. That is why they
    # are all evaluated against 'text' and chained in the inverse order
    # of application.
    def sub_sigla_y(m):
        cambios.append(("sigla_y", m.group(0), m.group(1) + " " + m.group(2) + " " + m.group(3),
                        _ambiguo(m.group(0))))
        return m.group(1) + " " + m.group(2) + " " + m.group(3)

    def sub_acr(m):
        cambios.append(("sigla", m.group(0), m.group(1) + " " + m.group(2),
                        _ambiguo(m.group(0))))
        return m.group(1) + " " + m.group(2)

    def sub_num_palabra(m):
        cambios.append(("num_palabra", m.group(0), m.group(0) + " ",
                        _ambiguo(m.group(0))))
        return m.group(0) + " "

    def sub_num(m):
        cambios.append(("glue_num", m.group(0), m.group(1) + " " + m.group(2),
                        _ambiguo(m.group(0))))
        return m.group(1) + " " + m.group(2)

    def sub_palabra_num(m):
        cambios.append(("palabra_num", m.group(0), m.group(1) + " " + m.group(2),
                        _ambiguo(m.group(0))))
        return m.group(1) + " " + m.group(2)

    out = GLUE_NUM_PALABRA_RE.sub(sub_num_palabra, text)
    out = GLUE_PALABRA_NUM_RE.sub(sub_palabra_num, out)
    out = GLUE_NUM_RE.sub(sub_num, out)
    out = GLUE_SIGLA_Y_RE.sub(sub_sigla_y, out)
    out = GLUE_ACR_RE.sub(sub_acr, out)
    out = GLUE_RE.sub(sub_glue, out)

    def sub_colon(m):
        cambios.append(("colon", m.group(0), m.group(0) + " ", False))
        return m.group(0) + " "

    out = COLON_RE.sub(sub_colon, out)

    def sub_paren(m):
        cambios.append(("paren", m.group(0), m.group(1) + " (" + m.group(2) + ")",
                        _ambiguo(m.group(0))))
        return m.group(1) + " (" + m.group(2) + ")"

    out = PAREN_RE.sub(sub_paren, out)
    out = _desmascara(out, guardados)
    return out, cambios


# ---------------------------------------------------------------------------
# Inline formatting -> segments (bold / italic)
# ---------------------------------------------------------------------------

INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*[^*\n]+?\*|__.+?__|_[^_\n]+?_)")


def parse_inline(text):
    """
    Converts 'text **bold** and *italic*' into a list of segments.
    Returns [{'text':..., 'bold':bool, 'italics':bool}, ...]
    """
    segs = []
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            segs.append({"text": text[pos:m.start()], "bold": False, "italics": False})
        tok = m.group(0)
        if tok.startswith("**") and tok.endswith("**"):
            segs.append({"text": tok[2:-2], "bold": True, "italics": False})
        elif tok.startswith("__") and tok.endswith("__"):
            segs.append({"text": tok[2:-2], "bold": True, "italics": False})
        elif tok.startswith("*") and tok.endswith("*"):
            segs.append({"text": tok[1:-1], "bold": False, "italics": True})
        else:
            segs.append({"text": tok[1:-1], "bold": False, "italics": True})
        pos = m.end()
    if pos < len(text):
        segs.append({"text": text[pos:], "bold": False, "italics": False})
    return [s for s in segs if s["text"] != ""]


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

HTML_TABLE_RE = re.compile(r"<table[^>]*>.*?</table>", re.I | re.S)
TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.I | re.S)
TD_RE = re.compile(r"<(td|th)[^>]*>(.*?)</\1>", re.I | re.S)
BR_RE = re.compile(r"<br\s*/?>", re.I)
TAG_RE = re.compile(r"<[^>]+>")


def cell_to_text(raw):
    txt = BR_RE.sub(" \u21b5 ", raw)   # line break inside the cell
    txt = TAG_RE.sub("", txt)
    txt = html.unescape(html.unescape(txt))
    txt = unescape_md(txt)
    txt = re.sub(r"[ \t]+", " ", txt)
    return txt.strip()


def parse_html_table(block):
    filas = []
    for tr in TR_RE.findall(block):
        celdas = [cell_to_text(c) for _, c in TD_RE.findall(tr)]
        if any(c for c in celdas):
            filas.append(celdas)
    if not filas:
        return []
    ncols = max(len(r) for r in filas)
    filas = [r + [""] * (ncols - len(r)) for r in filas]
    return filas


def parse_md_table(lines, start):
    """Detects a markdown table |a|b| ... Returns (rows, end_index)."""
    filas = []
    i = start
    while i < len(lines):
        ln = lines[i].strip()
        if not ln.startswith("|"):
            break
        celdas = [c.strip() for c in ln.strip("|").split("|")]
        if not all(re.fullmatch(r":?-{2,}:?", c) for c in celdas if c):
            filas.append([unescape_md(cell_to_text(c)) for c in celdas])
        i += 1
    if not filas:
        return [], start
    ncols = max(len(r) for r in filas)
    return [r + [""] * (ncols - len(r)) for r in filas], i


# ---------------------------------------------------------------------------
# References: detection of the italic range (heuristic + warnings)
# ---------------------------------------------------------------------------

YEAR_RE = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")
REVISTA_RE = re.compile(r"\b(revista|journal|revista|vol\.|volumen|n[o\u00ba]\.|pp\.|edition|ed\.)", re.I)
TESIS_RE = re.compile(r"\b(repositorio|tesis|trabajo de grado|universidad|universidad tecnica|monograf)", re.I)
URL_TAIL_RE = re.compile(r"((?:Recuperado de|Disponible en|consultado|Retrieved from|Available from)\s*:?\s*https?://\S+)\s*$", re.I)
BARE_URL_RE = re.compile(r"(https?://\S+)\s*$")


def propose_reference_italics(text):
    """
    PROPOSAL (not definitive) of which part of the reference goes in italics.

    APA 7 depends on whether the source is a book, a report, a journal article or
    a thesis, and that distinction cannot be determined reliably from the plain
    text alone that the conversion produces. Therefore:

      1. a conservative heuristic is applied, and
      2. a warning is returned so that the agent/user-agent reviews it.

    Rules:
      * Journal/periodical  -> italics on the journal name.
      * Repository/thesis   -> italics on the work title.
      * Web documentation   -> italics on the work title.
    """
    cuerpo = text
    m = URL_TAIL_RE.search(cuerpo)
    if m:
        cuerpo = cuerpo[: m.start()]
    else:
        m2 = BARE_URL_RE.search(cuerpo)
        if m2:
            cuerpo = cuerpo[: m2.start()]

    my = YEAR_RE.search(cuerpo)
    if not my:
        return None, "sin (anio) reconocible"
    ini = my.end()
    resto = cuerpo[ini:].lstrip(". ").strip()
    if not resto:
        return None, "sin titulo tras el anio"

    if REVISTA_RE.search(cuerpo):
        # The article title stays plain; the journal name goes in italics.
        mrev = REVISTA_RE.search(resto)
        if mrev:
            ini_rev = mrev.start()
            while ini_rev > 0 and resto[ini_rev - 1] not in ".,:":
                ini_rev -= 1
            fin_rev = mrev.end()
            cola = resto[fin_rev:]
            # Include the edition/volume parenthesis: '(Ed. 164)', '(Vol. 3, n. 2)'
            mpar = re.match(r"^[^()]{0,24}\([^()]*\)", cola)
            if mpar:
                fin_rev += mpar.end()
            else:
                msep = re.search(r"[,;]", cola)
                if msep:
                    fin_rev += msep.start()
            cand = resto[ini_rev:fin_rev].strip(" .,")
            if cand:
                return cand, "revista"
        return resto.split(".")[0], "articulo"

    # Book / report / thesis / documentation: italics on the work title.
    if TESIS_RE.search(cuerpo):
        # In a thesis the italic part spans title + institution.
        cortes = list(re.finditer(r"\.\s+(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da])", resto))
        for k, c in enumerate(cortes[:2]):
            fin = c.start() + 1
            if k == 1 or re.match(r"\s*(Universidad|Instituto|Corporaci\u00f3n|Instituto)", resto[c.end():c.end() + 16]):
                return resto[:fin].strip(), "tesis"
    titulo = resto.split(". ")[0]
    return titulo.strip(" .,"), "obra"


# ---------------------------------------------------------------------------
# Main analysis
# ---------------------------------------------------------------------------

RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
# Level 4 APA title ("**Texto.** the rest of the paragraph") or level 5
# ("***Texto.*** ..."). Group 1 says whether it is level 5 (triple asterisks),
# group 2 is the title text, which must end in a period.
RE_TITULO_EN_LINEA = re.compile(r"^(\*{3}|\*\*)(?!\s)(.+?)\.\1\s+(\S.*)$", re.S)
RE_IMAGEN = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
RE_NUM_HEADING = re.compile(r"^\s*(\d{1,2})[\.\)]\s+(.*)$")
RE_FIG_CAPTION = re.compile(r"^\s*(?:figura|imagen|image|ilustraci\u00f3n)\s*(\d+)\s*[.:\-\u2013]?\s*(.*)$", re.I)
RE_TAB_CAPTION = re.compile(r"^\s*(?:tabla|cuadro|table)\s*(\d+)\s*[.:\-\u2013]?\s*(.*)$", re.I)
# A note written in the .md next to a table or a figure: "Nota. Elaboracion
# propia" / "Note: own elaboration". The leading emphasis markers are optional
# because strip_leading_markers() already ate the opening "**" of "**Nota.**",
# which can leave a stray "**" right after the label.
RE_NOTA_LINEA = re.compile(r"^\s*\**\s*(?:nota|note)\s*[.:]\s*\**\s*(.+?)\s*$", re.I)
RE_NRC = re.compile(r"\bNRC\s*[:\s]*\s*([0-9]{4,}[-\u2013A-Za-z0-9]*)", re.I)
RE_DATE = re.compile(r"(\d{1,2})\s+de\s+([a-z\u00e1\u00e9\u00ed\u00f3\u00fa\u00f1]+)\s+de\s+(\d{4})", re.I)

CAMPOS_PORTADA = [
    ("materia_nrc", RE_NRC),
    ("docente", re.compile(r"profesor|profesora|docente|teacher", re.I)),
    ("facultad", re.compile(r"facultad\s+de", re.I)),
    ("vicerrectoria", re.compile(r"vicerrector", re.I)),
    ("fecha", RE_DATE),
]




# ---------------------------------------------------------------------------
# Enrichment from MinerU / Docling
# ---------------------------------------------------------------------------
# The real input contract is not just "a .md": MinerU and Docling deliver the
# prose .md PLUS a *_content_list.json with the structure (tables with their HTML,
# their bbox and their page) and an images/ folder with the crops of each table.
# The .md on its own loses exactly what APA demands be named: the title of each
# table. This block recovers that information instead of inventing it.
#
# What is NOT done here, on purpose:
#   - the document is not reordered with the json (the .md rules on ordering);
#   - tables that the .md does not bring are not added, because json entries with
#     0 rows are empty detections from Docling, not lost tables;
#   - a title is not deduced for a table that lacks one: it is left None and
#     asked about, like with the rest of the gaps.

RE_ETIQUETA_TABLA = re.compile(r"^\s*(tabla|cuadro|table)\s*[n\.#]?\s*\d*\s*[:.\-]?\s*",
                              re.I)

# Removes the number from a caption that already carries it in the .md, so that the .docx
# does not duplicate it by prepending its own ("Figura 1. Figura 1. ...").
RE_NUM_LEYENDA = re.compile(
    r"^\s*(figura|fig|imagen|table|tabla|cuadro|chart|grafico|grafico\s*\d*)\s*"
    r"(n[.°º]?\s*)?\d+\s*[:.\)\-–]\s*", re.I)


def _primer_texto_util(campos):
    """First non-empty string among the captions/footnotes in the json.

    CAREFUL with the shape of the data: `table_caption` and `table_footnote` are lists of
    strings, so `[e.get("table_footnote")]` would be a list inside a list and no element
    would be a string. It is flattened on purpose, because that is how it is
    read from the json and how it would look if a caption were added by hand.
    """
    for c in campos or []:
        if isinstance(c, str) and c.strip():
            return c.strip()
        if isinstance(c, (list, tuple)):
            for sub in c:
                if isinstance(sub, str) and sub.strip():
                    return sub.strip()
    return None


def limpiar_titulo_json(txt, aplicar_deglue):
    """Normalizes a title coming from the json (it comes with the same dirt as the
    rest of the .md: glued, with no space after ':' and sometimes with a label)."""
    t = unescape_md(strip_leading_markers(txt))
    if aplicar_deglue:
        t, _ = deglue(t, True)
    t = fix_colon_spacing(t).strip()
    t = RE_ETIQUETA_TABLA.sub("", t).strip()
    return t or None


def _firmas_tabla(filas):
    """Fingerprint of a table to match it with its json entry: the first
    complete row. Two different tables almost never share the first 3
    cells, and that datum survives even if the rest of the table was split."""
    if not filas or not filas[0]:
        return ""
    return norm_key(" ".join(filas[0][:3]))


def _entradas_tabla_json(cl):
    """Entries in content_list.json that are tables WITH content."""
    out = []
    for it in cl:
        if not isinstance(it, dict) or it.get("type") != "table":
            continue
        body = it.get("table_body") or ""
        if "<td" not in body:
            continue          # empty entry: detection without content
        out.append(it)
    return out


def resolver_tamano_pagina(md_path, docling_json, pdf, log):
    """(width_pt, height_pt, origin). The json bbox comes in points of the
    page, and to convert it to inches you must know how big the page is."""
    if pdf and os.path.isfile(pdf):
        try:
            import pymupdf
            d = pymupdf.open(pdf)
            r = d[0].rect
            log("Page size from the PDF: %.0f x %.0f pt" % (r.width, r.height))
            return r.width, r.height, "pdf"
        except Exception as e:                      # pragma: no cover
            log("WARNING  the PDF could not be read (%s); continuing with another method." % e)

    # Docling stores page_size per page in the *_middle.json
    base = docling_json or ""
    if base:
        guess = os.path.join(os.path.dirname(base), "_middle.json")
        cands = [guess]
        stem = os.path.basename(base).replace("_content_list.json", "")
        cands.append(os.path.join(os.path.dirname(base), stem + "_middle.json"))
        for c in cands:
            if c and os.path.isfile(c):
                try:
                    d = json.loads(read_text(c))
                    for pg in d.get("pdf_info", []):
                        ps = pg.get("page_size")
                        if ps and len(ps) == 2:
                            log("Page size from %s: %.0f x %.0f pt"
                                % (os.path.basename(c), ps[0], ps[1]))
                            return float(ps[0]), float(ps[1]), "middle.json"
                except Exception as e:              # pragma: no cover
                    log("WARNING  %s unreadable (%s)." % (os.path.basename(c), e))

    log("WARNING  there is no PDF nor middle.json: Letter 612x792 pt is assumed to "
        "convert the boxes. If the original was A4, the tables will come out a bit "
        "narrow; it can be forced by passing --pdf.")
    return 612.0, 792.0, "supuesto-carta"


def enriquecer_desde_docling(tablas, figuras, docling_json, base_dir, md_path,
                             pdf, ancho_max, log, aplicar_deglue=True):
    """Completes tables and figures with the MinerU/Docling json."""
    if not docling_json:
        return {"usado": False}

    if not os.path.isfile(docling_json):
        log("WARNING  the given json does not exist: %s" % docling_json)
        return {"usado": False}

    try:
        cl = json.loads(read_text(docling_json))
    except Exception as e:
        log("WARNING  %s unreadable: %s" % (os.path.basename(docling_json), e))
        return {"usado": False}

    log("Enrichment from %s (%d entries)"
        % (os.path.basename(docling_json), len(cl)))

    ancho_pt, alto_pt, origen_pagina = resolver_tamano_pagina(
        md_path, docling_json, pdf, log)

    espacio = _detectar_espacio_bbox(cl, ancho_pt, alto_pt)
    log("Coordinate space of the bbox: %s%s" % (
        espacio,
        " (normalized 0..1000; calibrated with page_size)" if espacio == "normalizado-1000" else ""))

    jt = _entradas_tabla_json(cl)
    log("Tables with content in the json: %d (out of %d 'table' entries)"
        % (len(jt), sum(1 for x in cl if isinstance(x, dict) and x.get("type") == "table")))

    # Matching by first-row fingerprint; if it does not fit, by order, which
    # is the right thing because both files come from the same PDF.
    por_firma = {}
    for e in jt:
        f = _firmas_tabla(parse_html_table(e.get("table_body") or ""))
        if f:
            por_firma.setdefault(f, []).append(e)
    usados = set()
    orphans = []

    for tb in tablas:
        f = _firmas_tabla(tb.get("filas"))
        cand = [e for e in por_firma.get(f, []) if id(e) not in usados] if f else []
        if not cand:
            orphans.append(tb)
        else:
            usados.add(id(cand[0]))
            _aplicar_entrada(tb, cand[0], base_dir, ancho_pt, alto_pt, ancho_max,
                             log, aplicar_deglue, espacio, es_titulo=True)

    # Those that could not be matched by fingerprint: they are matched by order, which is
    # correct unless the .md brings only a subset of the tables.
    if orphans and jt:
        libres = [e for e in jt if id(e) not in usados][:len(orphans)]
        for tb, e in zip(orphans, libres):
            usados.add(id(e))
            log("WARNING  the table on line %d was matched with the json one by "
                "order, not by content." % tb.get("linea", 0))
            _aplicar_entrada(tb, e, base_dir, ancho_pt, alto_pt, ancho_max,
                             log, aplicar_deglue, espacio, es_titulo=True)

    # Figures. The .md and the json may declare the SAME images:
    #   - MinerU/Docling write "![](ruta)" in the .md and also list the
    #     image in the json. Adding the json entry without checking would duplicate the
    #     figure and inflate the figure index (and the verifier would claim a
    #     "Figura N" that is not in the body).
    #   - A .md without image markers leaves the figures to the json alone.
    # Therefore they are first MATCHED by file name; only what has no
    # counterpart in the .md is ADDED.
    jf = [x for x in cl if isinstance(x, dict)
          and x.get("type") in ("image", "picture", "figure")]
    emparejadas = 0
    anadidas = 0
    ignoradas = 0
    if jf:
        por_nombre = {}
        for f in figuras:
            for clave in (f.get("ruta"), f.get("ruta_absoluta")):
                nom = _norm_ref_imagen(clave)
                if nom:
                    por_nombre.setdefault(nom, f)

        pendientes = []
        for k, e in enumerate(jf):
            ruta = e.get("img_path") or e.get("image_path") or e.get("path")
            if not ruta:
                continue
            abs_ruta = ruta if os.path.isabs(ruta) else os.path.join(base_dir, ruta)
            if not os.path.isfile(abs_ruta):
                log("WARNING  figure %d points to %s, which does not exist." % (k + 1, ruta))
                continue
            objetivo = por_nombre.get(_norm_ref_imagen(ruta))
            if objetivo is None:
                pendientes.append((e, ruta, abs_ruta))
                continue
            _medir(objetivo, e.get("bbox"), ancho_pt, alto_pt, ancho_max, espacio, es_figura=True)
            if not objetivo.get("titulo"):
                cap = _primer_texto_util([e.get("caption"), e.get("text")])
                if cap:
                    objetivo["titulo"] = cap
            objetivo["origen_medida"] = "docling-json"
            emparejadas += 1

        # Those that did not match by name: if the .md declared figures, they are matched
        # by order (same source PDF) and the leftovers are ignored. They are only
        # added to the body when the .md declared NO figure at all (pure Docling
        # case, .md without ![]() markers).
        if pendientes:
            if not figuras:
                for e, ruta, abs_ruta in pendientes:
                    fig = {
                        "ruta": ruta,
                        "ruta_absoluta": abs_ruta,
                        "titulo": _primer_texto_util([e.get("caption"), e.get("text")]),
                        "nota": None,
                        "origen": "docling-json",
                        "linea": e.get("page_idx"),
                    }
                    _medir(fig, e.get("bbox"), ancho_pt, alto_pt, ancho_max, espacio, es_figura=True)
                    fig["origen_medida"] = "docling-json"
                    # `analizar` already numbered its figures; the ones from the json arrive
                    # afterwards and must be given indice/existe or the inventory
                    # blows up with KeyError.
                    fig["indice"] = len(figuras) + 1
                    fig["existe"] = os.path.isfile(abs_ruta)
                    figuras.append(fig)
                    anadidas += 1
            else:
                libres = [f for f in figuras if not f.get("origen_medida")]
                for (e, ruta, abs_ruta), f in zip(pendientes, libres):
                    _medir(f, e.get("bbox"), ancho_pt, alto_pt, ancho_max, espacio, es_figura=True)
                    if not f.get("titulo"):
                        cap = _primer_texto_util([e.get("caption"), e.get("text")])
                        if cap:
                            f["titulo"] = cap
                    f["origen_medida"] = "docling-json"
                    emparejadas += 1
                ignoradas = max(0, len(pendientes) - len(libres))

        log("Figures from the json: %d (matched: %d, added: %d, ignored: %d)"
            % (len(jf), emparejadas, anadidas, ignoradas))

    return {
        "usado": True,
        "json": os.path.basename(docling_json),
        "pagina_pt": [ancho_pt, alto_pt],
        "tamano_pagina_origen": origen_pagina,
        "bbox_espacio": espacio,
        "ancho_max_tabla_in": ancho_max,
        "tablas_json": len(jt),
        "tablas_enriquecidas": len(tablas) - len(orphans) if orphans else len(tablas),
        "figuras_json": len(jf),
        "figuras_emparejadas": emparejadas,
        "figuras_anadidas": anadidas,
    }


def _aplicar_entrada(tb, e, base_dir, ancho_pt, alto_pt, ancho_max,
                     log, aplicar_deglue, espacio, es_titulo=True):
    cap = _primer_texto_util([e.get("table_caption")])
    foot = _primer_texto_util([e.get("table_footnote")])
    # Only the caption is the table title. The footnote is the text that goes BELOW
    # it, and in practice MinerU puts there the heading of the section that
    # groups several tables ("EJERCICIO 2: ..."), not the table name: using it
    # as a title would invent a name the document does not have. It is stored as a
    # note, which is what it is.
    if es_titulo and not tb.get("titulo"):
        limpio = limpiar_titulo_json(cap, aplicar_deglue) if cap else None
        if limpio:
            tb["titulo"] = limpio
            tb["titulo_origen"] = "json"
            log("Table title (line %d) taken from the json: %s"
                % (tb.get("linea", 0), limpio[:60]))
        else:
            log("The json brings no caption for the table on line %d; the user is "
                "asked." % tb.get("linea", 0))
    if foot and not tb.get("nota"):
        nota = limpiar_titulo_json(foot, aplicar_deglue)
        if nota:
            tb["nota"] = nota
            tb["nota_origen"] = "json-footnote"

    if e.get("img_path"):
        abs_img = e["img_path"] if os.path.isabs(e["img_path"]) else \
            os.path.join(base_dir, e["img_path"])
        if os.path.isfile(abs_img):
            tb["ruta_imagen"] = e["img_path"]
            tb["ruta_imagen_absoluta"] = abs_img
    _medir(tb, e.get("bbox"), ancho_pt, alto_pt, ancho_max, espacio, es_figura=False)
    if e.get("page_idx") is not None:
        tb["pagina_origen"] = e["page_idx"]


def _detectar_espacio_bbox(entradas, ancho_pt, alto_pt):
    """    MinerU/Docling do not always store the bbox in the same unit.

    In `content_list.json` the bbox comes NORMALIZED to 0..1000 on each axis
    (checked against `middle.json`, which IS in points: the same blocks
    give 612*1.63 = 998 and 792*1.26 = 999). In `middle.json` the bbox is in
    real points of the page.

    The difference is not academic: read as points, a box of 890 wide
    is 12.4 inches, more than an 8.5-inch sheet, and all the tables would come out
    deliberately distorted. It is decided with the data itself instead of
    assuming a format.
    """
    mx = my = 0.0
    for it in entradas:
        b = it.get("bbox") if isinstance(it, dict) else None
        if b and len(b) >= 4:
            try:
                mx = max(mx, float(b[0]), float(b[2]))
                my = max(my, float(b[1]), float(b[3]))
            except (TypeError, ValueError):
                pass
    if not mx and not my:
        return "puntos"
    if mx > ancho_pt * 1.02 or my > alto_pt * 1.02:
        return "normalizado-1000"
    return "puntos"


def _a_pulgadas(valor, ref_pt, espacio):
    return (valor / 1000.0) * ref_pt / 72.0 if espacio == "normalizado-1000" \
        else valor / 72.0


def _norm_ref_imagen(valor):
    """Normalizes an image reference to its file name.

    The .md and the json may cite the same image with different paths
    ("images/fig1.png" versus "fig1.png", or "fig1.png" versus
    "C:\\...\\fig1.png"); what identifies the figure is the base name.
    """
    if not valor:
        return None
    s = str(valor).replace("\\", "/").strip()
    return os.path.basename(s).lower() or None


def _medir(obj, bbox, ancho_pt, alto_pt, ancho_max, espacio, es_figura):
    """    bbox [l, t, r, b] -> width/height in inches, with caps.

    From a table only the WIDTH is taken. Its height is useless: when a table is
    split across two pages, Docling stores one entry per fragment, so
    the box height is that of the piece on that page, not of the table. The
    width IS reliable, and it is the only thing needed to avoid stretching it.
    """
    if not bbox or len(bbox) < 4:
        return
    try:
        l, t, r, b = (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
    except (TypeError, ValueError):
        return
    w_in = _a_pulgadas(r - l, ancho_pt, espacio)
    if w_in <= 0:
        return
    obj["bbox"] = [l, t, r, b]
    obj["bbox_espacio"] = espacio
    obj["ancho_original_in"] = round(w_in, 3)
    obj["ancho_in"] = round(min(w_in, ancho_max), 3)
    if not es_figura:
        return
    h_in = _a_pulgadas(b - t, alto_pt, espacio)
    if h_in > 0:
        obj["alto_original_in"] = round(h_in, 3)
        # A figure should not exceed half the usable height: 9 inches
        # tall in the middle of a body of text is unreadable.
        obj["alto_in"] = round(min(h_in, (alto_pt - 144) / 72.0), 3)


def analizar(md_path, base_dir, log, aplicar_deglue=True,
             detectar_en_linea=False):
    texto = read_text(md_path)
    lineas = texto.splitlines()

    bloques = []           # {tipo, ...}
    tablas = []            # {indice, titulo, filas, nota, origen}
    figuras = []           # {indice, titulo, ruta, nota, origen}
    referencias = []       # {texto, cursivas:[...], heuristica, aviso}
    secciones = []         # {hid, nivel, texto}
    portada_lineas = []
    avisos = []
    cambios_deglue = []
    hid = 0
    en_referencias = False
    dup_titulos = {}
    primera_seccion_vista = False

    i = 0
    n = len(lineas)

    while i < n:
        linea = lineas[i]
        cruda = linea
        t = linea.strip()

        if not t:
            i += 1
            continue

        # --- HTML table (may span one or more lines) ---------------
        if t.lower().startswith("<table"):
            bloque_html = t
            j = i
            while "</table>" not in bloque_html.lower() and j + 1 < n:
                j += 1
                bloque_html += "\n" + lineas[j]
            filas = parse_html_table(bloque_html)
            if filas:
                tablas.append({
                    "filas": filas,
                    "titulo": None,
                    "nota": None,
                    "origen": "html",
                    "linea": i + 1,
                })
                bloques.append({"tipo": "tabla", "indice": len(tablas) - 1})
            else:
                avisos.append("HTML table without cells on line %d; it is skipped." % (i + 1))
            i = j + 1
            continue

        # --- markdown table ---------------------------------------------
        if t.startswith("|"):
            filas, nuevo = parse_md_table(lineas, i)
            if filas:
                tablas.append({
                    "filas": filas,
                    "titulo": None,
                    "nota": None,
                    "origen": "markdown",
                    "linea": i + 1,
                })
                bloques.append({"tipo": "tabla", "indice": len(tablas) - 1})
                i = nuevo
                continue
            i += 1
            continue

        # --- heading -------------------------------------------------
        mh = RE_HEADING.match(t)
        if mh:
            nivel_md = len(mh.group(1))
            texto_h = mh.group(2).strip()

            if nivel_md == 1:
                # The document title: it is the cover title, not a section.
                portada_lineas.append({"tipo": "titulo", "texto": unescape_md(texto_h)})
                bloques.append({"tipo": "portada", "texto": unescape_md(texto_h)})
                i += 1
                continue

            limpio, camb = deglue(strip_leading_markers(unescape_md(texto_h)), aplicar_deglue)
            limpio = fix_colon_spacing(dedupe_numbering(limpio)).strip()
            cambios_deglue.extend(camb)

            # Is it the references section?
            if re.search(r"referencias?\s+bibliogr", norm_key(limpio)) or norm_key(limpio).startswith("referencias"):
                en_referencias = True

            mnum = RE_NUM_HEADING.match(limpio)
            if nivel_md >= 3:
                # APA 7 has 5 levels. '##' covers the first two according to this
                # document (numbered -> level 1, unnumbered -> level 2) and the
                # deeper hashes map to their natural level. Before, EVERYTHING
                # ended up at level 2, so a .md with '###' came out flattened and
                # levels 3, 4 and 5 did not exist in the output.
                nivel = min(nivel_md, 5)
                etiqueta = None
                texto_h2 = limpio
            elif mnum:
                # Numbered main section: "3. Analisis de BSIMM" -> level 1
                nivel = 1
                etiqueta = mnum.group(1)
                texto_h2 = mnum.group(2).strip()
            else:
                nivel = 2
                etiqueta = None
                texto_h2 = limpio

            texto_h2 = fix_colon_spacing(texto_h2).strip().rstrip(":").strip()
            hid += 1
            sec = {"hid": hid, "nivel": nivel, "texto": texto_h2, "etiqueta": etiqueta}
            secciones.append(sec)
            bloques.append({"tipo": "h", "hid": hid, "nivel": nivel, "texto": texto_h2,
                            "segmentos": parse_inline(texto_h2)})
            clave = norm_key(texto_h2)
            dup_titulos.setdefault(clave, []).append(texto_h2)
            primera_seccion_vista = True
            i += 1
            continue

        # --- markdown image -------------------------------------------
        if "![" in t and "](" in t:
            mimg = RE_IMAGEN.search(t)
            if mimg:
                alt, ruta = mimg.group(1).strip(), mimg.group(2).strip()
                # a single line can bring the image and its caption
                cola = t[mimg.end():].strip()
                figs = []
                for tag, num, resto in RE_FIG_CAPTION.findall(cola):
                    figs.append((num, resto))
                titulo = alt
                if figs:
                    titulo = figs[0][1] or titulo
                elif cola and not cola.startswith("|"):
                    titulo = alt or cola
                # The .docx prepends its own "Figura N", so if the caption
                # in the .md already carries it ("Figura 1. Grafico de barras") it must be removed
                # or the PDF shows "Figura 1. Figura 1. Grafico de barras".
                if titulo:
                    titulo = RE_NUM_LEYENDA.sub("", titulo).strip() or None
                figuras.append({
                    "ruta": ruta,
                    "titulo": titulo or None,
                    "nota": None,
                    "origen": "markdown",
                    "linea": i + 1,
                })
                bloques.append({"tipo": "figura", "indice": len(figuras) - 1})
                i += 1
                continue

        # --- references -------------------------------------------------
        if en_referencias:
            limpia = strip_leading_markers(unescape_md(t))
            limpia = fix_colon_spacing(limpia).strip()
            if limpia:
                referencias.append({"texto": limpia, "linea": i + 1})
            i += 1
            continue

        # --- cover: lines before the first section ----------------
        if not primera_seccion_vista:
            limpia = strip_leading_markers(unescape_md(t))
            if limpia:
                portada_lineas.append({"tipo": "dato", "texto": limpia})
                bloques.append({"tipo": "portada_dato", "texto": limpia})
            i += 1
            continue

        # --- paragraph / list --------------------------------------------
        crudo = unescape_md(t)

        # --- level 4 or 5 APA title "inline" (opt-in) ---------------
        # APA 7 defines them like this: "**Titulo.** El texto del parrafo sigue en la
        # misma linea", with first-line indent. Markdown cannot express
        # that structure, so it is recognized by convention. It is opt-in because
        # bold at the start of a paragraph is also plain normal text and
        # enabling it without control would flood the TOC with fake titles.
        #
        # It is matched against the RAW text on purpose: strip_leading_markers()
        # eats the leading asterisks ("**Titulo.**" -> "Titulo.") and if it were
        # detected afterwards there would be no way to tell it from normal text.
        if detectar_en_linea and not en_referencias:
            m_en = RE_TITULO_EN_LINEA.match(crudo.strip())
            if m_en and i + 1 < n:
                titulo_en = m_en.group(2).strip()
                # The body is group 3, NOT what comes after m_en.end(): group 3
                # IS the body and m_en.end() falls just after it.
                cuerpo_en, camb = deglue(m_en.group(3).strip(),
                                         aplicar_deglue)
                cambios_deglue.extend(camb)
                cuerpo_en = fix_colon_spacing(dedupe_numbering(cuerpo_en)).strip()
                hid += 1
                # Three leading asterisks = level 5; two = level 4.
                nivel_en = 5 if m_en.group(1) == "***" else 4
                bloques.append({
                    "tipo": "h", "hid": hid,
                    "nivel": nivel_en,
                    "texto": titulo_en,
                    "en_linea": cuerpo_en,
                    # The segments are resolved by the parser and not by the .docx: the
                    # inline markdown analysis (bold, italic, code)
                    # has a single implementation.
                    "en_linea_segmentos": parse_inline(cuerpo_en),
                    "segmentos": parse_inline(titulo_en),
                })
                secciones.append({
                    "hid": hid,
                    "nivel": nivel_en,
                    "texto": titulo_en,
                    "etiqueta": None,
                })
                i += 1
                continue

        limpia = strip_leading_markers(crudo)
        limpia, camb = deglue(limpia, aplicar_deglue)
        cambios_deglue.extend(camb)
        limpia = fix_colon_spacing(dedupe_numbering(limpia)).strip()

        if limpia:
            # keep the list marker if there is one
            mnum_item = re.match(r"^(\d+[\.\)]|[" + BULLET_CHARS + r"])\s+", limpia)
            bloques.append({
                "tipo": "lista" if mnum_item else "p",
                "texto": limpia,
                "marcador": (mnum_item.group(1) if mnum_item else None),
                "segmentos": parse_inline(limpia),
            })
        i += 1

    # --- table captions: they may come BEFORE or AFTER the table ---------
    # Docling and MinerU do not agree: one emits them before the <table> and the other
    # after. Both orders are accepted in a later pass, which is more
    # reliable than deciding it during the line-by-line walkthrough.
    for k, b in enumerate(bloques):
        if b.get("tipo") != "p" or "segmentos" not in b:
            continue
        texto = b["texto"]
        mt = RE_TAB_CAPTION.match(texto)
        if not mt:
            continue
        # table right after
        if k + 1 < len(bloques) and bloques[k + 1].get("tipo") == "tabla":
            destino = bloques[k + 1]["indice"]
        # table right before
        elif k > 0 and bloques[k - 1].get("tipo") == "tabla":
            destino = bloques[k - 1]["indice"]
        else:
            continue
        if tablas[destino]["titulo"] is None:
            tablas[destino]["titulo"] = fix_colon_spacing(mt.group(2)).strip().rstrip(":")
            b["tipo"] = "nota_tabla"
            b["indice"] = destino
            b.pop("segmentos", None)

    # --- notes written in the .md, right under a table or a figure --------
    # MinerU/Docling do not always deliver the footnote as data, and in practice
    # the .md carries it as an ordinary paragraph: "Nota. Elaboracion propia".
    # Without this pass that paragraph reaches the .docx as body text (indented,
    # in the middle of the document) and the table/figure gets the DEFAULT note
    # instead: the document ends up with two notes, one of them wrong.
    #
    # The .md wins over the footnote of the json: `enriquecer_desde_docling`
    # only writes a note when there is none, exactly like it already does with
    # the title of a table that the .md already names. The explicit override
    # `--notas-tabla-json` (applied later) still wins over both.
    def _destino_de_nota(pos):
        """Index of the table/figure the note at `pos` belongs to, and whether it
        is a figure. Looks forward first (the note goes BELOW) and then backward,
        skipping the caption paragraphs already consumed by the pass above."""
        for salto in (1, -1):
            for pasos in range(1, 4):
                j = pos + salto * pasos
                if j < 0 or j >= len(bloques):
                    break
                v = bloques[j]
                if v.get("tipo") in ("tabla", "figura"):
                    return v["indice"], v.get("tipo") == "figura"
                # A consumed caption belongs to the table/figure: the note can sit
                # on the far side of it. Anything else ends the search.
                if v.get("tipo") != "nota_tabla":
                    break
        return None, None

    for k, b in enumerate(bloques):
        if b.get("tipo") != "p" or "segmentos" not in b:
            continue
        mn = RE_NOTA_LINEA.match(b["texto"])
        if not mn:
            continue
        destino, es_figura = _destino_de_nota(k)
        if destino is None:
            continue
        nota_md = fix_colon_spacing(mn.group(1)).strip()
        if not nota_md:
            continue
        objetos = figuras if es_figura else tablas
        if objetos[destino].get("nota"):
            # Already has a note; the .md line is dropped instead of duplicated.
            b["tipo"] = "nota_tabla" if not es_figura else "nota_figura"
            b["indice"] = destino
            b.pop("segmentos", None)
            continue
        objetos[destino]["nota"] = nota_md
        objetos[destino]["nota_origen"] = "md"
        b["tipo"] = "nota_tabla" if not es_figura else "nota_figura"
        b["indice"] = destino
        b.pop("segmentos", None)
        log("%s note (line %d) taken from the .md: %s"
            % ("Figure" if es_figura else "Table", k + 1, nota_md[:60]))

    # --- table and figure numbering ---------------------------------
    for k, tb in enumerate(tablas):
        tb["indice"] = k + 1
        if not tb["titulo"]:
            avisos.append("Table %d brings no title in the .md: it must be drafted or "
                          "the user must be asked." % (k + 1))
    for k, fg in enumerate(figuras):
        fg["indice"] = k + 1
        ruta_abs = fg["ruta"]
        if not os.path.isabs(ruta_abs):
            ruta_abs = os.path.join(base_dir, ruta_abs)
        fg["ruta_absoluta"] = os.path.normpath(ruta_abs)
        fg["existe"] = os.path.isfile(fg["ruta_absoluta"])
        if not fg["existe"]:
            avisos.append("Figure %d points to '%s', which does not exist on disk. The "
                          "image will NOT be inserted." % (k + 1, fg["ruta"]))
        if not fg["titulo"]:
            avisos.append("Figure %d brings no caption in the .md: it must be drafted or "
                          "the user must be asked." % (k + 1))

    # --- italics proposed for the references -------------------------
    for ref in referencias:
        rango, tipo = propose_reference_italics(ref["texto"])
        if rango:
            ref["cursiva_propuesta"] = rango
            ref["tipo_heuristica"] = tipo
        else:
            ref["cursiva_propuesta"] = None
            ref["tipo_heuristica"] = None
            ref["aviso"] = "the italic part could not be determined (%s)" % tipo
        ref["cursiva_confirmada"] = ref["cursiva_propuesta"]
        ref["revisar_cursiva"] = True
    if referencias:
        avisos.append(
            "The %d references come from the .md and their italic part is a heuristic "
            "PROPOSAL: they must be reviewed one by one before delivery "
            "(references/apa7-format.md, section Referencias)." % len(referencias))

    # --- duplicate titles --------------------------------------------
    dups = {k: v for k, v in dup_titulos.items() if len(v) > 1}
    if dups:
        lista = "; ".join(sorted({v[0] for v in dups.values()}))
        avisos.append("There are %d repeated section titles in the document (%s). The TOC "
                      "will show them as they are: ask the user whether they want them "
                      "distinguished or renamed." % (len(dups), lista))

    # --- glued-word residues the deglue CANNOT resolve ---------------
    # The 'SegunelMarcoBSIMM' case joined two LOWERCASE words, so the
    # lowercase+uppercase rule does not see them. Guessing where each space goes would
    # require a language dictionary: it is more honest to report the residue and let the agent
    # fix it in the source .md instead of producing corrupted text.
    PEGADO_RE = re.compile(r"[A-Za-z\u00c0-\u00ff]{18,}")
    residuos = {}
    for b in bloques:
        txt = b.get("texto")
        if not txt:
            continue
        for m in PEGADO_RE.finditer(txt):
            if " " in m.group(0):
                continue
            residuos[m.group(0)] = residuos.get(m.group(0), 0) + 1
    if residuos:
        # They are sorted by decreasing length, not by frequency: a stretch of 40
        # glued characters is far more suspicious than one of 20, and a long valid
        # word ("internacionalmente") is the typical false positive.
        # Raising the PEGADO_RE threshold to avoid it would really lose 29
        # genuine residues from the reference document, which is a worse trade.
        top = sorted(residuos.items(), key=lambda kv: (-len(kv[0]), -kv[1]))[:12]
        avisos.append(
            "There are %d spaceless stretches that the deglue cannot separate with "
            "confidence (glued lowercase words). They were NOT fixed "
            "automatically: review them in the source .md. Among the longest "
            "(could any of them be a valid word without spaces?): %s"
            % (len(residuos), ", ".join("%s (x%d)" % (t, c) for t, c in top)))
        preguntas_residuo = ["Review in the .md: '%s' (x%d, %d characters). If it is "
                             "a glued word, split it; if it is a valid "
                             "word, leave it."
                             % (t, c, len(t)) for t, c in top]
    else:
        preguntas_residuo = []

    return {
        "bloques": bloques,
        "tablas": tablas,
        "figuras": figuras,
        "referencias": referencias,
        "secciones": secciones,
        "portada_lineas": portada_lineas,
        "avisos": avisos,
        "cambios_deglue": cambios_deglue,
        "residuos_pegado": residuos,
        "preguntas_residuo": preguntas_residuo,
        "n_lineas": n,
    }


# ---------------------------------------------------------------------------
# Cover: automatic detection + gaps that must be asked about
# ---------------------------------------------------------------------------

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def _es_nombre_persona(txt):
    """'Nombre Apellido Segundo' -> True. Requires 2 or more capitalized words."""
    if not txt or re.search(r"\d{1,2}\s+de\s+[a-z]", txt, re.I):
        return False
    if re.search(r"facultad|vicerrector|profesor|docente|nrc|universidad|corporaci", txt, re.I):
        return False
    partes = [p for p in re.split(r"\s+", txt.strip()) if p]
    if len(partes) < 2:
        return False
    cap = 0
    for p in partes:
        if re.match(r"^[A-Z\u00c1\u00c9\u00cd\u00d3\u00da][a-z\u00e0-\u00fc]+", p):
            cap += 1
    return cap >= max(2, len(partes) - 1)


def detectar_portada(lineas, overrides):
    """
    Tries to infer the cover fields from the .md. Whatever cannot be inferred
    is NOT invented: it is listed in 'campos_faltantes' so that the agent asks.

    Typical format in MinerU/Docling documents of academic works:
        # Titulo del trabajo
        Nombre Apellido, Nombre Apellido
        NRC: 00000000
        PROFESOR: Nombre Apellido
        Facultad de <facultad>, <universidad>
            <sede / unidad academica>
        <dia> de <mes> de <ano>

    The institutional fields (facultad, sede, NRC) are detected by pattern and
    are optional: if the document does not bring them, they are asked about. They are never invented.
    """
    portada = {
        "titulo": None,
        "autores": [],
        "facultad": None,
        "vicerrectoria": None,
        "sede": None,
        "materia_nrc": None,
        "docente": None,
        "docente_titulo": None,
        "fecha": None,
        "logo": None,
    }

    for ln in lineas:
        txt = ln["texto"].strip()
        if ln["tipo"] == "titulo":
            portada["titulo"] = portada["titulo"] or txt
            continue

        mn = RE_NRC.search(txt)
        if mn and not portada["materia_nrc"]:
            # 'Nombre Apellido Segundo NRC: 00000000' -> author + NRC separated
            resto = RE_NRC.sub("", txt).strip(" :-\u2013,")
            portada["materia_nrc"] = "NRC " + mn.group(1)
            if _es_nombre_persona(resto):
                portada["autores"] = [p.strip() for p in resto.split(",")]
            continue

        mdoc = re.search(r"(profesor|profesora|docente|teacher)\s*[:\-]?\s*(.*)$", txt, re.I)
        if mdoc and not portada["docente"]:
            nombre = mdoc.group(2).strip(" :-\u2013,")
            partes = [p.strip() for p in nombre.split(",") if p.strip()]
            portada["docente"] = partes[0] if partes else nombre
            if len(partes) > 1 and partes[1]:
                portada["docente_titulo"] = partes[1]
            continue

        # The institution line usually brings facultad + vicerrectoria + sede.
        if re.search(r"facultad|vicerrector|universidad|corporaci\u00f3n", txt, re.I):
            if not portada["vicerrectoria"] and re.search(r"vicerrector", txt, re.I):
                m = re.search(r"vicerrector\u00eda[^,]*", txt, re.I)
                if m:
                    portada["vicerrectoria"] = m.group(0).strip()
                    txt = (txt[: m.start()] + " " + txt[m.end():]).strip(" ,")
            if not portada["sede"] and re.search(r"\bsede\b", txt, re.I):
                m = re.search(r"\bsede\b.*$", txt, re.I)
                if m:
                    portada["sede"] = m.group(0).strip()
                    txt = (txt[: m.start()] + " " + txt[m.end():]).strip(" ,")
            if not portada["facultad"] and re.search(r"facultad|universidad|corporaci\u00f3n", txt, re.I):
                portada["facultad"] = re.sub(r"\s+", " ", txt).strip(" ,")
            continue

        md = RE_DATE.search(txt)
        if md and not portada["fecha"]:
            portada["fecha"] = "%s de %s de %s" % (md.group(1), md.group(2).capitalize(), md.group(3))
            continue

        if not portada["autores"] and "," in txt and _es_nombre_persona(txt):
            portada["autores"] = [p.strip() for p in txt.split(",") if p.strip()]
            continue

        if not portada["autores"] and _es_nombre_persona(txt):
            portada["autores"] = [txt]

    # the user's overrides (what they answered when asked) always win
    for k, v in (overrides or {}).items():
        if v not in (None, "", []):
            portada[k] = v

    # 'autores' is the only cover field that is a list. If the user delivers
    # it as a string, storing it verbatim would make build-docx.js walk it
    # character by character and the cover would come out with one letter per line. What
    # was written is respected literally (it is not split on 'y' nor on commas: that
    # would be inventing authors). For one per line, pass a JSON array.
    if isinstance(portada.get("autores"), str):
        _a = portada["autores"].strip()
        portada["autores"] = [_a] if _a else []
    elif isinstance(portada.get("autores"), list):
        portada["autores"] = [str(x).strip() for x in portada["autores"] if str(x).strip()]

    faltan = []
    for campo, etiqueta in [
        ("titulo", "titulo del trabajo"),
        ("autores", "autores"),
        ("facultad", "facultad / universidad"),
        ("vicerrectoria", "vicerrectoria"),
        ("materia_nrc", "materia y NRC"),
        ("docente", "nombre del docente"),
        ("fecha", "fecha"),
    ]:
        if not portada.get(campo):
            faltan.append(campo)

    # The professional title of the docente is ALWAYS asked about (skill rule),
    # even if the rest of the cover is complete.
    if not portada.get("docente_titulo"):
        faltan.append("docente_titulo")

    # The logo is also ALWAYS asked about. It is optional as to the result,
    # but it cannot be omitted silently: without this line the generator has no
    # way of telling "the user has no logo" from "we asked nobody".
    if not portada.get("logo"):
        faltan.append("logo")

    portada["campos_faltantes"] = faltan
    return portada


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description="Converts a .md into MANIFEST.json")
    ap.add_argument("--md", required=True, help="path of the source .md")
    ap.add_argument("--out", required=True, help="path of the MANIFEST.json to write")
    ap.add_argument("--log", help="log file of the analysis")
    ap.add_argument("--base-dir", help="base directory to resolve image paths")

    ap.add_argument("--portada", help="portada.json with the user's answers")
    ap.add_argument("--sin-deglue", action="store_true", help="do not split glued words")
    ap.add_argument("--terminos-protegidos", metavar="ARCHIVO",
                    help="text file with one protected term per line (exact, "
                         "case-sensitive). Terms are never split by deglue: use it "
                         "for product names and variables such as JavaScript, "
                         "GitHub, SHA256 or NodeJS. Blank lines and lines starting "
                         "with '#' are ignored")
    ap.add_argument("--titulos-tabla-json", help="JSON {index: title} from the agent")
    ap.add_argument("--titulos-figura-json", help="JSON {index: title} from the agent")
    ap.add_argument("--notas-tabla-json",
                    help="JSON {index: note} from the agent; default 'Elaboracion propia'")
    ap.add_argument("--notas-figura-json",
                    help="JSON {index: note} from the agent; default 'Elaboracion propia'")
    ap.add_argument("--detectar-niveles-en-linea", action="store_true",
                    help="treat a paragraph starting with '**Texto.**' as a level 4 title "
                         "and '***Texto.***' as level 5 (APA 7 inline). "
                         "Opt-in: without the option, bold at the start of a paragraph is normal text. "
                         "Disabled by default because it generates false positives.")
    ap.add_argument("--sin-indice-tablas", action="store_true")
    ap.add_argument("--sin-indice-figuras", action="store_true")
    ap.add_argument("--docling-json", dest="docling_json",
                    help="content_list.json from MinerU/Docling to complete titles, "
                         "boxes (bbox) and table images. If omitted, only "
                         "<name>_content_list.json next to the .md is searched")
    ap.add_argument("--pdf", help="original PDF: it gives the real page size to "
                                  "convert bbox into inches")
    ap.add_argument("--ancho-max-tabla", type=float, default=6.5,
                    help="usable width in inches (Letter with 1-inch margins)")
    args = ap.parse_args()

    global TERMINOS_PROTEGIDOS
    if args.terminos_protegidos:
        if not os.path.isfile(args.terminos_protegidos):
            print("ERROR: the protected-terms file does not exist: %s"
                  % args.terminos_protegidos, file=sys.stderr)
            return 2
        with open(args.terminos_protegidos, "r", encoding="utf-8") as fh:
            TERMINOS_PROTEGIDOS = [ln.strip() for ln in fh
                                   if ln.strip() and not ln.lstrip().startswith("#")]

    log_lines = []

    def log(msg):
        line = "[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg)
        log_lines.append(line)
        print(line)

    if not os.path.isfile(args.md):
        print("ERROR: the .md does not exist: %s" % args.md, file=sys.stderr)
        return 2

    base_dir = args.base_dir or os.path.dirname(os.path.abspath(args.md))

    log("=== PHASE 1: .md analysis ===")
    log("Source       : %s" % args.md)
    log("sha256       : %s" % sha256_of(args.md))
    log("Size         : %d bytes" % os.path.getsize(args.md))
    log("Base dir     : %s" % base_dir)
    log("")

    # MinerU/Docling deliver the prose .md next to a *_content_list.json with the
    # structure. If it is at hand it is used alone: that is the normal case for those flows.
    docling_json = args.docling_json
    if docling_json is None:
        raiz = os.path.splitext(os.path.abspath(args.md))[0]
        for cand in (raiz + "_content_list.json", raiz + "_content_list_v2.json"):
            if os.path.isfile(cand):
                docling_json = cand
                log("MinerU/Docling json detected on its own: %s" % os.path.basename(cand))
                break

    res = analizar(args.md, base_dir, log,
                   aplicar_deglue=not args.sin_deglue,
                   detectar_en_linea=args.detectar_niveles_en_linea)

    info_docling = {}
    if docling_json:
        info_docling = enriquecer_desde_docling(
            res["tablas"], res["figuras"], docling_json, base_dir, args.md,
            args.pdf, args.ancho_max_tabla, log,
            aplicar_deglue=not args.sin_deglue)
        log("")

    log("--- Inventory ---")
    log("Sections (headings)  : %d" % len(res["secciones"]))
    log("Tables               : %d" % len(res["tablas"]))
    log("Figures              : %d" % len(res["figuras"]))
    log("References           : %d" % len(res["referencias"]))
    log("Cover lines          : %d" % len(res["portada_lineas"]))
    log("")

    if res["cambios_deglue"]:
        ambiguos = [c for c in res["cambios_deglue"] if c[3]]
        claros = [c for c in res["cambios_deglue"] if not c[3]]
        log("--- Glued words correction (%d changes) ---" % len(res["cambios_deglue"]))
        if claros:
            log("  Clear corrections (%d):" % len(claros))
            vistos = set()
            for tipo, antes, despues, _ in claros:
                clave = (antes, despues)
                if clave in vistos:
                    continue
                vistos.add(clave)
                log("    %-9s %-34s -> %s" % (tipo, antes, despues))
        if ambiguos:
            log("  REVIEW these (%d) - mixed alphanumeric / camelCase, they may be"
                % len(ambiguos))
            log("  real product names or variables. Protect them with"
                " --terminos-protegidos and re-run if so:")
            vistos = set()
            for tipo, antes, despues, _ in ambiguos:
                clave = (antes, despues)
                if clave in vistos:
                    continue
                vistos.add(clave)
                log("    %-9s %-34s -> %s" % (tipo, antes, despues))
        log("")

    log("--- Detected sections ---")
    for s in res["secciones"]:
        log("  H%-5d [n%d] %s" % (s["hid"], s["nivel"], s["texto"][:88]))
    log("")

    log("--- Detected tables ---")
    for tb in res["tablas"]:
        log("  Table %-2d %d rows x %d cols  title=%s  (line %d, %s)"
            % (tb["indice"], len(tb["filas"]), len(tb["filas"][0]),
               (tb["titulo"] or "NO TITLE"), tb["linea"], tb["origen"]))
    log("")

    if res["figuras"]:
        log("--- Detected figures ---")
        for fg in res["figuras"]:
            log("  Figure %-2d %s  exists=%s  title=%s"
                % (fg["indice"], fg["ruta"], fg["existe"], (fg["titulo"] or "NO TITLE")))
        log("")

    if res["referencias"]:
        log("--- Detected references (alphabetical order is applied in the .docx) ---")
        for ref in res["referencias"]:
            log("  - %s" % ref["texto"][:110])
            log("      proposed italic: %s  [%s]  REVIEW" %
                (ref["cursiva_propuesta"], ref.get("tipo_heuristica") or "?"))
        log("")

    # --- titles supplied by the agent -----------------------------------
    if args.titulos_tabla_json:
        tt = read_json(args.titulos_tabla_json)
        for k, v in tt.items():
            for tb in res["tablas"]:
                if tb["indice"] == int(k):
                    tb["titulo"] = v
        log("Table titles overwritten from %s" % args.titulos_tabla_json)
    if args.titulos_figura_json:
        tf = read_json(args.titulos_figura_json)
        for k, v in tf.items():
            for fg in res["figuras"]:
                if fg["indice"] == int(k):
                    fg["titulo"] = v
        log("Figure titles overwritten from %s" % args.titulos_figura_json)

    # --- table and figure notes ----------------------------------------
    if args.notas_tabla_json:
        nt = read_json(args.notas_tabla_json)
        for k, v in nt.items():
            for tb in res["tablas"]:
                if tb["indice"] == int(k):
                    tb["nota"] = v
                    tb["nota_origen"] = "agente"
        log("Table notes overwritten from %s" % args.notas_tabla_json)
    if args.notas_figura_json:
        nf = read_json(args.notas_figura_json)
        for k, v in nf.items():
            for fg in res["figuras"]:
                if fg["indice"] == int(k):
                    fg["nota"] = v
                    fg["nota_origen"] = "agente"
        log("Figure notes overwritten from %s" % args.notas_figura_json)

    # references/apa7-format.md: "Por defecto 'Nota. Elaboracion propia.',
    # salvo que el .md indique otra fuente". That default is mandatory because
    # a table or figure without a note does not comply with APA. Before the note was only printed
    # if the .md or the json brought it, so in an ordinary document none appeared and the verifier did not
    # detect it: it was delivered without one.
    NOTA_PROPIA = "Elaboración propia"
    for tb in res["tablas"]:
        if not tb.get("nota"):
            tb["nota"] = NOTA_PROPIA
            tb["nota_origen"] = "default"
    for fg in res["figuras"]:
        if not fg.get("nota"):
            fg["nota"] = NOTA_PROPIA
            fg["nota_origen"] = "default"

    # --- portada ---------------------------------------------------------
    overrides = {}
    if args.portada and os.path.isfile(args.portada):
        overrides = read_json(args.portada)
    portada = detectar_portada(res["portada_lineas"], overrides)

    log("--- Detected cover ---")
    for k in ("titulo", "autores", "facultad", "vicerrectoria", "materia_nrc",
              "docente", "docente_titulo", "fecha", "logo"):
        log("  %-15s : %s" % (k, portada.get(k)))
    log("  MISSING FIELDS (ask the user, do NOT infer): %s"
        % (", ".join(portada["campos_faltantes"]) or "(none)"))
    log("")

    # --- index policy --------------------------------------------
    hay_tablas = len(res["tablas"]) > 0
    hay_figuras = len(res["figuras"]) > 0
    idx_tablas = hay_tablas and not args.sin_indice_tablas
    idx_figuras = hay_figuras and not args.sin_indice_figuras

    log("--- Index policy ---")
    log("  Table of contents : ALWAYS")
    log("  Table index       : %s (tables found: %d)"
        % ("YES" if idx_tablas else "NO", len(res["tablas"])))
    log("  Figure index      : %s (figures found: %d)"
        % ("YES" if idx_figuras else "NO", len(res["figuras"])))
    if not idx_tablas:
        log("  -> No tables in the .md. ASK the user whether the document has no tables,")
        log("     and do NOT generate an empty table index.")
    if not idx_figuras:
        log("  -> No figures in the .md. ASK the user whether the document has no figures,")
        log("     and do NOT generate an empty figure index.")
    log("")

    # --- warnings ----------------------------------------------------------
    if res["avisos"]:
        log("--- Warnings (%d) ---" % len(res["avisos"]))
        for a in res["avisos"]:
            log("  * %s" % a)
        log("")

    # --- questions for the user ---------------------------------------------
    # Each question declares whether it BLOCKS delivery. Blocking ones are not
    # resolved by inventing: they are asked, and the document is not delivered until
    # the user answers. Before they were loose strings in the log, with no way of
    # telling them from a warning, and the document came out with "(sin titulo)".
    preguntas = []

    def preguntar(pid, bloqueante, campo, tipo, objetivo, indice, texto,
                  opciones=None, como_resolver=None):
        preguntas.append({
            "id": pid,
            "bloqueante": bool(bloqueante),
            "campo": campo,
            "tipo": tipo,
            "objetivo": objetivo,
            "indice": indice,
            "pregunta": texto,
            "opciones": opciones or [],
            "como_resolver": como_resolver,
        })

    # Cover fields that are ASKED about but do NOT block: the document is
    # built just fine without them. This is consistent with the warnings (critico=False) that
    # verificar-pdf.py emits for the same fields.
    #
    # vicerrectoria and materia_nrc used to block delivery. They are optional
    # (references/institutional-cover.md): the cover prints the line only if the
    # value exists, and a missing one never leaves a blank gap.
    PORTADA_SIN_BLOQUEAR = {"logo", "docente_titulo", "vicerrectoria"}
    PREGUNTA_PORTADA = {
        "logo": ("There is no logo on the cover. If one is provided it is placed above the "
                 "title. Is it included or omitted?"),
        "docente_titulo": "What is the title or profession of the docente?",
        "materia_nrc": ("There is no course/NRC line on the cover. Provide the course and its "
                        "code, or omit the line?"),
        "vicerrectoria": ("There is no vice-rector's office line on the cover. Provide it, or "
                          "omit the line?"),
    }

    for c in portada["campos_faltantes"]:
        preguntar("portada_%s" % c, c not in PORTADA_SIN_BLOQUEAR, c,
                  "agregar_u_omitir", "Cover", None,
                  PREGUNTA_PORTADA.get(c, "Missing '%s' on the cover: add it or omit it?" % c),
                  ["agregar", "omitir"],
                  "Record the decision in portada.json and run the "
                  "parser again with --portada.")
    if not idx_tablas:
        preguntar("sin_tablas", True, "estructura", "confirmacion", "Document",
                  None, "The .md contains no tables: confirm that the document "
                        "has no tables.",
                  ["confirmar", "falta_alguno"],
                  "If a table is missing, review the .md before continuing.")
    if not idx_figuras:
        preguntar("sin_figuras", True, "estructura", "confirmacion", "Document",
                  None, "The .md contains no figures: confirm that the document "
                        "has no figures.",
                  ["confirmar", "falta_alguno"],
                  "If a figure is missing, review the .md before continuing.")

    OPCION_TITULO = ["redactar_desde_contexto", "usuario_proporciona"]
    for tb in res["tablas"]:
        if not tb["titulo"]:
            preguntar(
                "titulo_tabla_%d" % tb["indice"], True, "titulo",
                "eleccion_binaria", "Table", tb["indice"],
                "Table %d has no title. Should I draft the title from the "
                "report context, or will you provide it?" % tb["indice"],
                OPCION_TITULO,
                "If you choose 'redactar_desde_contexto': write the title in "
                "titulos-tabla.json as {\"%d\": \"...\"} and run the "
                "parser again with --titulos-tabla-json." % tb["indice"])
    for fg in res["figuras"]:
        if not fg["titulo"]:
            preguntar(
                "titulo_figura_%d" % fg["indice"], True, "titulo",
                "eleccion_binaria", "Figure", fg["indice"],
                "Figure %d has no caption. Should I draft the caption from the "
                "report context, or will you provide it?" % fg["indice"],
                OPCION_TITULO,
                "If you choose 'redactar_desde_contexto': write the caption in "
                "titulos-figura.json as {\"%d\": \"...\"} and run the "
                "parser again with --titulos-figura-json." % fg["indice"])

    if res["referencias"]:
        preguntar("cursiva_referencias", False, "referencias", "revision",
                  "References", None,
                  "Review the italic part of the %d references (heuristic)."
                  % len(res["referencias"]))
    # The residues are detected by the analyzer but were copied neither to the manifest
    # nor to the question list: the agent learned about the problem from the
    # console, not from the output that is actually consumed.
    for n, pr in enumerate(res["preguntas_residuo"]):
        preguntar("residuo_%d" % n, False, "deglue", "revision_manual",
                  "Glued words", None, pr)

    preguntas_texto = [q["pregunta"] for q in preguntas]
    pendientes_bloqueantes = [q["id"] for q in preguntas if q["bloqueante"]]


    log("--- Questions pending for the user (%d, %d blocking) ---"
        % (len(preguntas), len(pendientes_bloqueantes)))
    for q in preguntas:
        log("  ? [%s] %s" % ("BLOCKS" if q["bloqueante"] else "warning", q["pregunta"]))
    if pendientes_bloqueantes:
        log("")
        log("  PENDING BLOCKING: %s" % ", ".join(pendientes_bloqueantes))
        log("  The document cannot be delivered until they are resolved.")
    log("")

    # --- manifest ------------------------------------------------------
    manifiesto = {
        "version": VERSION_MANIFIESTO,
        "generado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "fuente": {
            "md": os.path.abspath(args.md),
            "sha256": sha256_of(args.md),
            "bytes": os.path.getsize(args.md),
            "base_dir": base_dir,
        },
        "opciones": {
            "numerar_pagina_portada": False,
            "dedupe_toc": False,
            "deglue_aplicado": not args.sin_deglue,
            "indice_tablas": idx_tablas,
            "indice_figuras": idx_figuras,
        },
        "portada": portada,
        "docling": info_docling,
        "secciones": res["secciones"],
        "bloques": res["bloques"],
        "tablas": res["tablas"],
        "figuras": res["figuras"],
        "referencias": res["referencias"],
        "diagnostico": {
            "n_lineas": res["n_lineas"],
            "tablas_encontradas": len(res["tablas"]),
            "figuras_encontradas": len(res["figuras"]),
            "referencias_encontradas": len(res["referencias"]),
            "cambios_deglue": len(res["cambios_deglue"]),
            "residuos_pegado": res["residuos_pegado"],
            "avisos": res["avisos"],
            "preguntas": preguntas,
            "preguntas_texto": preguntas_texto,
            "pendientes_bloqueantes": pendientes_bloqueantes,
        },
    }

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(manifiesto, fh, ensure_ascii=False, indent=2)

    log("MANIFEST written: %s (%d bytes)" % (args.out, os.path.getsize(args.out)))

    if args.log:
        log_dir = os.path.dirname(os.path.abspath(args.log))
        if log_dir and not os.path.isdir(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        with open(args.log, "w", encoding="utf-8") as fh:
            fh.write("\n".join(log_lines) + "\n")
        log("Log written: %s" % args.log)

    return 0


if __name__ == "__main__":
    sys.exit(main())


