#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md-a-manifiesto.py - Convierte un .md (y sus JSON de layout) en MANIFIEST.json

Este script es el CEREBRO del pipeline. Antes, cada documento obligaba a
reescribir un generador .js de cientos de lineas con todo hardcodeado (portada,
titulos, tablas, referencias). Ahora el .md se analiza UNA vez, se produce un
manifiesto declarativo, y build-docx.js solo lo aplica.

Que arregla respecto al pipeline anterior:
  * Las referencias SIEMPRE salen del .md. Antes se ignoraban por completo y
    se emitian 5 referencias literales fijas del documento de ejemplo.
  * La correccion de "palabras pegadas" (deglue) se implementa de verdad, con
    lista blanca. Antes solo estaba documentada y nunca se ejecutaba.
  * Se inventarioan tablas, figuras y referencias, y se DECLARA cuando falta
    algo, para que build-docx.js aplique la politica de indices condicionales.
  * Se detectan vinetas U+2022 y el artefacto "o " de Docling, y la
    numeracion duplicada "1. 1.".

Solo libreria estandar: no requiere pip install de nada.

Uso tipico:
    python md-a-manifiesto.py --md doc.md --out salida/MANIFEST.json \
        --log salida/_logs/01-analisis.log [--imagenes salida/imagenes.json] \
        [--portada portada.json] [--ancho-pagina-landscape]

Respuestas del agente (JSON {indice: texto}, indice base 1):
    --titulos-tabla-json   titulo de cada tabla, si el .md no lo trae
    --titulos-figura-json  leyenda de cada figura, si el .md no la trae
    --notas-tabla-json      nota de cada tabla; por defecto "Elaboracion propia"
    --notas-figura-json     nota de cada figura; por defecto "Elaboracion propia"

    Se pasan solo para los indices que falten. Sirven para responder a las
    preguntas bloqueantes que el propio parser deja en
    MANIFEST.json -> diagnostico.preguntas: si el usuario elige que se redacte
    el titulo desde el contexto, se escribe aqui y se vuelve a ejecutar el
    parser. build-docx.js NO acepta estos ficheros.

Opciones de analisis:
    --sin-deglue            no aplicar la correccion de "palabras pegadas"
    --whitelist FICH        terminos intocables (sin lista base en el codigo)
    --detectar-niveles-en-linea
                            tratar "**Titulo.** texto" como nivel 4 y
                            "***Titulo.*** texto" como nivel 5 (APA 7 en linea).
                            Opt-in: la negrita al abrir parrafo tambien es texto
                            normal y activarlo sin control crearia titulos falsos.

Salida: MANIFEST.json + un informe de analisis legible. Las preguntas
bloqueantes pendientes quedan en diagnostico.pendientes_bloqueantes; los
titulos o leyendas ausentes hacen que build-docx.js aborte en vez de inventarlos.
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
# Lista blanca de terminos que NUNCA se separan al pegar palabras.
#
# No hay lista base en el codigo, a proposito: cualquier termino incrustado
# aqui contaminationa el pipeline con el dominio de un trabajo concreto. La
# lista blanca de un documento se define por documento, en
#   references/terminos-whitelist.txt
# y se pasa con  --whitelist <fichero>.
#
# La comparacion es SENSIBLE A MAYUSCULAS y por subcadena, asi que conviene
# anadir el termino con su forma exacta tal como aparece en el .md.
#
# Pauta al elegir terminos:
#   - Anade los que tienen camelCase interno (NodeJS -> 'Node JS', MySQL ->
#     'My SQL', OpenID -> 'Open ID') y los que mezclan letra y numero
#     (IPv4, SHA256, AES128 -> 'IPv 4', 'SHA 256', 'AES 128').
#   - Las siglas en MAYUSCULAS puras (OWASP, CVSS, SDL) NO necesitan entrar: la
#     regla solo parte minuscula seguida de mayuscula, asi que ya son seguras.
#   - EVITA terminos de menos de 4 caracteres (Git, CI, QA, In). Comparados por
#     subcadena protegen palabras sin relacion: "Di-git-al", "de-ci-sion", "c-d".
#   - Los apellidos con prefijo (McDonald, MacArthur) si se rompen, se anaden al
#     fichero del documento; no se meten aqui.
#
# Sin lista blanca el deglue puede partir un termino legitimo. No es un fallo
# silencioso: todo lo que cambia queda en el informe de residuos para revision.
# ---------------------------------------------------------------------------

# Vinetas y simbolos que las conversiones dejan pegados al inicio de linea.
BULLET_CHARS = "\u2022\u25cf\u25aa\u25e6\u2043\u2219\u00b7\u2023\u2027"

# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------


def norm_key(text):
    """Minusculas sin acentos, para comparar titulos de forma tolerante."""
    t = unicodedata.normalize("NFD", text.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def read_text(path):
    """Lee texto detectando UTF-8/BOM. Los .md de MinerU y Docling son UTF-8."""
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
    """Carga un JSON de entrada tolerando BOM.

    Importa usar read_text y no open(..., encoding='utf-8'): en Windows,
    Set-Content -Encoding UTF8 y la mayoria de editores guardan con BOM, y
    json.load aborta con 'Unexpected UTF-8 BOM' si el fichero la lleva.
    """
    return json.loads(read_text(path))


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def unescape_md(text):
    """Quita el escape de markdown: 'A00\\_2021' -> 'A00_2021'."""
    return re.sub(r"\\([_\\*`~\[\]()#+\-.!])", r"\1", text)


def strip_leading_markers(text):
    """
    Quita vinetas de conversion al inicio de la linea.

    Casos reales medidos en documentos de Docling:
      '\u2022 Objetivos Especificos:'  -> 'Objetivos Especificos:'
      '\u2022 1. Entrenamiento (Training): ...' -> '1. Entrenamiento ...'
      'o Descripcion Tecnica: ...'       -> 'Descripcion Tecnica: ...'
    El patron 'o ' solo se quita si NO es una palabra legitima: exigimos que
    tras el marcador venga una mayuscula o un numero, y que el resto sea largo.
    """
    original = text
    # 1) Vinetas unicode (posiblemente repetidas: '\u2022 \u2022 texto')
    text = re.sub(r"^[\s" + BULLET_CHARS + r"]+[\s]*", "", text)
    # 2) Marcadores ascii repetidos: '- - texto', '* * texto', '\u2013 \u2022 texto'
    text = re.sub(r"^[\s\-\*\u2013\u2014]{2,}", "", text)
    # 3) Artefacto 'o ' de Docling (vinieta de segundo nivel convertida a letra)
    m = re.match(r"^[oO][\s\t]+(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]|\d)", text)
    if m and len(text) > 3:
        text = text[m.end():]
    return text if text.strip() else original


def dedupe_numbering(text):
    """
    Normaliza la numeracion de lista duplicada por artefacto de extraccion.

    '1. 1. Analizar ...' -> '1. Analizar ...'
    '\u2022 1. 1. texto'  -> '\u2022 1. texto'   (el prefijo ya se quito antes)
    Se conserva el primer numero y se descarta el repetido.
    """
    text = re.sub(r"^(\d+)\.\s+\d+\.\s+", r"\1. ", text)
    text = re.sub(r"^(\d+)\)\s+\d+\)\s+", r"\1) ", text)
    return text


def fix_colon_spacing(text):
    """'ACTIVIDAD 4 : Title' -> 'ACTIVIDAD 4: Title'  (espacio antes de ':')."""
    return re.sub(r"[ \t]+([:;])", r"\1", text)


# ---------------------------------------------------------------------------
# deglue: separar palabras pegadas por la conversion
# ---------------------------------------------------------------------------

GLUE_RE = re.compile(r"([a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0])([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc\u00c0\u00c8][a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0]{0,})")
# 'ACTIVIDAD1:' / 'CR1.4' -> numero pegado a palabra EN MAYUSCULAS.
# El guardia (?![-\u2013]\d) evita partir identificadores con guion: 'CVE2021-44228'
# debe quedarse entero, no convertirse en 'CVE 2021-44228'.
GLUE_NUM_RE = re.compile(r"\b([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{3,})(\d+)\b(?![-\u2013]\d)")
# 'SegunelMarco' ya lo cubre GLUE_RE; ':Diferenciacion' cubre el degglu de ':'
COLON_RE = re.compile(r"(:)(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc])")
# 'Digital(Portal)' -> 'Digital (Portal)': falta el espacio antes del parentesis.
# Se exige >=4 caracteres dentro para no tocar abreviaturas como '(Ed.', '(n. 2)'.
PAREN_RE = re.compile(r"([a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0])\(([^()]{4,})\)")
# 'SCAySBOM' -> 'SCA y SBOM': sigla, conjuncion de una letra, otra sigla.
# Es la pegadura de siglas mas comun en informes y GLUE_RE no la ve, porque
# exige minuscula seguida de mayuscula y aqui solo hay una minuscula en medio.
GLUE_SIGLA_Y_RE = re.compile(
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{3,})"
    r"([a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0])"
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{2,})")
# 'OWASPTop10' -> 'OWASP Top10': sigla en mayusculas pegada a capitalizada.
GLUE_ACR_RE = re.compile(
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc]{3,})"
    r"([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc][a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0]+)")
# 'Top10' -> 'Top 10', 'Junio2021' -> 'Junio 2021': palabra pegada a numero.
# La lista blanca cubre los casos en que NO se separan: IPv6, SHA256, Base64...
# De ahi la importancia de anadir a terminos-whitelist.txt los terminos con
# letra y numero del documento en curso.
GLUE_PALABRA_NUM_RE = re.compile(
    r"([A-Za-z\u00c0-\u00ff]{3,})(\d{1,4})(?![0-9A-Za-z-])")
# '10de' -> '10 de', '2Capa' -> '2 Capa', '3Mes5' -> '3 Mes 5': numero pegado a
# la palabra siguiente. Solo se acepta capitalizada o una de las palabras
# funcionales del espanol, para no partir '3er' ni 'v2beta'.
GLUE_NUM_PALABRA_RE = re.compile(
    r"(\d{1,4})(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00dc][a-z\u00f1\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00e0]"
    r"|(?:de|del|la|el|los|las|un|una|y|en|con|para|por)\b)")


def deglue(text, whitelist, apply=True):
    """
    Inserta espacios en palabras pegadas por la conversion del PDF a texto.

    Casos reales del documento de referencia:
      'SegunelMarcoBSIMM'                  -> 'Segun el Marco BSIMM'
      'modeladodeamenazas'                 -> 'modelado de amenazas'
      'ACTIVIDAD4'                         -> 'ACTIVIDAD 4'
      '2:Diferenciacion'                   -> '2: Diferenciacion'

    Por que NO separar:
      los terminos de la lista blanca del documento (--whitelist), que son
      camelCase legitimos: separarlos produciria 'Node JS', que es un error
      igual de grave que no separar nada.
      Las siglas en MAYUSCULAS puras ('OWASP', 'CVSS') no se tocan: la regla
      exige minuscula seguida de mayuscula.

    La lista blanca se evalua sobre la PALABRA completa que contiene el punto de
    union, no sobre una ventana de contexto. Esto es deliberado: con una ventana
    de contexto, terminos cortos de la lista ('In', 'CI', 'CD') aparecen dentro
    de casi cualquier frase y desactivarian la correccion en todo el documento.
    A nivel de palabra, un termino protegido queda intacto y 'yDatos' se separa
    igual.

    Devuelve (texto_corregido, lista_de_cambios) para poder reportarlos.
    """
    cambios = []
    if not apply or not text:
        return text, cambios

    # Regex de lista blanca, escaped y SENSIBLE A MAYUSCULAS.
    #
    # La sensibilidad no es un detalle: con re.I el termino "Git" (el control de
    # versiones) aparecia dentro de "DiGITal" y protegia la palabra completa, y
    # "CI" aparecia dentro de "deCIsion". Con comparacion insensible, cualquier
    # sigla corta anula la correccion en medio el documento. Comparando con las
    # mayusculas exactas, "Git" solo protege "Git", "GitLab" y "GitHub".
    if whitelist:
        pat = "|".join(re.escape(t) for t in sorted(whitelist, key=len, reverse=True))
        wl_re = re.compile("(?:" + pat + ")")
        # Subconjunto que puede proteger por subcadena. El criterio NO es
        # "tiene mayusculas": "OWASP" y "SCA" las tienen y sin embargo son
        # justamente las siglas que hay que poder partir ("OWASPTop10",
        # "SCAySBOM"). Entran los nombres MEZCLADOS ("OpenID", "JavaScript",
        # "IPv6") y los que llevan digitos ("ISO27002", "MD5"), que si conviene
        # proteger aunque sean mayusculas.
        wl_camel = frozenset(
            t for t in whitelist
            if (any(c.islower() for c in t) and any(c.isupper() for c in t[1:]))
            or (any(c.isdigit() for c in t) and any(c.isupper() for c in t))
        )
    else:
        wl_re = None
        wl_camel = frozenset()

    # Maximo run de caracteres que pueden formar una palabra pegada.
    token_chars = r"[0-9A-Za-z\u00c0-\u00ff_./\\-]"

    def palabra_en(pos, longitud, cadena):
        ini = pos
        while ini > 0 and re.match(token_chars, cadena[ini - 1]):
            ini -= 1
        fin = pos + longitud
        while fin < len(cadena) and re.match(token_chars, cadena[fin]):
            fin += 1
        return cadena[ini:fin]

    def protegida(pos, longitud, cadena):
        if wl_re is None:
            return False
        palabra = palabra_en(pos, longitud, cadena)
        # Dos modos de proteccion, y la distincion es lo que hace que la lista
        # blanca sirva para algo:
        #
        #  1. Coincidencia EXACTA: la palabra pegada es el termino ("Portal",
        #     "SCA"). Protege siempre.
        #  2. Subcadena: el termino esta DENTRO de una palabra pegada mas larga
        #     ("OpenID" dentro de "OpenIDConnect"). Solo se aplica a terminos
        #     en camelCase o con digitos.
        #
        # Las siglas EN MAYUSCULAS no protegen por subcadena a proposito:
        # si lo hicieran, "SCA" protegeria "SCAySBOM" y "OWASP" protegeria
        # "OWASPTop10", que son exactamente las pegaduras que hay que deshacer.
        # El precio es que un producto como "OWASPJuiceShop" habra que anadirlo
        # completo a la lista blanca; la lista es editable justamente para eso.
        if palabra in whitelist:
            return True
        return any(t in palabra for t in wl_camel)

    def sub_glue(m):
        a, b = m.group(1), m.group(2)
        if protegida(m.start(), len(m.group(0)), text):
            return m.group(0)
        cambios.append(("glue", m.group(0), a + " " + b))
        return a + " " + b

    out = GLUE_RE.sub(sub_glue, text)

    # El orden importa. GLUE_RE va primero porque ve 'SegunelMarco'; pero parte
    # 'SCAySBOM' en 'SCAy SBOM' y despues ya no queda nada que una regla de
    # siglas pueda arreglar, asi que las reglas de sigla se aplican ANTES que
    # GLUE_RE. Por eso se evaluan todas contra 'text' y se encadenan en orden
    # inverso al de aplicacion.
    def sub_sigla_y(m):
        if protegida(m.start(), len(m.group(0)), text):
            return m.group(0)
        cambios.append(("sigla_y", m.group(0), m.group(1) + " " + m.group(2) + " " + m.group(3)))
        return m.group(1) + " " + m.group(2) + " " + m.group(3)

    def sub_acr(m):
        if protegida(m.start(), len(m.group(0)), text):
            return m.group(0)
        cambios.append(("sigla", m.group(0), m.group(1) + " " + m.group(2)))
        return m.group(1) + " " + m.group(2)

    def sub_num_palabra(m):
        if protegida(m.start(), len(m.group(0)), text):
            return m.group(0)
        cambios.append(("num_palabra", m.group(0), m.group(0) + " "))
        return m.group(0) + " "

    def sub_num(m):
        if protegida(m.start(), len(m.group(0)), text):
            return m.group(0)
        cambios.append(("glue_num", m.group(0), m.group(1) + " " + m.group(2)))
        return m.group(1) + " " + m.group(2)

    def sub_palabra_num(m):
        if protegida(m.start(), len(m.group(0)), text):
            return m.group(0)
        cambios.append(("palabra_num", m.group(0), m.group(1) + " " + m.group(2)))
        return m.group(1) + " " + m.group(2)

    out = GLUE_NUM_PALABRA_RE.sub(sub_num_palabra, text)
    out = GLUE_PALABRA_NUM_RE.sub(sub_palabra_num, out)
    out = GLUE_NUM_RE.sub(sub_num, out)
    out = GLUE_SIGLA_Y_RE.sub(sub_sigla_y, out)
    out = GLUE_ACR_RE.sub(sub_acr, out)
    out = GLUE_RE.sub(sub_glue, out)

    def sub_colon(m):
        if protegida(m.start(), len(m.group(0)), out):
            return m.group(0)
        cambios.append(("colon", m.group(0), m.group(0) + " "))
        return m.group(0) + " "

    out = COLON_RE.sub(sub_colon, out)

    def sub_paren(m):
        # Proteger solo la palabra ANTERIOR al parentesis. Si se midiera el
        # match entero, palabra_en devolveria "Digital(Portal)" completo y
        # el termino en lista blanca de dentro del parentesis protegeria toda la
        # frase, con lo que esta regla no corregiria nada util.
        if protegida(m.start(), 1, out):
            return m.group(0)
        cambios.append(("paren", m.group(0), m.group(1) + " (" + m.group(2) + ")"))
        return m.group(1) + " (" + m.group(2) + ")"

    out = PAREN_RE.sub(sub_paren, out)
    return out, cambios


# ---------------------------------------------------------------------------
# Formato en linea -> segmentos (negrita / cursiva)
# ---------------------------------------------------------------------------

INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*[^*\n]+?\*|__.+?__|_[^_\n]+?_)")


def parse_inline(text):
    """
    Convierte 'texto **negrita** y *cursiva*' en una lista de segmentos.
    Devuelve [{'text':..., 'bold':bool, 'italics':bool}, ...]
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
# Tablas
# ---------------------------------------------------------------------------

HTML_TABLE_RE = re.compile(r"<table[^>]*>.*?</table>", re.I | re.S)
TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.I | re.S)
TD_RE = re.compile(r"<(td|th)[^>]*>(.*?)</\1>", re.I | re.S)
BR_RE = re.compile(r"<br\s*/?>", re.I)
TAG_RE = re.compile(r"<[^>]+>")


def cell_to_text(raw):
    txt = BR_RE.sub(" \u21b5 ", raw)   # salto de linea dentro de la celda
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
    """Detecta una tabla markdown |a|b| ... Devuelve (filas, fin_indice)."""
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
# Referencias: deteccion del rango en cursiva (heuristica + avisos)
# ---------------------------------------------------------------------------

YEAR_RE = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")
REVISTA_RE = re.compile(r"\b(revista|journal|revista|vol\.|volumen|n[o\u00ba]\.|pp\.|edition|ed\.)", re.I)
TESIS_RE = re.compile(r"\b(repositorio|tesis|trabajo de grado|universidad|universidad tecnica|monograf)", re.I)
URL_TAIL_RE = re.compile(r"((?:Recuperado de|Disponible en|consultado|Retrieved from|Available from)\s*:?\s*https?://\S+)\s*$", re.I)
BARE_URL_RE = re.compile(r"(https?://\S+)\s*$")


def propose_reference_italics(text):
    """
    PROPUESTA (no definitiva) de que parte de la referencia va en cursiva.

    APA 7 depende de si la fuente es un libro, informe, articulo de revista o
    tesis, y esa distincion no se puede determinar de forma fiable solo con el
    texto plano que produce la conversion. Por eso:

      1. se aplica una heuristica conservadora, y
      2. se devuelve un aviso para que el agente/agente-usuario la revise.

    Reglas:
      * Revista/periodico  -> cursiva en el nombre de la revista.
      * Repositorio/tesis  -> cursiva en el titulo de la obra.
      * Documentacion web   -> cursiva en el titulo de la obra.
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
        # El titulo del articulo va plano; el nombre de la revista, en cursiva.
        mrev = REVISTA_RE.search(resto)
        if mrev:
            ini_rev = mrev.start()
            while ini_rev > 0 and resto[ini_rev - 1] not in ".,:":
                ini_rev -= 1
            fin_rev = mrev.end()
            cola = resto[fin_rev:]
            # Incluir el parentesis de edicion/volumen: '(Ed. 164)', '(Vol. 3, n. 2)'
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

    # Libro / informe / tesis / documentacion: cursiva en el titulo de la obra.
    if TESIS_RE.search(cuerpo):
        # En una tesis la parte en cursiva abarca titulo + institucion.
        cortes = list(re.finditer(r"\.\s+(?=[A-Z\u00c1\u00c9\u00cd\u00d3\u00da])", resto))
        for k, c in enumerate(cortes[:2]):
            fin = c.start() + 1
            if k == 1 or re.match(r"\s*(Universidad|Instituto|Corporaci\u00f3n|Instituto)", resto[c.end():c.end() + 16]):
                return resto[:fin].strip(), "tesis"
    titulo = resto.split(". ")[0]
    return titulo.strip(" .,"), "obra"


# ---------------------------------------------------------------------------
# Analisis principal
# ---------------------------------------------------------------------------

RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
# Titulo APA de nivel 4 ("**Texto.** el resto del parrafo") o de nivel 5
# ("***Texto.*** ..."). El grupo 1 dice si es nivel 5 (asteriscos triples),
# el grupo 2 es el texto del titulo, que debe terminar en punto.
RE_TITULO_EN_LINEA = re.compile(r"^(\*{3}|\*\*)(?!\s)(.+?)\.\1\s+(\S.*)$", re.S)
RE_IMAGEN = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
RE_NUM_HEADING = re.compile(r"^\s*(\d{1,2})[\.\)]\s+(.*)$")
RE_FIG_CAPTION = re.compile(r"^\s*(?:figura|imagen|image|ilustraci\u00f3n)\s*(\d+)\s*[.:\-\u2013]?\s*(.*)$", re.I)
RE_TAB_CAPTION = re.compile(r"^\s*(?:tabla|cuadro|table)\s*(\d+)\s*[.:\-\u2013]?\s*(.*)$", re.I)
RE_NRC = re.compile(r"\bNRC\s*[:\s]*\s*([0-9]{4,}[-â€“A-Za-z0-9]*)", re.I)
RE_DATE = re.compile(r"(\d{1,2})\s+de\s+([a-z\u00e1\u00e9\u00ed\u00f3\u00fa\u00f1]+)\s+de\s+(\d{4})", re.I)

CAMPOS_PORTADA = [
    ("materia_nrc", RE_NRC),
    ("docente", re.compile(r"profesor|profesora|docente|teacher", re.I)),
    ("facultad", re.compile(r"facultad\s+de", re.I)),
    ("vicerrectoria", re.compile(r"vicerrector", re.I)),
    ("fecha", RE_DATE),
]


def cargar_whitelist(ruta_extra):
    """Terminos intocables. Solo proceden del fichero indicado: no hay base en codigo."""
    terms = set()
    if ruta_extra and os.path.isfile(ruta_extra):
        for ln in read_text(ruta_extra).splitlines():
            ln = ln.strip()
            if ln and not ln.startswith("#"):
                terms.add(ln)
    return sorted(terms, key=len, reverse=True)


# ---------------------------------------------------------------------------
# Enriquecimiento desde MinerU / Docling
# ---------------------------------------------------------------------------
# El contrato real de entrada no es solo "un .md": MinerU y Docling entregan el
# .md de prosa MAS un *_content_list.json con la estructura (tablas con su HTML,
# su bbox y su pagina) y una carpeta images/ con los recortes de cada tabla.
# El .md por si solo pierde justo lo que APA exige nombrar: el titulo de cada
# tabla. Este bloque recupera esa informacion en vez de inventarla.
#
# Que NO se hace aqui, a proposito:
#   - no se reordena el documento con el json (el .md manda en el orden);
#   - no se anaden tablas que el .md no trae, porque las entradas del json con
#     0 filas son detecciones vacias de Docling, no tablas perdidas;
#   - no se deduce un titulo de una tabla que no lo tiene: se deja None y se
#     pregunta, como con el resto de huecos.

RE_ETIQUETA_TABLA = re.compile(r"^\s*(tabla|cuadro|table)\s*[n\.#]?\s*\d*\s*[:.\-]?\s*",
                              re.I)

# Quita el numero de una leyenda que ya lo trae en el .md, para que el .docx no
# lo duplique al anteponer el suyo ("Figura 1. Figura 1. ...").
RE_NUM_LEYENDA = re.compile(
    r"^\s*(figura|fig|imagen|table|tabla|cuadro|chart|grafico|grafico\s*\d*)\s*"
    r"(n[.°º]?\s*)?\d+\s*[:.\)\-–]\s*", re.I)


def _primer_texto_util(campos):
    """Primer string no vacio de los captions/footnotes del json.

    OJO con la forma del dato: `table_caption` y `table_footnote` son listas de
    strings, asi que `[e.get("table_footnote")]` seria una lista dentro de una
    lista y ningun elemento seria un string. Se aplana a proposito, porque asi se
    lee desde el json y asi se veria al anadirle un caption a mano.
    """
    for c in campos or []:
        if isinstance(c, str) and c.strip():
            return c.strip()
        if isinstance(c, (list, tuple)):
            for sub in c:
                if isinstance(sub, str) and sub.strip():
                    return sub.strip()
    return None


def limpiar_titulo_json(txt, whitelist, aplicar_deglue):
    """Normaliza un titulo que viene del json (viene con la misma mugre que el
    resto del .md: pegado, sin espacio tras ':' y a veces con una etiqueta)."""
    t = unescape_md(strip_leading_markers(txt))
    if aplicar_deglue:
        t, _ = deglue(t, whitelist, True)
    t = fix_colon_spacing(t).strip()
    t = RE_ETIQUETA_TABLA.sub("", t).strip()
    return t or None


def _firmas_tabla(filas):
    """Huella de una tabla para emparejarla con su entrada del json: la primera
    fila completa. Dos tablas distintas casi nunca comparten las 3 primeras
    celdas, y ese dato sobrevive aunque el resto de la tabla se haya partido."""
    if not filas or not filas[0]:
        return ""
    return norm_key(" ".join(filas[0][:3]))


def _entradas_tabla_json(cl):
    """Entradas del content_list.json que son tablas CON contenido."""
    out = []
    for it in cl:
        if not isinstance(it, dict) or it.get("type") != "table":
            continue
        body = it.get("table_body") or ""
        if "<td" not in body:
            continue          # entrada vacia: deteccion sin contenido
        out.append(it)
    return out


def resolver_tamano_pagina(md_path, docling_json, pdf, log):
    """(ancho_pt, alto_pt, origen). El bbox del json viene en puntos de la
    pagina, y para pasarlo a pulgadas hay que saber cuanto mide la pagina."""
    if pdf and os.path.isfile(pdf):
        try:
            import pymupdf
            d = pymupdf.open(pdf)
            r = d[0].rect
            log("Tamano de pagina desde el PDF: %.0f x %.0f pt" % (r.width, r.height))
            return r.width, r.height, "pdf"
        except Exception as e:                      # pragma: no cover
            log("AVISO  no se pudo leer el PDF (%s); se continua con otro metodo." % e)

    # Docling guarda page_size por pagina en el *_middle.json
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
                            log("Tamano de pagina desde %s: %.0f x %.0f pt"
                                % (os.path.basename(c), ps[0], ps[1]))
                            return float(ps[0]), float(ps[1]), "middle.json"
                except Exception as e:              # pragma: no cover
                    log("AVISO  %s ilegible (%s)." % (os.path.basename(c), e))

    log("AVISO  no hay PDF ni middle.json: se asume Carta 612x792 pt para "
        "convertir las cajas. Si el original era A4, las tablas quedaran un poco "
        "estrechas; se puede forzar pasando --pdf.")
    return 612.0, 792.0, "supuesto-carta"


def enriquecer_desde_docling(tablas, figuras, docling_json, base_dir, md_path,
                             pdf, ancho_max, whitelist, log, aplicar_deglue=True):
    """Completa tablas y figuras con el json de MinerU/Docling."""
    if not docling_json:
        return {"usado": False}

    if not os.path.isfile(docling_json):
        log("AVISO  no existe el json indicado: %s" % docling_json)
        return {"usado": False}

    try:
        cl = json.loads(read_text(docling_json))
    except Exception as e:
        log("AVISO  %s ilegible: %s" % (os.path.basename(docling_json), e))
        return {"usado": False}

    log("Enriquecimiento desde %s (%d entradas)"
        % (os.path.basename(docling_json), len(cl)))

    ancho_pt, alto_pt, origen_pagina = resolver_tamano_pagina(
        md_path, docling_json, pdf, log)

    espacio = _detectar_espacio_bbox(cl, ancho_pt, alto_pt)
    log("Espacio de coordenadas del bbox: %s%s" % (
        espacio,
        " (normalizado 0..1000; se calibra con page_size)" if espacio == "normalizado-1000" else ""))

    jt = _entradas_tabla_json(cl)
    log("Tablas con contenido en el json: %d (de %d entradas 'table')"
        % (len(jt), sum(1 for x in cl if isinstance(x, dict) and x.get("type") == "table")))

    # Emparejamiento por huella de la primera fila; si no cuaja, por orden, que
    # es lo que corresponde porque ambos ficheros salen del mismo PDF.
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
                             whitelist, log, aplicar_deglue, espacio, es_titulo=True)

    # Las que no se pudieron emparejar por firma: se emparejan por orden, que es
    # correcto salvo que el .md traiga subconjunto de las tablas.
    if orphans and jt:
        libres = [e for e in jt if id(e) not in usados][:len(orphans)]
        for tb, e in zip(orphans, libres):
            usados.add(id(e))
            log("AVISO  la tabla de la linea %d se emparejo con la del json por "
                "orden, no por contenido." % tb.get("linea", 0))
            _aplicar_entrada(tb, e, base_dir, ancho_pt, alto_pt, ancho_max,
                             whitelist, log, aplicar_deglue, espacio, es_titulo=True)

    # Figuras: MinerU/Docling solo las declaran si el documento las trae.
    jf = [x for x in cl if isinstance(x, dict)
          and x.get("type") in ("image", "picture", "figure")]
    for k, e in enumerate(jf):
        ruta = e.get("img_path") or e.get("image_path") or e.get("path")
        if not ruta:
            continue
        abs_ruta = ruta if os.path.isabs(ruta) else os.path.join(base_dir, ruta)
        if not os.path.isfile(abs_ruta):
            log("AVISO  la figura %d apunta a %s, que no existe." % (k + 1, ruta))
            continue
        fig = {
            "ruta": ruta,
            "ruta_absoluta": abs_ruta,
            "titulo": _primer_texto_util([e.get("caption"), e.get("text")]),
            "nota": None,
            "origen": "docling-json",
            "linea": e.get("page_idx"),
        }
        _medir(fig, e.get("bbox"), ancho_pt, alto_pt, ancho_max, espacio, es_figura=True)
        figuras.append(fig)

    if jf:
        log("Figuras declaradas por el json: %d" % len(figuras))

    return {
        "usado": True,
        "json": os.path.basename(docling_json),
        "pagina_pt": [ancho_pt, alto_pt],
        "tamano_pagina_origen": origen_pagina,
        "bbox_espacio": espacio,
        "ancho_max_tabla_in": ancho_max,
        "tablas_json": len(jt),
        "tablas_enriquecidas": len(tablas) - len(orphans) if orphans else len(tablas),
    }


def _aplicar_entrada(tb, e, base_dir, ancho_pt, alto_pt, ancho_max, whitelist,
                     log, aplicar_deglue, espacio, es_titulo=True):
    cap = _primer_texto_util([e.get("table_caption")])
    foot = _primer_texto_util([e.get("table_footnote")])
    # Solo el caption es titulo de la tabla. El footnote es el texto que va DEBAJO
    # de ella, y en la practica MinerU le pone ahi el encabezado de la seccion que
    # agrupa varias tablas ("EJERCICIO 2: ..."), no el nombre de la tabla: usarlo
    # como titulo inventaria un nombre que el documento no tiene. Se guarda como
    # nota, que es lo que es.
    if es_titulo and not tb.get("titulo"):
        limpio = limpiar_titulo_json(cap, whitelist, aplicar_deglue) if cap else None
        if limpio:
            tb["titulo"] = limpio
            tb["titulo_origen"] = "json"
            log("Titulo de la tabla (linea %d) tomado del json: %s"
                % (tb.get("linea", 0), limpio[:60]))
        else:
            log("El json no trae caption para la tabla de la linea %d; se pregunta "
                "al usuario." % tb.get("linea", 0))
    if foot and not tb.get("nota"):
        nota = limpiar_titulo_json(foot, whitelist, aplicar_deglue)
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
    """MinerU/Docling no siempre guardan el bbox en la misma unidad.

    En `content_list.json` el bbox viene NORMALIZADO a 0..1000 en cada eje
    (comprobado contra `middle.json`, que si esta en puntos: los mismos bloques
    dan 612*1.63 = 998 y 792*1.26 = 999). En `middle.json` el bbox esta en
    puntos reales de la pagina.

    La diferencia no es academica: leida como puntos, una caja de 890 de ancho
    son 12,4 pulgadas, mas que una hoja de 8,5, y todas las tablas saldrian
    Deliberadamente deformadas. Se decide con los propios datos en vez de
    suponer un formato.
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


def _medir(obj, bbox, ancho_pt, alto_pt, ancho_max, espacio, es_figura):
    """bbox [l, t, r, b] -> ancho/alto en pulgadas, con topes.

    De una tabla solo se toma el ANCHO. Su alto no sirve: cuando una tabla se
    parte entre dos paginas, Docling guarda una entrada por fragmento, asi que
    la altura de la caja es la del trozo de esa pagina, no la de la tabla. El
    ancho si es fiable, y es lo unico que hace falta para no estirarla.
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
        # Una figura no debería pasar de la mitad de la altura útil: 9 pulgadas
        # de alto en medio de un cuerpo de texto no hay quien las lea.
        obj["alto_in"] = round(min(h_in, (alto_pt - 144) / 72.0), 3)


def analizar(md_path, base_dir, whitelist, log, aplicar_deglue=True,
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

        # --- tabla HTML (puede ocupar una o varias lineas) ---------------
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
                avisos.append("Tabla HTML sin celdas en la linea %d; se omite." % (i + 1))
            i = j + 1
            continue

        # --- tabla markdown ---------------------------------------------
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

        # --- encabezado -------------------------------------------------
        mh = RE_HEADING.match(t)
        if mh:
            nivel_md = len(mh.group(1))
            texto_h = mh.group(2).strip()

            if nivel_md == 1:
                # El titulo del documento: es el titulo de la portada, no una seccion.
                portada_lineas.append({"tipo": "titulo", "texto": unescape_md(texto_h)})
                bloques.append({"tipo": "portada", "texto": unescape_md(texto_h)})
                i += 1
                continue

            limpio, camb = deglue(strip_leading_markers(unescape_md(texto_h)), whitelist, aplicar_deglue)
            limpio = fix_colon_spacing(dedupe_numbering(limpio)).strip()
            cambios_deglue.extend(camb)

            # Â¿es la seccion de referencias?
            if re.search(r"referencias?\s+bibliogr", norm_key(limpio)) or norm_key(limpio).startswith("referencias"):
                en_referencias = True

            mnum = RE_NUM_HEADING.match(limpio)
            if nivel_md >= 3:
                # APA 7 tiene 5 niveles. '##' cubre los dos primeros segun este
                # documento (numerada -> nivel 1, sin numerar -> nivel 2) y los
                # hashes mas profundos mapean a su nivel natural. Antes TODO
                # acababa en nivel 2, asi que un .md con '###' salia aplanado y
                # los niveles 3, 4 y 5 no existian en la salida.
                nivel = min(nivel_md, 5)
                etiqueta = None
                texto_h2 = limpio
            elif mnum:
                # Seccion principal numerada: "3. Analisis de BSIMM" -> nivel 1
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

        # --- imagen markdown -------------------------------------------
        if "![" in t and "](" in t:
            mimg = RE_IMAGEN.search(t)
            if mimg:
                alt, ruta = mimg.group(1).strip(), mimg.group(2).strip()
                # una sola linea puede traer la imagen y su leyenda
                cola = t[mimg.end():].strip()
                figs = []
                for tag, num, resto in RE_FIG_CAPTION.findall(cola):
                    figs.append((num, resto))
                titulo = alt
                if figs:
                    titulo = figs[0][1] or titulo
                elif cola and not cola.startswith("|"):
                    titulo = alt or cola
                # El .docx antepone su propio "Figura N", asi que si la leyenda
                # del .md ya lo trae ("Figura 1. Grafico de barras") hay que quitarlo
                # o en el PDF sale "Figura 1. Figura 1. Grafico de barras".
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

        # --- referencias -------------------------------------------------
        if en_referencias:
            limpia = strip_leading_markers(unescape_md(t))
            limpia = fix_colon_spacing(limpia).strip()
            if limpia:
                referencias.append({"texto": limpia, "linea": i + 1})
            i += 1
            continue

        # --- portada: lineas antes de la primera seccion ----------------
        if not primera_seccion_vista:
            limpia = strip_leading_markers(unescape_md(t))
            if limpia:
                portada_lineas.append({"tipo": "dato", "texto": limpia})
                bloques.append({"tipo": "portada_dato", "texto": limpia})
            i += 1
            continue

        # --- parrafo / lista --------------------------------------------
        crudo = unescape_md(t)

        # --- titulo APA de nivel 4 o 5 "en linea" (opt-in) ---------------
        # APA 7 los define asi: "**Titulo.** El texto del parrafo sigue en la
        # misma linea", con sangria de primera linea. Markdown no puede expresar
        # esa estructura, asi que se reconoce por convencion. Es opt-in porque la
        # negrita al abrir un parrafo es tambien texto normal corriente y
        # activarlo sin control inundaria la TOC de titulos falsos.
        #
        # Se matchea sobre el texto CRUDO a proposito: strip_leading_markers()
        # se come los asteriscos iniciales ("**Titulo.**" -> "Titulo.") y si se
        # detectara despues ya no habria forma de distinguirlo del texto normal.
        if detectar_en_linea and not en_referencias:
            m_en = RE_TITULO_EN_LINEA.match(crudo.strip())
            if m_en and i + 1 < n:
                titulo_en = m_en.group(2).strip()
                # El cuerpo es el grupo 3, NO lo que va tras m_en.end(): el
                # grupo 3 ES el cuerpo y m_en.end() cae justo despues de el.
                cuerpo_en, camb = deglue(m_en.group(3).strip(), whitelist,
                                         aplicar_deglue)
                cambios_deglue.extend(camb)
                cuerpo_en = fix_colon_spacing(dedupe_numbering(cuerpo_en)).strip()
                hid += 1
                # Tres asteriscos de entrada = nivel 5; dos = nivel 4.
                nivel_en = 5 if m_en.group(1) == "***" else 4
                bloques.append({
                    "tipo": "h", "hid": hid,
                    "nivel": nivel_en,
                    "texto": titulo_en,
                    "en_linea": cuerpo_en,
                    # Los segmentos los resuelve el parser y no el .docx: el
                    # analisis de markdown en linea (negrita, cursiva, codigo)
                    # tiene una sola implementacion.
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
        limpia, camb = deglue(limpia, whitelist, aplicar_deglue)
        cambios_deglue.extend(camb)
        limpia = fix_colon_spacing(dedupe_numbering(limpia)).strip()

        if limpia:
            # conservar el marcador de lista si lo hay
            mnum_item = re.match(r"^(\d+[\.\)]|[" + BULLET_CHARS + r"])\s+", limpia)
            bloques.append({
                "tipo": "lista" if mnum_item else "p",
                "texto": limpia,
                "marcador": (mnum_item.group(1) if mnum_item else None),
                "segmentos": parse_inline(limpia),
            })
        i += 1

    # --- leyendas de tabla: pueden ir ANTES o DESPUES de la tabla ---------
    # Docling y MinerU no coinciden: uno las emite antes del <table> y el otro
    # despues. Se aceptan ambos ordenes en una pasada posterior, que es mas
    # fiable que decidirlo durante el recorrido linea a linea.
    for k, b in enumerate(bloques):
        if b.get("tipo") != "p" or "segmentos" not in b:
            continue
        texto = b["texto"]
        mt = RE_TAB_CAPTION.match(texto)
        if not mt:
            continue
        # tabla justo despues
        if k + 1 < len(bloques) and bloques[k + 1].get("tipo") == "tabla":
            destino = bloques[k + 1]["indice"]
        # tabla justo antes
        elif k > 0 and bloques[k - 1].get("tipo") == "tabla":
            destino = bloques[k - 1]["indice"]
        else:
            continue
        if tablas[destino]["titulo"] is None:
            tablas[destino]["titulo"] = fix_colon_spacing(mt.group(2)).strip().rstrip(":")
            b["tipo"] = "nota_tabla"
            b["indice"] = destino
            b.pop("segmentos", None)

    # --- numeracion de tablas y figuras ---------------------------------
    for k, tb in enumerate(tablas):
        tb["indice"] = k + 1
        if not tb["titulo"]:
            avisos.append("La tabla %d no trae titulo en el .md: hay que redactarlo o "
                          "preguntar al usuario." % (k + 1))
    for k, fg in enumerate(figuras):
        fg["indice"] = k + 1
        ruta_abs = fg["ruta"]
        if not os.path.isabs(ruta_abs):
            ruta_abs = os.path.join(base_dir, ruta_abs)
        fg["ruta_absoluta"] = os.path.normpath(ruta_abs)
        fg["existe"] = os.path.isfile(fg["ruta_absoluta"])
        if not fg["existe"]:
            avisos.append("La figura %d apunta a '%s', que no existe en disco. La "
                          "imagen NO se insertara." % (k + 1, fg["ruta"]))
        if not fg["titulo"]:
            avisos.append("La figura %d no trae leyenda en el .md: hay que redactarla o "
                          "preguntar al usuario." % (k + 1))

    # --- cursivas propuestas en las referencias -------------------------
    for ref in referencias:
        rango, tipo = propose_reference_italics(ref["texto"])
        if rango:
            ref["cursiva_propuesta"] = rango
            ref["tipo_heuristica"] = tipo
        else:
            ref["cursiva_propuesta"] = None
            ref["tipo_heuristica"] = None
            ref["aviso"] = "no se pudo determinar la parte en cursiva (%s)" % tipo
        ref["cursiva_confirmada"] = ref["cursiva_propuesta"]
        ref["revisar_cursiva"] = True
    if referencias:
        avisos.append(
            "Las %d referencias vienen del .md y su parte en cursiva es una PROPUESTA "
            "heuristica: hay que revisarlas una a una antes de entregar "
            "(references/normas-apa7.md, seccion Referencias)." % len(referencias))

    # --- titulos duplicados --------------------------------------------
    dups = {k: v for k, v in dup_titulos.items() if len(v) > 1}
    if dups:
        lista = "; ".join(sorted({v[0] for v in dups.values()}))
        avisos.append("Hay %d titulos de seccion repetidos en el documento (%s). La TOC "
                      "los mostrara tal cual: preguntar al usuario si quiere "
                      "distinguirlos o renombrarlos." % (len(dups), lista))

    # --- residuos de pegado que el deglue NO puede resolver ---------------
    # El caso 'SegunelMarcoBSIMM' uni dos palabras MINUSCULAS, asi que la regla
    # minuscula+M mayuscula no las ve. Adivinar donde va cada espacio exigiria un
    # diccionario del idioma: es mas honesto reportar el residuo y que el agente
    # lo corrija en el .md de origen en vez de producir texto corrupto.
    PEGADO_RE = re.compile(r"[A-Za-z\u00c0-\u00ff]{18,}")
    residuos = {}
    for b in bloques:
        txt = b.get("texto")
        if not txt:
            continue
        for m in PEGADO_RE.finditer(txt):
            if " " in m.group(0):
                continue
            # los terminos de la lista blanca son largamente pegados por diseÃ±o
            if any(w.lower() in m.group(0).lower() for w in whitelist if len(w) >= 8):
                continue
            residuos[m.group(0)] = residuos.get(m.group(0), 0) + 1
    if residuos:
        # Se ordenan por longitud decreciente, no por frecuencia: un tramo de 40
        # caracteres pegados es mucho mas sospechoso que uno de 20, y la palabra
        # larga y valida ("internacionalmente") es el falso positivo tipico.
        # Subir el umbral de PEGADO_RE para evitarlo perderia de verdad 29
        # residuos reales del documento de referencia, que es peor canje.
        top = sorted(residuos.items(), key=lambda kv: (-len(kv[0]), -kv[1]))[:12]
        avisos.append(
            "Hay %d tramos sin espacios que el deglue no puede separar con "
            "confianza (pegado de palabras en minuscula). NO se corrigieron "
            "automaticamente: revisarlos en el .md de origen. Entre los mas "
            "largos (puede alguno ser una palabra valida sin espacios): %s"
            % (len(residuos), ", ".join("%s (x%d)" % (t, c) for t, c in top)))
        preguntas_residuo = ["Revisar en el .md: '%s' (x%d, %d caracteres). Si es "
                             "una palabra pegada, separarla; si es una palabra "
                             "valida, dejarla."
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
# Portada: deteccion automatica + huecos que hay que preguntar
# ---------------------------------------------------------------------------

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def _es_nombre_persona(txt):
    """'Nombre Apellido Segundo' -> True. Exige 2 o mas palabras capitalizadas."""
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
    Intenta deducir los campos de portada del .md. Lo que no se deduzca
    NO se inventa: se lista en 'campos_faltantes' para que el agente pregunte.

    Formato tipico en documentos MinerU/Docling de trabajos academicos:
        # Titulo del trabajo
        Nombre Apellido, Nombre Apellido
        NRC: 00000000
        PROFESOR: Nombre Apellido
        Facultad de <facultad>, <universidad>
            <sede / unidad academica>
        <dia> de <mes> de <ano>

    Los campos institucionales (facultad, sede, NRC) se detectan por patron y
    son opcionales: si el documento no los trae, se preguntan. Nunca se inventan.
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
            # 'Nombre Apellido Segundo NRC: 00000000' -> autor + NRC separados
            resto = RE_NRC.sub("", txt).strip(" :-â€“,")
            portada["materia_nrc"] = "NRC " + mn.group(1)
            if _es_nombre_persona(resto):
                portada["autores"] = [p.strip() for p in resto.split(",")]
            continue

        mdoc = re.search(r"(profesor|profesora|docente|teacher)\s*[:\-]?\s*(.*)$", txt, re.I)
        if mdoc and not portada["docente"]:
            nombre = mdoc.group(2).strip(" :-â€“,")
            partes = [p.strip() for p in nombre.split(",") if p.strip()]
            portada["docente"] = partes[0] if partes else nombre
            if len(partes) > 1 and partes[1]:
                portada["docente_titulo"] = partes[1]
            continue

        # La linea de institucion suele traer facultad + vicerrectoria + sede.
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

    # el overrides del usuario (lo que respondio al preguntar) manda siempre
    for k, v in (overrides or {}).items():
        if v not in (None, "", []):
            portada[k] = v

    # 'autores' es el unico campo de portada que es lista. Si el usuario lo
    # entrega como cadena, meterla tal cual haria que build-docx.js la recorriera
    # caracter a caracter y la portada saldria con una letra por linea. Se
    # respeta literalmente lo escrito (no se parte por 'y' ni por comas: eso
    # seria inventar autores). Para uno por linea, pasar un array JSON.
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

    # El titulo profesional del docente SIEMPRE se pregunta (regla de la skill),
    # aunque el resto de la portada este completo.
    if not portada.get("docente_titulo"):
        faltan.append("docente_titulo")

    # El logo tambien se pregunta SIEMPRE. Es opcional en cuanto al resultado,
    # pero no se puede omitir en silencio: sin esta linea el generador no tiene
    # forma de distinguir "el usuario no tiene logo" de "a nadie le preguntamos".
    if not portada.get("logo"):
        faltan.append("logo")

    portada["campos_faltantes"] = faltan
    return portada


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description="Convierte un .md en MANIFIEST.json")
    ap.add_argument("--md", required=True, help="ruta del .md fuente")
    ap.add_argument("--out", required=True, help="ruta del MANIFIEST.json a escribir")
    ap.add_argument("--log", help="fichero de log del analisis")
    ap.add_argument("--base-dir", help="directorio base para resolver rutas de imagenes")
    ap.add_argument("--imagenes", help="imagenes.json de calcular_tamano_imagenes.py")
    ap.add_argument("--portada", help="portada.json con las respuestas del usuario")
    ap.add_argument("--whitelist", help="fichero extra de terminos intocables")
    ap.add_argument("--sin-deglue", action="store_true", help="no separar palabras pegadas")
    ap.add_argument("--titulos-tabla-json", help="JSON {indice: titulo} del agente")
    ap.add_argument("--titulos-figura-json", help="JSON {indice: titulo} del agente")
    ap.add_argument("--notas-tabla-json",
                    help="JSON {indice: nota} del agente; por defecto 'Elaboracion propia'")
    ap.add_argument("--notas-figura-json",
                    help="JSON {indice: nota} del agente; por defecto 'Elaboracion propia'")
    ap.add_argument("--detectar-niveles-en-linea", action="store_true",
                    help="tratar un parrafo que empieza por '**Texto.**' como titulo "
                         "de nivel 4 y '***Texto.***' como nivel 5 (APA 7 en linea). "
                         "Opt-in: sin la opcion, negrita al abrir parrafo es texto normal. "
                         "Desactivado por defecto porque genera falsos positivos.")
    ap.add_argument("--sin-indice-tablas", action="store_true")
    ap.add_argument("--sin-indice-figuras", action="store_true")
    ap.add_argument("--docling-json", dest="docling_json",
                    help="content_list.json de MinerU/Docling para completar titulos, "
                         "cajas (bbox) e imagenes de tabla. Si se omite se busca solo "
                         "el <nombre>_content_list.json junto al .md")
    ap.add_argument("--pdf", help="PDF original: da el tamano de pagina real para "
                                  "convertir bbox en pulgadas")
    ap.add_argument("--ancho-max-tabla", type=float, default=6.5,
                    help="ancho util en pulgadas (Letter con margenes de 1 pulgada)")
    args = ap.parse_args()

    log_lines = []

    def log(msg):
        line = "[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg)
        log_lines.append(line)
        print(line)

    if not os.path.isfile(args.md):
        print("ERROR: no existe el .md: %s" % args.md, file=sys.stderr)
        return 2

    base_dir = args.base_dir or os.path.dirname(os.path.abspath(args.md))
    whitelist = cargar_whitelist(args.whitelist)

    log("=== FASE 1: analisis del .md ===")
    log("Fuente        : %s" % args.md)
    log("sha256        : %s" % sha256_of(args.md))
    log("Tamano        : %d bytes" % os.path.getsize(args.md))
    log("Dir base      : %s" % base_dir)
    log("Terminos en lista blanca: %d" % len(whitelist))
    log("")

    # MinerU/Docling entregan el .md de prosa junto a un *_content_list.json con la
    # estructura. Si esta a mano, se usa solo: es el caso normal de esos flujos.
    docling_json = args.docling_json
    if docling_json is None:
        raiz = os.path.splitext(os.path.abspath(args.md))[0]
        for cand in (raiz + "_content_list.json", raiz + "_content_list_v2.json"):
            if os.path.isfile(cand):
                docling_json = cand
                log("Json de MinerU/Docling detectado solo: %s" % os.path.basename(cand))
                break

    res = analizar(args.md, base_dir, whitelist, log,
                   aplicar_deglue=not args.sin_deglue,
                   detectar_en_linea=args.detectar_niveles_en_linea)

    info_docling = {}
    if docling_json:
        info_docling = enriquecer_desde_docling(
            res["tablas"], res["figuras"], docling_json, base_dir, args.md,
            args.pdf, args.ancho_max_tabla, whitelist, log,
            aplicar_deglue=not args.sin_deglue)
        log("")

    log("--- Inventario ---")
    log("Secciones (encabezados) : %d" % len(res["secciones"]))
    log("Tablas                  : %d" % len(res["tablas"]))
    log("Figuras                 : %d" % len(res["figuras"]))
    log("Referencias             : %d" % len(res["referencias"]))
    log("Lineas de portada       : %d" % len(res["portada_lineas"]))
    log("")

    if res["cambios_deglue"]:
        log("--- Correccion de palabras pegadas (%d cambios) ---" % len(res["cambios_deglue"]))
        vistos = set()
        for tipo, antes, despues in res["cambios_deglue"]:
            clave = (antes, despues)
            if clave in vistos:
                continue
            vistos.add(clave)
            log("  %-9s %-34s -> %s" % (tipo, antes, despues))
        log("")

    log("--- Secciones detectadas ---")
    for s in res["secciones"]:
        log("  H%-5d [n%d] %s" % (s["hid"], s["nivel"], s["texto"][:88]))
    log("")

    log("--- Tablas detectadas ---")
    for tb in res["tablas"]:
        log("  Tabla %-2d %d filas x %d cols  titulo=%s  (linea %d, %s)"
            % (tb["indice"], len(tb["filas"]), len(tb["filas"][0]),
               (tb["titulo"] or "SIN TITULO"), tb["linea"], tb["origen"]))
    log("")

    if res["figuras"]:
        log("--- Figuras detectadas ---")
        for fg in res["figuras"]:
            log("  Figura %-2d %s  existe=%s  titulo=%s"
                % (fg["indice"], fg["ruta"], fg["existe"], (fg["titulo"] or "SIN TITULO")))
        log("")

    if res["referencias"]:
        log("--- Referencias detectadas (orden alfabetico se aplica en el .docx) ---")
        for ref in res["referencias"]:
            log("  - %s" % ref["texto"][:110])
            log("      cursiva propuesta: %s  [%s]  REVISAR" %
                (ref["cursiva_propuesta"], ref.get("tipo_heuristica") or "?"))
        log("")

    # --- titulos que aporta el agente -----------------------------------
    if args.titulos_tabla_json:
        tt = read_json(args.titulos_tabla_json)
        for k, v in tt.items():
            for tb in res["tablas"]:
                if tb["indice"] == int(k):
                    tb["titulo"] = v
        log("Titulos de tabla sobrescritos desde %s" % args.titulos_tabla_json)
    if args.titulos_figura_json:
        tf = read_json(args.titulos_figura_json)
        for k, v in tf.items():
            for fg in res["figuras"]:
                if fg["indice"] == int(k):
                    fg["titulo"] = v
        log("Titulos de figura sobrescritos desde %s" % args.titulos_figura_json)

    # --- notas de tablas y figuras ----------------------------------------
    if args.notas_tabla_json:
        nt = read_json(args.notas_tabla_json)
        for k, v in nt.items():
            for tb in res["tablas"]:
                if tb["indice"] == int(k):
                    tb["nota"] = v
                    tb["nota_origen"] = "agente"
        log("Notas de tabla sobrescritas desde %s" % args.notas_tabla_json)
    if args.notas_figura_json:
        nf = read_json(args.notas_figura_json)
        for k, v in nf.items():
            for fg in res["figuras"]:
                if fg["indice"] == int(k):
                    fg["nota"] = v
                    fg["nota_origen"] = "agente"
        log("Notas de figura sobrescritas desde %s" % args.notas_figura_json)

    # references/normas-apa7.md: "Por defecto 'Nota. Elaboracion propia.',
    # salvo que el .md indique otra fuente". Ese default es obligatorio porque
    # una tabla o figura sin nota no cumple APA. Antes solo se imprimia la nota
    # si el .md o el json la traian, de modo que en un documento corriente no
    # salia ninguna y el verificador no lo detectaba: se entregaba sin ella.
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

    log("--- Portada detectada ---")
    for k in ("titulo", "autores", "facultad", "vicerrectoria", "materia_nrc",
              "docente", "docente_titulo", "fecha", "logo"):
        log("  %-15s : %s" % (k, portada.get(k)))
    log("  CAMPOS FALTANTES (preguntar al usuario, NO inferir): %s"
        % (", ".join(portada["campos_faltantes"]) or "(ninguno)"))
    log("")

    # --- tamanos de imagen ----------------------------------------------
    if args.imagenes and os.path.isfile(args.imagenes):
        imgs = read_json(args.imagenes)
        log("--- Tamanos de imagen recibidos (%d) ---" % len(imgs))
        by_path = {}
        for it in imgs:
            by_path[norm_key(os.path.basename(it.get("ruta", "")))] = it
        for fg in res["figuras"]:
            clave = norm_key(os.path.basename(fg["ruta"]))
            hit = by_path.get(clave)
            if hit:
                fg["ancho_in"] = hit.get("ancho_in")
                fg["alto_in"] = hit.get("alto_in")
                log("  Figura %-2d %-46s -> %.2f x %.2f in"
                    % (fg["indice"], os.path.basename(fg["ruta"]),
                       hit.get("ancho_in", 0), hit.get("alto_in", 0)))
            else:
                fg["ancho_in"] = 6.5
                fg["alto_in"] = None
                avisos_fg = "sin medida en imagenes.json: se usa el ancho de contenido"
                fg["aviso_medida"] = avisos_fg
                log("  Figura %-2d %-46s -> SIN MEDIDA (ancho por defecto 6.5 in)"
                    % (fg["indice"], os.path.basename(fg["ruta"])))
        log("")

    # --- politica de indices --------------------------------------------
    hay_tablas = len(res["tablas"]) > 0
    hay_figuras = len(res["figuras"]) > 0
    idx_tablas = hay_tablas and not args.sin_indice_tablas
    idx_figuras = hay_figuras and not args.sin_indice_figuras

    log("--- Politica de indices ---")
    log("  Tabla de contenido : SIEMPRE")
    log("  Indice de tablas  : %s (tablas encontradas: %d)"
        % ("SI" if idx_tablas else "NO", len(res["tablas"])))
    log("  Indice de figuras : %s (figuras encontradas: %d)"
        % ("SI" if idx_figuras else "NO", len(res["figuras"])))
    if not idx_tablas:
        log("  -> Sin tablas en el .md. PREGUNTAR al usuario si el documento no tiene tablas,")
        log("     y NO generar un indice de tablas vacio.")
    if not idx_figuras:
        log("  -> Sin figuras en el .md. PREGUNTAR al usuario si el documento no tiene figuras,")
        log("     y NO generar un indice de figuras vacio.")
    log("")

    # --- avisos ----------------------------------------------------------
    if res["avisos"]:
        log("--- Avisos (%d) ---" % len(res["avisos"]))
        for a in res["avisos"]:
            log("  * %s" % a)
        log("")

    # --- preguntas al usuario ---------------------------------------------
    # Cada pregunta declara si BLOQUEA la entrega. Las bloqueantes no se
    # resuelven inventando: se preguntan y el documento no se entrega hasta que
    # el usuario responda. Antes eran cadenas sueltas en el log, sin forma de
    # distinguirlas de un aviso, y el documento salia con "(sin titulo)".
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

    # Campos de portada que se PREGUNTAN pero NO bloquean: el documento se
    # construye igual sin ellos. Es coherente con los avisos (critico=False) que
    # emite verificar-pdf.py por los mismos dos campos.
    PORTADA_SIN_BLOQUEAR = {"logo", "docente_titulo"}
    PREGUNTA_PORTADA = {
        "logo": ("No hay logo en la portada. Si lo aporta se coloca encima del "
                 "titulo. ¿Se incluye o se omite?"),
        "docente_titulo": "¿Cual es el titulo o profesion del docente?",
    }

    for c in portada["campos_faltantes"]:
        preguntar("portada_%s" % c, c not in PORTADA_SIN_BLOQUEAR, c,
                  "agregar_u_omitir", "Portada", None,
                  PREGUNTA_PORTADA.get(c, "Falta '%s' en la portada: se agrega o se omite?" % c),
                  ["agregar", "omitir"],
                  "Anotar la decision en portada.json y volver a ejecutar el "
                  "parser con --portada.")
    if not idx_tablas:
        preguntar("sin_tablas", True, "estructura", "confirmacion", "Documento",
                  None, "El .md no contiene tablas: confirma que el documento "
                        "no tiene tablas.",
                  ["confirmar", "falta_alguno"],
                  "Si falta alguna tabla, revisar el .md antes de continuar.")
    if not idx_figuras:
        preguntar("sin_figuras", True, "estructura", "confirmacion", "Documento",
                  None, "El .md no contiene figuras: confirma que el documento "
                        "no tiene figuras.",
                  ["confirmar", "falta_alguno"],
                  "Si falta alguna figura, revisar el .md antes de continuar.")

    OPCION_TITULO = ["redactar_desde_contexto", "usuario_proporciona"]
    for tb in res["tablas"]:
        if not tb["titulo"]:
            preguntar(
                "titulo_tabla_%d" % tb["indice"], True, "titulo",
                "eleccion_binaria", "Tabla", tb["indice"],
                "La Tabla %d no tiene titulo. ¿Redacto el titulo a partir del "
                "contexto del informe, o me lo proporcionas tu?" % tb["indice"],
                OPCION_TITULO,
                "Si eliges 'redactar_desde_contexto': escribir el titulo en "
                "titulos-tabla.json como {\"%d\": \"...\"} y volver a ejecutar el "
                "parser con --titulos-tabla-json." % tb["indice"])
    for fg in res["figuras"]:
        if not fg["titulo"]:
            preguntar(
                "titulo_figura_%d" % fg["indice"], True, "titulo",
                "eleccion_binaria", "Figura", fg["indice"],
                "La Figura %d no tiene leyenda. ¿Redacto la leyenda a partir del "
                "contexto del informe, o me la proporcionas tu?" % fg["indice"],
                OPCION_TITULO,
                "Si eliges 'redactar_desde_contexto': escribir la leyenda en "
                "titulos-figura.json como {\"%d\": \"...\"} y volver a ejecutar el "
                "parser con --titulos-figura-json." % fg["indice"])

    if res["referencias"]:
        preguntar("cursiva_referencias", False, "referencias", "revision",
                  "Referencias", None,
                  "Revisar la parte en cursiva de las %d referencias (heuristica)."
                  % len(res["referencias"]))
    # Los residuos los detecta el analizador pero no se copiaban ni al manifiesto
    # ni a la lista de preguntas: el agente se enteraba del problema por el
    # console, no por la salida que realmente consume.
    for n, pr in enumerate(res["preguntas_residuo"]):
        preguntar("residuo_%d" % n, False, "deglue", "revision_manual",
                  "Palabras pegadas", None, pr)

    preguntas_texto = [q["pregunta"] for q in preguntas]
    pendientes_bloqueantes = [q["id"] for q in preguntas if q["bloqueante"]]


    log("--- Preguntas pendientes para el usuario (%d, %d bloqueantes) ---"
        % (len(preguntas), len(pendientes_bloqueantes)))
    for q in preguntas:
        log("  ? [%s] %s" % ("BLOQUEA" if q["bloqueante"] else "aviso", q["pregunta"]))
    if pendientes_bloqueantes:
        log("")
        log("  PENDIENTES BLOQUEANTES: %s" % ", ".join(pendientes_bloqueantes))
        log("  El documento NO se puede entregar hasta resolverlas.")
    log("")

    # --- manifiesto ------------------------------------------------------
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

    log("MANIFIEST escrito: %s (%d bytes)" % (args.out, os.path.getsize(args.out)))

    if args.log:
        log_dir = os.path.dirname(os.path.abspath(args.log))
        if log_dir and not os.path.isdir(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        with open(args.log, "w", encoding="utf-8") as fh:
            fh.write("\n".join(log_lines) + "\n")
        log("Log escrito: %s" % args.log)

    return 0


if __name__ == "__main__":
    sys.exit(main())


