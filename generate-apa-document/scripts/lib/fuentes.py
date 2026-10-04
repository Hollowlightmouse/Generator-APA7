"""fuentes.py - which embedded font names are acceptable in the PDF.

When the requested font is not available, LibreOffice silently substitutes a
metrically compatible one. That is what happened on the Linux and macOS
machines this was checked on: the substitute shares advance widths with Times
New Roman, so line breaks, page breaks and the line advance are unaffected, but
the embedded font name is different. The old check only accepted names
containing "timesnewroman", so every correct document whose machine lacked the
real font was reported as using a font it was not allowed to use.

A separate module because verificar-pdf.py can only be imported where pymupdf
exists, and this rule needs no PDF and no third-party package to be tested.
"""

import re

# Normalised (lowercase, letters only) names accepted as Times New Roman.
#
# These are the families listed as metric-compatible with Times (or Times New
# Roman) in the usual substitution tables: Liberation Serif (Red Hat), Nimbus
# Roman (URW), TeX Gyre Termes (GUST), FreeSerif (GNU) and Tinos (Google's
# Croscore serif, metric-compatible with Times New Roman; its sans partner
# Arimo is the Arial-compatible one). They share advance widths with Times, so
# line and page breaks are unaffected.
#
# Matched as PREFIXES, because the embedded name carries the style:
# "LiberationSerif-Bold" normalises to "liberationserifbold".
METRICOS = (
    "timesnewroman",
    "liberationserif",
    "nimbusroman",      # URW clone, ships with Ghostscript
    "texgyretermes",
    "freeserif",
    "tinos",            # Croscore serif, metric-compatible with Times New Roman
)

# Fonts that are known and deliberately NOT accepted. They are a different
# family with different metrics, so the line advance measured in the PDF would
# no longer be comparable with the Times New Roman values the check assumes.
# They get their own message because "unknown font" would send the reader
# looking for a substitution problem that does not exist.
CONOCIDAS_NO_METRICAS = {
    "dejavuserif": "DejaVu Serif",
    "liberationsans": "Liberation Sans",
    "arial": "Arial",
    "calibri": "Calibri",
    "cambria": "Cambria",
    "georgia": "Georgia",
    "verdana": "Verdana",
    "couriernew": "Courier New",
    "consolas": "Consolas",
}


def normaliza(nombre):
    """Strip the PDF subset prefix and reduce to comparable letters only.

    PDFs embed fonts as "BAAAAA+TimesNewRomanPS-BoldMT", so both the prefix
    and the separators have to go before anything can be compared.
    """
    nombre = re.sub(r"^[A-Z]{6}\+", "", nombre or "")
    return re.sub(r"[^a-z]", "", (nombre or "").lower())


def clasifica(nombre):
    """Return (aceptada, es_sustitucion, nombre_legible).

    aceptada        -- the font may appear in the PDF.
    es_sustitucion  -- True when it is a substitute for Times New Roman rather
                       than Times New Roman itself. Worth surfacing: the reader
                       should know the machine did not have the real font.
    """
    limpio = normaliza(nombre)
    if not limpio:
        return False, False, nombre or ""

    if any(limpio.startswith(conocida) for conocida in METRICOS):
        # Times New Roman itself is not a substitution.
        sustitucion = not limpio.startswith("timesnewroman")
        return True, sustitucion, nombre

    for conocido, legible in CONOCIDAS_NO_METRICAS.items():
        if limpio.startswith(conocido):
            return False, False, legible

    return False, False, nombre


def particiona(nombres):
    """Split embedded font names into (aceptadas, sustituciones, rechazadas).

    Devuelve each rejection with the reason, because "unknown font" and "known
    font with different metrics" call for different fixes.
    """
    aceptadas, sustituciones, rechazadas = [], [], []
    for nombre in nombres:
        ok, sustitucion, legible = clasifica(nombre)
        if ok:
            (sustituciones if sustitucion else aceptadas).append(nombre)
        elif legible in CONOCIDAS_NO_METRICAS.values():
            rechazadas.append((legible, "known font, but its metrics differ from "
                                      "Times New Roman, so the line advance "
                                      "measured here cannot be compared"))
        else:
            rechazadas.append((legible, "not a declared or metric-compatible font"))
    return aceptadas, sustituciones, rechazadas