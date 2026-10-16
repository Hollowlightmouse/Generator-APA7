"""fuentes.py - which font families the document may use, and their metrics.

APA 7 (and the UNIMINUTO guide) allow ONE typeface for the whole document,
chosen from a short list, each with its own size:

    Times New Roman 12, Arial 11, Calibri 11, Lucida Sans Unicode 10,
    Georgia 11, Computer Modern 10.

Two things depend on the family:

  * which embedded font names are ACCEPTED in the PDF. When the requested font
    is not installed, LibreOffice silently substitutes a metrically compatible
    one that shares its advance widths (line breaks and line advance are
    unaffected), so those substitutes must pass. A font of a DIFFERENT family
    must still fail, because its metrics are not comparable.
  * the expected LINE ADVANCE. The double-spacing check measures the advance in
    points and compares it with the value the chosen family/size should give.
    Times New Roman 12 advances ~13.8 pt at single spacing; each family has its
    own line-height factor, so the threshold is derived from the family instead
    of being a single hardcoded number.

A separate module because verificar-pdf.py can only be imported where pymupdf
exists, and these rules need no PDF and no third-party package to be tested.
"""

import re

# Factor that turns the single-spacing advance into the minimum a double-spaced
# body must show. Calibrated so Times New Roman 12 keeps its historical 20.0 pt
# threshold (20.0 / 13.8 = 1.4493). Anything clearly below single spacing fails.
FACTOR_UMBRAL_DOBLE = 20.0 / 13.8

# The families the guide allows. Key = normalised name (letters only, lowercase).
#   etiqueta   -- the name printed in messages and stored in the manifest.
#   pt         -- the size the guide asks for that family (points).
#   em         -- single-spacing line advance as a multiple of the font size
#                 (Word/LibreOffice single line height). Approximate on purpose:
#                 the double-spacing check has a wide tolerance.
#   sustitutos -- normalised names of metric-compatible substitutes, matched as
#                 PREFIXES (the embedded name carries the style, so
#                 "LiberationSerif-Bold" normalises to "liberationserifbold").
FAMILIAS = {
    "timesnewroman": {
        "etiqueta": "Times New Roman",
        "pt": 12,
        "em": 1.15,
        "sustitutos": ("liberationserif", "nimbusroman", "texgyretermes",
                       "freeserif", "tinos"),
    },
    "arial": {
        "etiqueta": "Arial",
        "pt": 11,
        "em": 1.15,
        "sustitutos": ("liberationsans", "arimo", "helvetica", "nimbussans",
                       "freesans"),
    },
    "calibri": {
        "etiqueta": "Calibri",
        "pt": 11,
        "em": 1.22,
        "sustitutos": ("carlito",),
    },
    "lucidasansunicode": {
        "etiqueta": "Lucida Sans Unicode",
        "pt": 10,
        "em": 1.16,
        "sustitutos": (),
    },
    "georgia": {
        "etiqueta": "Georgia",
        "pt": 11,
        "em": 1.14,
        "sustitutos": ("gelasio",),
    },
    "computermodern": {
        "etiqueta": "Computer Modern",
        "pt": 10,
        "em": 1.20,
        "sustitutos": ("latinmodernroman", "lmroman", "cmuserif",
                       "computermodernroman"),
    },
}

FAMILIA_POR_DEFECTO = "Times New Roman"

# Fonts that are known and deliberately NOT accepted when the declared family is
# something else. They are a different family with different metrics, so the line
# advance measured in the PDF would no longer be comparable. They get their own
# message because "unknown font" would send the reader looking for a substitution
# problem that does not exist.
CONOCIDAS_NO_METRICAS = {
    "dejavuserif": "DejaVu Serif",
    "dejavusans": "DejaVu Sans",
    "liberationsans": "Liberation Sans",
    "liberationserif": "Liberation Serif",
    "arial": "Arial",
    "calibri": "Calibri",
    "cambria": "Cambria",
    "georgia": "Georgia",
    "lucidasans": "Lucida Sans Unicode",
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


def clave_familia(nombre):
    """Normalised key of the declared family, or None if it is not one we know.

    Matching is by PREFIX so that "Times New Roman", "times new roman" and the
    embedded "TimesNewRomanPSMT" all resolve to the same family.
    """
    limpio = normaliza(nombre)
    if not limpio:
        return None
    for clave in FAMILIAS:
        if limpio.startswith(clave):
            return clave
    return None


def es_familia_valida(nombre):
    """True when the name is one of the families the guide allows."""
    return clave_familia(nombre) is not None


def etiqueta_familia(nombre):
    """Printed name of the family, or the input unchanged when unknown."""
    clave = clave_familia(nombre)
    return FAMILIAS[clave]["etiqueta"] if clave else (nombre or "")


def tamano_por_defecto(nombre):
    """Guide size (points) for the family, or None for an unknown one."""
    clave = clave_familia(nombre)
    return FAMILIAS[clave]["pt"] if clave else None


def familias_disponibles():
    """[(label, size_pt)] of every allowed family, in a stable order."""
    return [(datos["etiqueta"], datos["pt"]) for datos in FAMILIAS.values()]


def datos_familia(nombre):
    """The family's dict, falling back to Times New Roman when unknown."""
    clave = clave_familia(nombre) or "timesnewroman"
    return FAMILIAS[clave]


def avance_simple_pt(nombre, pt=None):
    """Expected single-spacing line advance in points for the family/size."""
    datos = datos_familia(nombre)
    tam = pt or datos["pt"]
    return datos["em"] * tam


def umbral_doble(nombre, pt=None):
    """Minimum line advance (points) a double-spaced body of this font shows."""
    return FACTOR_UMBRAL_DOBLE * avance_simple_pt(nombre, pt)


def clasifica(nombre, familia=FAMILIA_POR_DEFECTO):
    """Return (aceptada, es_sustitucion, nombre_legible) for the given family.

    aceptada        -- the font may appear in a document of that family.
    es_sustitucion  -- True when it is a metric-compatible substitute for the
                       declared family rather than the family itself. Worth
                       surfacing: the reader should know the machine did not have
                       the real font.
    """
    limpio = normaliza(nombre)
    if not limpio:
        return False, False, nombre or ""

    clave = clave_familia(familia) or "timesnewroman"

    if limpio.startswith(clave):
        return True, False, nombre

    if any(limpio.startswith(s) for s in FAMILIAS[clave]["sustitutos"]):
        return True, True, nombre

    for conocido, legible in CONOCIDAS_NO_METRICAS.items():
        if limpio.startswith(conocido):
            return False, False, legible

    return False, False, nombre


def particiona(nombres, familia=FAMILIA_POR_DEFECTO):
    """Split embedded font names into (aceptadas, sustituciones, rechazadas).

    Each rejection carries its reason, because "unknown font" and "known font of
    another family" call for different fixes.
    """
    aceptadas, sustituciones, rechazadas = [], [], []
    for nombre in nombres:
        ok, sustitucion, legible = clasifica(nombre, familia)
        if ok:
            (sustituciones if sustitucion else aceptadas).append(nombre)
        elif legible in CONOCIDAS_NO_METRICAS.values():
            rechazadas.append((legible, "known font, but it belongs to another "
                                      "family, so the line advance measured here "
                                      "cannot be compared"))
        else:
            rechazadas.append((legible, "not a declared or metric-compatible font"))
    return aceptadas, sustituciones, rechazadas
