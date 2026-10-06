"""Which engine turns a .docx into a PDF: Microsoft Word first, LibreOffice as backup.

STANDARD LIBRARY ONLY, same reason as everything else under lib/: `check` runs
before anything is installed.

The order is a decision, not an accident. Word is what the user already has and
what their supervisor will open the PDF in, so a document that has to be checked
page by page is best rendered by the same program that will display it. Word is
also the engine whose pagination matches Word's own page numbers. But Word is
only preferred when it is PROVEN to work here (lib/word.py), never because an
executable exists.

LibreOffice stays as the fallback for three concrete reasons: it is free and
scriptable on every platform, it is what the engine was for years, and it has no
requirement that the user's word processor be closed while it runs.
"""

from dataclasses import dataclass

from . import rutas
from . import word as word_motor

WORD = "word"
LIBREOFFICE = "libreoffice"
AUTO = "auto"

# The order is also the preference order for `auto`.
MOTORES = (WORD, LIBREOFFICE)
PREFIJOS = (AUTO,) + MOTORES


@dataclass
class Motor:
    """One engine's state, plus the sentence a human needs to understand it."""

    nombre: str
    ok: bool
    version: str = ""
    detalle: str = ""
    motivo: str = ""
    # True when the answer came from a real export probe, not from a lookup.
    # "Installed" and "can export" are different questions and this says which
    # one was asked, so a caller can say it out loud instead of implying more
    # than it knows.
    sondeado: bool = False

    def etiqueta(self):
        """Short name for logs: 'Microsoft Word 16.0' / 'headless LibreOffice 7.x'."""
        base = {"word": "Microsoft Word", "libreoffice": "headless LibreOffice"}[self.nombre]
        return ("%s %s" % (base, self.version)).strip()


def normaliza(valor):
    """'Word', 'WORD', 'auto', '' -> the canonical name. ValueError otherwise.

    The accepted spellings are deliberately forgiving (case, a leading dash, an
    empty value) because this string reaches us from a command line a human or an
    agent typed, and a typo must produce a clear message, not a silent fallback to
    a different engine than the one that was asked for.
    """
    texto = (valor or "").strip().strip("-_ ").lower()
    alias = {
        "": AUTO,
        "microsoft word": WORD,
        "word": WORD,
        "winword": WORD,
        "office": WORD,
        "lo": LIBREOFFICE,
        "soffice": LIBREOFFICE,
        "headless": LIBREOFFICE,
        "default": AUTO,
    }
    texto = alias.get(texto, texto)
    if texto not in PREFIJOS:
        raise ValueError("unknown PDF engine %r: use --motor %s"
                         % (valor, ", ".join(PREFIJOS)))
    return texto


def estado_word(probar=True, timeout=word_motor.TIMEOUT_SONDEO, reevaluar=False):
    """Microsoft Word as an engine. `probar=False` answers the cheap question only."""
    estado = word_motor.disponible(probar=probar, timeout=timeout, reevaluar=reevaluar)
    motor = Motor(nombre=WORD, ok=estado.ok, version=estado.version,
                  sondeado=estado.sondeado)
    if estado.ok:
        if estado.sondeado:
            motor.detalle = "%s  (%s, probe: minimal .docx -> PDF)" % (
                motor.etiqueta(), estado.ruta)
        else:
            motor.detalle = "%s  (%s, not verified)" % (motor.etiqueta(), estado.ruta)
    else:
        motor.motivo = estado.motivo
        motor.detalle = estado.motivo
    return motor


def estado_libreoffice():
    """Headless LibreOffice as an engine.

    The version is always queried through run_soffice, never by calling the
    launcher directly: soffice.exe detaches, the child inherits the output pipe
    and the caller waits forever. That hangs the whole preflight.
    """
    soffice = rutas.soffice_path()
    if not soffice:
        buscados = [p for p in rutas.soffice_candidates() if p]
        return Motor(nombre=LIBREOFFICE, ok=False,
                     motivo="not found. Searched: %s"
                            % (" | ".join(buscados)
                               or "(no candidates: set APA7_SOFFICE)"))
    try:
        probed = rutas.run_soffice(["--version"], timeout=60)
    except rutas.SofficeNotFound as exc:
        return Motor(nombre=LIBREOFFICE, ok=False, motivo=str(exc))

    first = ""
    for line in (probed.stdout or "").replace("\r\n", "\n").split("\n"):
        if line.strip():
            first = line.strip()
            break
    if probed.exit_code != 0 or not first:
        return Motor(nombre=LIBREOFFICE, ok=False,
                     motivo="could not read the version (exit code %d)"
                            % probed.exit_code)
    version = first.split(" ", 1)[1].strip() if " " in first else first
    return Motor(nombre=LIBREOFFICE, ok=True, version=version,
                 detalle="%s  (%s)" % (Motor(LIBREOFFICE, True, version).etiqueta(), soffice))


def elegir(preferido=AUTO, probar_word=True, timeout=word_motor.TIMEOUT_SONDEO,
           reevaluar=False):
    """`(chosen, all_states)`.

    `chosen` is the Motor to use, or None when nothing can convert a PDF here.
    `all_states` is every engine that was looked at, in preference order, so the
    caller can explain the decision instead of just acting on it.

    `preferido` naming an engine that is NOT viable returns None rather than
    quietly using the other one: an explicit `--motor word` must never be
    answered with LibreOffice, or the user gets a PDF from a renderer they did
    not ask for and no way to know.
    """
    preferido = normaliza(preferido)

    if preferido == WORD:
        estado = estado_word(probar=probar_word, timeout=timeout, reevaluar=reevaluar)
        return (estado if estado.ok else None), [estado]

    if preferido == LIBREOFFICE:
        estado = estado_libreoffice()
        return (estado if estado.ok else None), [estado]

    # auto: Word first, LibreOffice second. Word that turns out not to be viable
    # must not end the search: the whole point of having two engines is that the
    # common case (no Word on a Linux or CI box) still produces a PDF.
    estados = []
    estado_w = estado_word(probar=probar_word, timeout=timeout, reevaluar=reevaluar)
    estados.append(estado_w)
    if estado_w.ok:
        return estado_w, estados
    estado_l = estado_libreoffice()
    estados.append(estado_l)
    return (estado_l if estado_l.ok else None), estados
