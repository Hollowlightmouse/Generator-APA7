"""Microsoft Word as the .docx -> .pdf engine.

STANDARD LIBRARY ONLY, and no syntax newer than Python 3.9, for the same reason
as apa7.py: `check` runs BEFORE anything is installed. pywin32 is not an option
even on Windows, because `check` cannot assume anything is there yet.

How Word is driven
    Word has no command line that converts a document. On Windows it is driven
    through COM (PowerShell + `New-Object -ComObject Word.Application`); on
    macOS through AppleScript. The script travels base64-encoded
    (`-EncodedCommand`) so no path ever gets quoted into it by accident: the
    script is uploaded whole, with the caller's own `'` characters doubled.

Why viability is a PROBE and not a lookup
    "Word is installed" is cheap to answer and is not the question that matters.
    What matters is whether Word on THIS machine can really turn a .docx into a
    PDF: COM can be registered and still fail on activation, on first-run
    licensing dialogs, or on a locked-down profile. So viability is proven the
    only honest way, by exporting a minimal .docx to a PDF and looking at the
    result. That costs a few seconds, which is why the answer is cached in
    `.work/motor-word.json` and only re-run with --recheck-motor.

The one rule that is not negotiable
    Word is a single-instance COM server. `New-Object -ComObject
    Word.Application` ATTACHES to the instance the user already has open, and
    Quit-ing it would close their documents. So an export refuses to start while
    Word is already running, and only ever quits the instance it launched
    itself.

macOS status: implemented but NOT verified on real hardware. Every path behind
    it is probe-gated, so an untested AppleScript degrades to "Word not viable"
    instead of breaking the pipeline. `auto` then falls back to LibreOffice and
    says why in the log.
"""

import base64
import json
import os
import shutil
import subprocess
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from . import rutas

MOTOR = "word"

# Same conventions the LibreOffice engine already uses: 124 is "timed out" and
# 127 is "could not be launched", so the exporter can report both the same way.
TIMEOUT_SONDEO = 90
TIMEOUT_EXPORTACION = 300

# Word constants used below, spelled out so the driver script is self-contained.
WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0
WD_ALERTS_NONE = 0

_SENTINEL_OK = "APA7-PDF-OK"
_SENTINEL_ERROR = "APA7-PDF-ERROR"
_SENTINEL_VERSION = "APA7-VERSION:"

_CACHE_NAME = "motor-word.json"
_CACHE_VERSION = 1


@dataclass
class Estado:
    """What is known about Word on this machine, and how it was learned."""

    ok: bool = False
    version: str = ""
    motivo: str = ""
    ruta: str = ""
    registrada: bool = False
    sondeado: bool = False
    desde_cache: bool = False


@dataclass
class Resultado:
    """Outcome of one export attempt (or of the probe, which is the same code)."""

    exit_code: int = 0
    stdout: str = ""
    stderr: str = ""
    timed_out: bool = False
    launch_error: str = None
    pdf_creado: bool = False
    motivo: str = ""
    version: str = ""
    ruta_ejecutable: str = ""
    log_dir: str = None
    stdout_path: str = None
    stderr_path: str = None
    pids_restantes: list = field(default_factory=list)
    transitorio: bool = False
    # True when the attempt was skipped on purpose (--dry-run) instead of
    # failing: nothing was proven and nothing was disproven, which is not the
    # same thing as a failure and must never be reported as a probe.
    omitido: bool = False


# ---------------------------------------------------------------------------
# The platform: can Word be driven here at all?
# ---------------------------------------------------------------------------
def powershell_bin():
    """Path of PowerShell, or None. Windows PowerShell first: `pwsh` is not on
    every machine, and the COM support of Windows PowerShell 5.1 is the one
    that is actually installed with Windows."""
    for name in ("powershell", "pwsh"):
        found = shutil.which(name)
        if found:
            return found
    return None


def plataforma_soportada():
    """None when Word can be driven from here, or the reason it cannot.

    Deliberately conservative: it only answers "is there a driver for this OS at
    all", never "is Word installed", so it stays a pure, cheap fact.
    """
    if rutas.is_windows():
        if not powershell_bin():
            return "PowerShell not found (needed to drive Word through COM)"
        return None
    if rutas.is_macos():
        if not shutil.which("osascript"):
            return "osascript not found (needed to drive Word through AppleScript)"
        return None
    return "Microsoft Word cannot be automated on this platform"


# ---------------------------------------------------------------------------
# The driver scripts
# ---------------------------------------------------------------------------
def _comilla(valor):
    """A PowerShell/AppleScript single-quoted string literal for `valor`."""
    return "'" + str(valor).replace("'", "''") + "'"


def _literal(valor):
    return str(valor).replace("\\", "\\\\").replace("'", "\\'")


def script_powershell(docx, pdf, actualizar_campos=False):
    """The PowerShell that opens `docx` invisibly and exports it to `pdf`.

    Every COM object is released in `finally`, including on the error path, so a
    failed export cannot leave a headless WINWORD.EXE behind holding the file
    open. `exit 1` inside `try` still runs `finally`: that is what keeps the
    document and the application from leaking.

    `[ref]` on Close/Quit is not decoration. Word declares those optional
    parameters as by-reference VARIANTs, so PowerShell's late binding REFUSES a
    bare literal ("must be System.Management.Automation.PSReference. Use
    [ref].") and the exception is swallowed by the surrounding `catch {}`. The
    visible effect is a WINWORD.EXE that stays alive forever after a successful
    export, which then makes every later run refuse because "Word is open".
    """
    actualizar = ("[void]$documento.Fields.Update()\n    " if actualizar_campos else "")
    return (
        "$ErrorActionPreference = 'Stop'\n"
        "$cero = %d\n"
        "$palabra = $null\n"
        "$documento = $null\n"
        "try {\n"
        "  $palabra = New-Object -ComObject Word.Application\n"
        "  $palabra.Visible = $false\n"
        "  $palabra.DisplayAlerts = %d\n"
        "  Write-Output ('%s' + $palabra.Version)\n"
        "  $documento = $palabra.Documents.Open(%s, $false, $true, $false, "
        "'', '', $false, '', '', 0, 0, $false)\n"
        "  %s"
        "  $documento.ExportAsFixedFormat(%s, %d)\n"
        "  Write-Output '%s'\n"
        "} catch {\n"
        "  Write-Output ('%s: ' + $_.Exception.Message)\n"
        "  exit 1\n"
        "} finally {\n"
        "  if ($null -ne $documento) { try { $documento.Close([ref]$cero) } catch {} }\n"
        "  if ($null -ne $palabra) { try { $palabra.Quit([ref]$cero) } catch {} }\n"
        "}\n"
    ) % (
        WD_DO_NOT_SAVE_CHANGES,
        WD_ALERTS_NONE,
        _SENTINEL_VERSION,
        _comilla(docx),
        actualizar,
        _comilla(pdf),
        WD_EXPORT_FORMAT_PDF,
        _SENTINEL_OK,
        _SENTINEL_ERROR,
    )


def script_applescript(docx, pdf, actualizar_campos=False):
    """The AppleScript equivalent (macOS, unverified on hardware).

    `read only true` plus `close ... saving no` keeps the user's document
    untouched, and `quit saving no` only ever quits the instance this script
    launched -- which is why word.py refuses to run it while Word is open. The
    sentinels are the same ones the PowerShell driver prints, so the result of
    an export is interpreted the same way on both platforms.
    """
    actualizar = ("            update fields of active document\n" if actualizar_campos else "")
    return (
        "try\n"
        '    tell application "Microsoft Word"\n'
        "        with timeout of %d seconds\n"
        "            set laVersion to version\n"
        "            set elDocumento to open file name %s with read only\n"
        "%s"
        "            save as elDocumento file name %s file format format PDF\n"
        "            close elDocumento saving no\n"
        "        end timeout\n"
        "        quit saving no\n"
        "    end tell\n"
        "    return \"%s\" & laVersion & \"\\n%s\"\n"
        "on error mensaje number numero\n"
        "    return \"%s: \" & mensaje & \" (\" & numero & \")\"\n"
        "end try\n"
    ) % (
        max(30, TIMEOUT_EXPORTACION),
        _literal(docx),
        actualizar,
        _literal(pdf),
        _SENTINEL_VERSION,
        _SENTINEL_OK,
        _SENTINEL_ERROR,
    )


def comando(docx, pdf, actualizar_campos=False):
    """The argv that performs one export, or None when this platform has no driver."""
    if rutas.is_windows():
        powershell = powershell_bin()
        if not powershell:
            return None
        script = script_powershell(docx, pdf, actualizar_campos)
        codificado = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        return [powershell, "-NoProfile", "-NonInteractive", "-EncodedCommand", codificado]
    if rutas.is_macos():
        osascript = shutil.which("osascript")
        if not osascript:
            return None
        script = script_applescript(docx, pdf, actualizar_campos)
        return [osascript, "-e", script]
    return None


# ---------------------------------------------------------------------------
# Running it, with a real time bound and no inherited pipes
# ---------------------------------------------------------------------------
def _correr(argv, timeout, out_file, err_file):
    """Run `argv`, writing stdout/stderr to files. Exit 124 = timed out, 127 = could not launch.

    Files rather than pipes for the same reason run_soffice uses them: a child
    that inherits a redirected handle can outlive us and leave the pipeline
    hanging, and Word in particular spawns processes of its own.
    """
    exit_code = 0
    timed_out = False
    launch_error = None

    with open(str(out_file), "w", encoding="utf-8", errors="replace") as out_handle, \
            open(str(err_file), "w", encoding="utf-8", errors="replace") as err_handle:
        try:
            process = subprocess.Popen(
                argv,
                stdout=out_handle,
                stderr=err_handle,
                stdin=subprocess.DEVNULL,
                start_new_session=(not rutas.is_windows()),
            )
        except OSError as exc:
            launch_error = str(exc)
            exit_code = 127
        else:
            try:
                exit_code = process.wait(timeout=max(1, int(timeout)))
            except subprocess.TimeoutExpired:
                timed_out = True
                kill_word()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    pass
                exit_code = 124

    def _leer(path):
        try:
            return Path(path).read_text(encoding="utf-8", errors="replace")
        except OSError:
            return ""

    return exit_code, _leer(out_file), _leer(err_file), timed_out, launch_error


def kill_word(pids=None):
    """Kill Word processes (ours by default). Never raises: cleanup is best effort."""
    if pids is None:
        pids = rutas.word_pids()
    for pid in pids or ():
        try:
            rutas.kill_process_tree(pid)
        except Exception:      # noqa: BLE001 - cleanup must never break an export
            pass


def _normalizar(texto):
    return (texto or "").replace("\r\n", "\n")


def _interpretar(stdout, exit_code, pdf, timed_out):
    """Turn what the driver printed into (ok, version, motivo)."""
    salida = _normalizar(stdout)
    version = ""
    for linea in salida.split("\n"):
        limpia = linea.strip()
        if limpia.startswith(_SENTINEL_VERSION):
            version = limpia[len(_SENTINEL_VERSION):].strip()

    for linea in salida.split("\n"):
        limpia = linea.strip()
        if limpia.startswith(_SENTINEL_ERROR):
            # The driver prints "APA7-PDF-ERROR: <message>"; the ": " belongs to
            # the sentinel line, not to the message the user has to read.
            return False, version, limpia[len(_SENTINEL_ERROR):].lstrip(": ").strip() or "unknown error"

    if timed_out:
        return False, version, "Microsoft Word did not finish within the timeout"
    if _SENTINEL_OK not in salida:
        return False, version, "the driver produced no PDF (exit code %d)" % exit_code
    if not _es_pdf(pdf):
        return False, version, "the driver reported success but wrote no readable PDF"
    return True, version, ""


def _es_pdf(pdf):
    try:
        with open(str(pdf), "rb") as handle:
            return handle.read(5).startswith(b"%PDF-")
    except OSError:
        return False


def _esperar_que_word_salga(intentos=5, espera=1.0):
    """Give Word a moment to finish quitting, and report whatever is left.

    Quit() returns as soon as the request is made; WINWORD.EXE usually disappears
    a second or two later. Polling instead of assuming is what keeps the engine
    from declaring a leftover and killing a process that was on its way out.
    Each poll costs a shell-out, so this is a few tries, not a long poll.
    """
    for _ in range(max(1, intentos)):
        if not rutas.word_pids():
            return []
        time.sleep(max(0.0, espera))
    return rutas.word_pids()


def _sin_rastro_word(antes, log_dir=None):
    """Word processes that appeared during our run and are still there.

    Any PID here is one this run started (the engine refuses to start while Word
    is already open, so `antes` is empty), which is why killing them is safe.
    """
    time.sleep(1.0)
    restantes = _esperar_que_word_salga()
    if restantes and log_dir:
        rutas.log("Word left running after the export; stopping it (PID %s)"
                  % ", ".join(str(pid) for pid in restantes))
    return restantes


# ---------------------------------------------------------------------------
# The export itself
# ---------------------------------------------------------------------------
def exportar(docx, pdf, timeout=TIMEOUT_EXPORTACION, actualizar_campos=False, log_dir=None):
    """Export `docx` to `pdf` with Word. Returns a Resultado.

    The PDF is produced under a sibling temporary name and moved into place only
    after it is proven to be a PDF, so a failed attempt can never destroy the
    previous PDF or leave a truncated file behind.
    """
    docx = Path(docx).expanduser().resolve()
    pdf = Path(pdf).expanduser().resolve()

    if not docx.is_file():
        return Resultado(exit_code=1, motivo="%s does not exist" % docx)

    sin_plataforma = plataforma_soportada()
    if sin_plataforma:
        return Resultado(exit_code=1, motivo=sin_plataforma)

    abiertas = rutas.word_pids()
    if abiertas:
        # Word is single-instance: attaching would mean quitting the user's own
        # Word at the end. Refuse instead of risking their unsaved documents.
        return Resultado(
            exit_code=1,
            motivo="Microsoft Word is already open (PID %s). Close it and export again; "
                   "the engine never touches a Word session you started."
                   % ", ".join(str(pid) for pid in abiertas),
            ruta_ejecutable=rutas.word_path() or "",
            # A Word the user opened is not a verdict on Word: whoever looks this
            # up again must probe for real instead of reading a cached failure.
            transitorio=True,
        )

    # The PDF is produced under a sibling temporary name and moved into place
    # only after it is proven to be a PDF, so a failed attempt can never destroy
    # the previous PDF or leave a truncated file behind.
    destino = pdf.with_name(pdf.stem + ".word.tmp.pdf")
    try:
        destino.unlink()
    except OSError:
        pass
    pdf.parent.mkdir(parents=True, exist_ok=True)
    argv = comando(docx, destino, actualizar_campos)
    if argv is None:
        return Resultado(exit_code=1, motivo="no driver available for this platform")

    own_dir = log_dir is None
    run_dir = (Path(tempfile.gettempdir()) / ("apa7-word-%d" % os.getpid())
               if own_dir else Path(log_dir).expanduser().resolve())
    run_dir.mkdir(parents=True, exist_ok=True)
    out_file = run_dir / "word.out.log"
    err_file = run_dir / "word.err.log"

    exit_code, stdout, stderr, timed_out, launch_error = _correr(
        argv, timeout, out_file, err_file)
    ok, version, motivo = _interpretar(stdout, exit_code, destino, timed_out)
    if launch_error:
        ok, motivo = False, "could not launch the Word driver: %s" % launch_error

    pdf_creado = False
    if ok:
        try:
            os.replace(str(destino), str(pdf))
            pdf_creado = True
        except OSError as exc:
            ok, motivo = False, "could not move the exported PDF into place: %s" % exc
    try:
        destino.unlink()
    except OSError:
        pass

    restantes = _sin_rastro_word(abiertas, run_dir)
    # Unconditional: these PIDs are ours (the export refused to start while Word
    # was open), so a leftover is stopped whether the export worked or not.
    if restantes:
        kill_word(restantes)

    return Resultado(
        exit_code=0 if ok else 1,
        stdout=stdout,
        stderr=stderr,
        timed_out=timed_out,
        launch_error=launch_error,
        pdf_creado=pdf_creado,
        motivo=motivo,
        version=version,
        ruta_ejecutable=rutas.word_path() or "",
        log_dir=str(run_dir),
        stdout_path=str(out_file),
        stderr_path=str(err_file),
        pids_restantes=restantes,
    )


# ---------------------------------------------------------------------------
# Viability: the probe
# ---------------------------------------------------------------------------
# A .docx is a zip with three parts and nothing else. Building it here instead of
# shipping a binary fixture means no .docx is ever committed (they are all
# gitignored as build products) and the probe cannot go stale.
_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)
_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)
_DOCUMENTO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p><w:r><w:t>APA7</w:t></w:r></w:p></w:body></w:document>"
)


def docx_minimo(destino):
    """Write a valid, minimal .docx to `destino`. Returns the path."""
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(destino), "w", zipfile.ZIP_DEFLATED) as paquete:
        # [Content_Types].xml goes first: the OPC spec expects it at the head of
        # the package, and some readers are stricter than Word about it.
        paquete.writestr("[Content_Types].xml", _CONTENT_TYPES)
        paquete.writestr("_rels/.rels", _RELS)
        paquete.writestr("word/document.xml", _DOCUMENTO)
    return destino


def _limpiar_scratch(scratch, intentos=3):
    """Delete the probe's scratch directory, retrying while Word still holds it.

    Word releases its files a moment after Quit() returns, and rmtree is
    deliberately called with ignore_errors, so a single attempt can leave a
    `apa7-sondeo-word-*` directory in %TEMP% with a PDF in it. Retrying for a
    second costs nothing; leaking one directory per run would be noise the user
    finds later and cannot explain, so a directory that survives that is said
    out loud instead of being swallowed.
    """
    for intento in range(max(1, intentos)):
        if rutas.remove_tree(scratch):
            return True
        if intento + 1 < intentos:
            time.sleep(0.5)
    rutas.log("The probe could not delete its temporary directory %s; "
              "it is safe to remove it by hand." % scratch)
    return False


def sondear(timeout=TIMEOUT_SONDEO):
    """Prove that Word can really export, by exporting a one-word document.

    A Resultado with exit_code 0 is the only answer that means "viable".
    """
    if rutas.dry_run():
        return Resultado(exit_code=1, omitido=True,
                         motivo="skipped: --dry-run creates and launches nothing")

    scratch = Path(tempfile.mkdtemp(prefix="apa7-sondeo-word-"))
    try:
        docx = docx_minimo(scratch / "sondeo.docx")
        return exportar(docx, scratch / "sondeo.pdf", timeout=timeout)
    finally:
        _limpiar_scratch(scratch)


# ---------------------------------------------------------------------------
# The cached answer
# ---------------------------------------------------------------------------
def _cache_path():
    return rutas.workdir() / _CACHE_NAME


def _clave(ruta):
    """Identity of the Word build. An upgrade changes it and invalidates the cache."""
    try:
        info = Path(ruta).stat()
        return "%s|%d|%d" % (ruta, int(info.st_mtime), info.st_size)
    except OSError:
        return ruta


def _cargar_cache(clave):
    try:
        data = json.loads(_cache_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    if data.get("version") != _CACHE_VERSION or data.get("clave") != clave:
        return None
    return data


def _guardar_cache(clave, estado):
    if rutas.dry_run():
        return
    try:
        _cache_path().write_text(
            json.dumps({"version": _CACHE_VERSION, "clave": clave, "ok": estado.ok,
                        "word": estado.version, "motivo": estado.motivo,
                        "fecha": time.strftime("%Y-%m-%dT%H:%M:%S")},
                       ensure_ascii=False, indent=2),
            encoding="utf-8")
    except OSError:
        pass


def disponible(probar=True, timeout=TIMEOUT_SONDEO, reevaluar=False):
    """Is Word usable as the PDF engine here? Returns an Estado.

    Three escalating steps, cheapest first:
      1. is there an executable,
      2. does this machine know the COM class / AppleScript application,
      3. can it really export (the probe), cached against the Word build.

    `probar=False` stops after step 2 and says so (`sondeado=False`), for
    callers that want the cheap answer without launching anything.
    """
    sin_plataforma = plataforma_soportada()
    if sin_plataforma:
        return Estado(ok=False, motivo=sin_plataforma)

    ruta = rutas.word_path()
    if not ruta:
        buscados = rutas.word_candidates()
        return Estado(ok=False, motivo="not installed. Searched: %s"
                      % (", ".join(buscados) if buscados else "no known location"))

    registrada = rutas.word_clase_registrada()
    if registrada is False:
        return Estado(ok=False, ruta=ruta,
                      motivo="installed but the Word.Application automation class is "
                             "not registered, so it cannot be driven")

    if not probar:
        return Estado(ok=True, ruta=ruta, registrada=bool(registrada))

    clave = _clave(ruta)
    if not reevaluar:
        cache = _cargar_cache(clave)
        if cache:
            return Estado(ok=bool(cache.get("ok")), version=cache.get("word", ""),
                          motivo=cache.get("motivo", ""), ruta=ruta, registrada=True,
                          sondeado=True, desde_cache=True)

    resultado = sondear(timeout=timeout)
    if resultado.omitido:
        # A skipped probe is not a failed probe. Reporting it as either "not
        # viable" or "viable, sondeado" would be a claim nobody can back up, so
        # it comes back as unknown-with-a-reason and nothing is cached.
        return Estado(ok=False, motivo=resultado.motivo, ruta=ruta, registrada=True,
                      sondeado=False)
    estado = Estado(ok=resultado.exit_code == 0 and resultado.pdf_creado,
                    version=resultado.version, ruta=ruta, registrada=True, sondeado=True)
    if not estado.ok:
        estado.motivo = resultado.motivo or "the probe export failed"
    if resultado.transitorio:
        # Nothing was proven and nothing was disproven (the user had Word open).
        # Caching this would make the next run report a failure that no longer
        # exists, so the answer is simply left uncached.
        estado.sondeado = False
        return estado
    _guardar_cache(clave, estado)
    return estado
