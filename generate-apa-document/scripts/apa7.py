#!/usr/bin/env python3
"""apa7.py - single entry point of the generate-apa-document skill.

Cross-platform replacement for the four PowerShell scripts. Standard library
only, and no syntax newer than Python 3.9, because `check` is the command that
runs BEFORE anything is installed: it is what tells the user what to install,
so it cannot depend on a third-party package or on a newer interpreter.

Subcommands
    check     preflight (STEP 0): every required tool, one line per tool
    export    .docx -> .pdf with Microsoft Word or headless LibreOffice
    install   install whatever is missing
    parse     .md + layout JSON -> MANIFEST.json
    build     MANIFEST.json -> .docx
    verify    .pdf verification with pymupdf

Output contract
    Machine-readable output goes to STDOUT and nothing else does, so that
    SKILL.md can parse it. Progress, warnings and diagnostics go to STDERR.
    A human running the skill by hand still sees them.

    check prints one line per tool:
        OK|<tool>|<detail>
        MISSING|<tool>|<detail>
        INFO|<tool>|<detail>
    followed by RESULT: OK or RESULT: MISSING. Exit code 0 = complete
    environment, 1 = something is missing.

    INFO is new here. The package manager is only needed by `install`, so
    reporting it as MISSING made a perfectly usable environment look broken
    and sent the pipeline down an install that had nothing to do. INFO lines
    never affect RESULT nor the exit code.

The PDF engine
    `--motor auto|word|libreoffice`. `auto` (the default) uses Microsoft Word
    when a probe proves it can really export here, and headless LibreOffice
    otherwise. Word is preferred because it renders the document the way Word
    itself will display it, which is what makes page-by-page verification
    meaningful. Naming an engine that is not usable is an ERROR (exit code 2):
    an explicit request is never answered with a different renderer.

Optional environment variables
    APA7_SOFFICE, APA7_WORD, APA7_PYTHON, APA7_NODEDIR, APA7_WORKDIR,
    APA7_SKILL_ROOT
"""

import argparse
import contextlib
import io
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib import instalador  # noqa: E402
from lib import motores as motores_mod  # noqa: E402
from lib import rutas  # noqa: E402
from lib import word as word_motor  # noqa: E402

_FLOOR = "%d.%d" % rutas.MIN_PYTHON
_FLOOR_TUPLE = "(%d, %d)" % rutas.MIN_PYTHON

# Reports the version and, at the same time, whether it clears the floor, so one
# subprocess answers both questions and the two cannot disagree.
# Built by concatenation on purpose: the %d below belongs to the probe's own
# formatting, so it must not be consumed by the formatting that injects the floor.
_PY_PROBE = (
    "import sys; v = sys.version_info; "
    "ok = v[:2] >= " + _FLOOR_TUPLE + "; "
    "print('%d.%d.%d %s' % (v[0], v[1], v[2], 'OK' if ok else 'TOO-OLD')); "
    "sys.exit(0 if ok else 3)"
)


# ---------------------------------------------------------------------------
# check
# ---------------------------------------------------------------------------
def cmd_check(args=None):
    """Preflight (STEP 0).

    Every line goes to stdout because this output is the contract the skill
    parses; nothing human-facing is mixed in, so `RESULT:` can be read without
    having to filter noise out of the stream.

    `args` is the parsed namespace when the CLI called this, and None when the
    installer called it. Everything read from it is optional and defaults to the
    documented behaviour, so `check` on its own needs no flags at all.
    """
    args = args or argparse.Namespace()
    results = []

    def report(ok, tool, detail):
        state = "OK" if ok else "MISSING"
        results.append((state, tool))
        sys.stdout.write("%s|%s|%s\n" % (state, tool, detail))

    def inform(tool, detail):
        # Never appended to `results`: INFO must not be able to change RESULT.
        sys.stdout.write("INFO|%s|%s\n" % (tool, detail))

    # --- Python + pymupdf ----------------------------------------------------
    # The concrete venv is NOT fixed inside the skill: it is resolved by
    # environment variable or PATH. See references/system-requirements.md.
    python = rutas.python_path()

    if not python:
        # "No usable Python" has two very different causes and they need
        # different fixes, so they are not reported the same way. A venv whose
        # base interpreter was removed or upgraded still EXISTS but cannot run:
        # the executable is there, so a bare "not found" would send the user to
        # install Python, which does nothing for them.
        broken_venv = rutas.venv_python()
        if broken_venv:
            report(False, "Python",
                   "the skill's virtual environment exists but does not run: %s "
                   "(its base interpreter was moved, removed or upgraded). Recreate "
                   "it with: %s" % (broken_venv, rutas.comando_apa7("install")))
        else:
            report(False, "Python",
                   "not found (set APA7_PYTHON to an interpreter that has pymupdf)")
        report(False, "pymupdf", "not verifiable: Python missing")
    else:
        probed = rutas.run([python, "-c", _PY_PROBE])
        version = probed.first_line.split(" ")[0] if probed.first_line else "?"
        if probed.exit_code != 0 or "TOO-OLD" in probed.first_line:
            report(False, "Python",
                   "%s is below the %s floor required by the skill  (%s)"
                   % (version, _FLOOR, python))
        else:
            report(True, "Python", "%s  (%s)" % (version, python))

        pymupdf = rutas.run([python, "-c", "import pymupdf; print(pymupdf.__version__)"])
        ok = pymupdf.exit_code == 0
        # A version that is not the pinned one is reported INSIDE the detail and
        # the state stays OK: the library imports and works, it is simply not the
        # tested one. The documented contract stays untouched.
        note = ""
        if ok and pymupdf.first_line and pymupdf.first_line != rutas.DEPS["pymupdf"]:
            note = "  [MISMATCH: pinned %s]" % rutas.DEPS["pymupdf"]
        report(ok, "pymupdf",
               "%s%s  (interpreter: %s)" % (pymupdf.first_line or "not importable",
                                            note, python))

    # --- Node.js + npm -------------------------------------------------------
    node = rutas.node_path()
    if node:
        probed = rutas.run([node, "--version"])
        report(probed.exit_code == 0, "Node.js",
               "%s  (%s)" % (probed.first_line or "could not run", node))
    else:
        report(False, "Node.js", "not found on PATH")

    npm = rutas.npm_path()
    if npm:
        probed = rutas.run([npm, "--version"])
        report(probed.exit_code == 0, "npm",
               "v%s  (%s)" % (probed.first_line or "could not run", npm))
    else:
        report(False, "npm", "not found on PATH")

    # --- docx library (npm) --------------------------------------------------
    # REAL check: Node is asked to resolve the module from the node dir.
    # Verifying only that node_modules/docx/package.json exists gave false
    # positives when the module was installed elsewhere, and that was exactly
    # what broke the pipeline (the generator required it from a different
    # relative path than the installer).
    node_dir = rutas.node_dir()
    package_json = node_dir / "node_modules" / "docx" / "package.json"

    if not node:
        report(False, "docx (npm)", "not verifiable: Node.js missing")
    elif not package_json.is_file():
        report(False, "docx (npm)",
               "not installed in %s (run: %s)"
               % (node_dir, rutas.comando_apa7("install")))
    else:
        version = ""
        try:
            with open(str(package_json), encoding="utf-8") as handle:
                version = str(json.load(handle).get("version", ""))
        except (OSError, ValueError):
            pass
        probe = "require('%s'); console.log('ok')" % (node_dir / "node_modules" / "docx").as_posix()
        probed = rutas.run([node, "-e", probe])
        ok = probed.exit_code == 0 and probed.first_line == "ok"
        note = ""
        if ok and version and version != rutas.DEPS["docx"]:
            note = "  [MISMATCH: pinned %s]" % rutas.DEPS["docx"]
        report(ok, "docx (npm)", "v%s%s  (%s)" % (version or "?", note, node_dir))

    # --- PDF engine: Microsoft Word first, headless LibreOffice as backup -----
    # Both engines are reported, and then ONE line says which one will actually
    # be used. "Word is installed" is not the question: what matters is whether it
    # can really export here, which only a probe can answer, so `check` probes
    # (cached afterwards) instead of trusting a file lookup.
    #
    # `auto` does not probe during --dry-run: a dry run promises to create
    # nothing and launch nothing, and the probe launches Word.
    sondear_word = not (getattr(args, "sin_sondeo", False) or rutas.dry_run())
    try:
        pedido = motores_mod.normaliza(getattr(args, "motor", "auto"))
    except ValueError as exc:
        pedido = "auto"
        report(False, "Engine (.docx to .pdf)", str(exc))

    motor_elegido, estados = motores_mod.elegir(
        pedido,
        probar_word=sondear_word,
        timeout=motores_mod.word_motor.TIMEOUT_SONDEO,
        reevaluar=bool(getattr(args, "recheck_motor", False)),
    )

    for estado in estados:
        if estado.ok:
            inform("Microsoft Word" if estado.nombre == motores_mod.WORD else "LibreOffice",
                   estado.detalle)
        elif estado.nombre == motores_mod.WORD:
            # INFO, not MISSING: a machine without Word is a machine with a
            # perfectly good pipeline, as long as the other engine answers.
            inform("Microsoft Word", estado.motivo)
        else:
            report(False, "LibreOffice", estado.motivo)

    if motor_elegido is not None:
        # The reason belongs to the engine that was actually evaluated. Saying
        # "Microsoft Word not usable: " (with nothing after it) because the
        # user asked for LibreOffice and LibreOffice answered would be noise
        # that also implies Word was even looked at.
        nota = ""
        if motor_elegido.nombre == motores_mod.WORD:
            nota = "  (preferred)"
        elif estados and estados[0].nombre == motores_mod.WORD:
            nota = "  (Microsoft Word not usable: %s)" % estados[0].motivo
        report(True, "Engine (.docx to .pdf)", "%s%s" % (motor_elegido.etiqueta(), nota))
    elif pedido != "auto":
        report(False, "Engine (.docx to .pdf)",
               "--motor %s is not usable here: %s. Nothing else was used on purpose: "
               "rerun without --motor to fall back to the other engine, or install this one."
               % (pedido, estados[0].motivo if estados else "unknown"))
    else:
        report(False, "Engine (.docx to .pdf)", "no usable PDF engine: %s"
               % "; ".join("%s: %s" % (e.nombre, e.motivo) for e in estados))

    # --- Package manager (only `install` needs this) -------------------------
    manager = rutas.package_manager()
    if manager:
        inform("Package manager", "%s  (%s)" % (manager.name, manager.path))
    else:
        inform("Package manager",
               "none found: automatic installation cannot run. Install by hand "
               "according to references/system-requirements.md")

    # --- Summary -------------------------------------------------------------
    sys.stdout.write("\n")
    missing = [tool for state, tool in results if state == "MISSING"]
    if not missing:
        sys.stdout.write("RESULT: OK\n")
        sys.stdout.write("ENVIRONMENT OK: all required tools are present.\n")
        return 0

    sys.stdout.write("MISSING: %s\n" % ", ".join(missing))
    sys.stdout.write("RESULT: MISSING\n")
    sys.stdout.write("Run: %s\n" % rutas.comando_apa7("install"))
    sys.stdout.write("If that is not possible, stop the pipeline and tell the user.\n")
    return 1


# ---------------------------------------------------------------------------
# export
# ---------------------------------------------------------------------------
def _docx_a_exportar(docx, carpeta):
    """`(path, note, error)` for the .docx to export.

    An explicit `--docx` always wins. Without it the document is looked for in
    the working folder (`--carpeta-trabajo`) or, failing that, in the current
    folder: exactly one `.docx` is used, and zero or several are an ERROR
    rather than a guess, because picking one of two documents would convert the
    wrong one and report success doing it.
    """
    if docx:
        return Path(docx).expanduser(), "", None

    base = Path(carpeta).expanduser() if carpeta else Path.cwd()
    encontrados = sorted(p for p in base.glob("*.docx") if p.is_file())
    if len(encontrados) == 1:
        return (encontrados[0],
                "No --docx given; using the only .docx of the folder: %s"
                % encontrados[0].resolve(),
                None)
    if not encontrados:
        return None, "", ("No .docx in %s. Pass --docx FILE.docx."
                          % base.resolve())
    return None, "", ("%d .docx files in %s, so which one to export is ambiguous: "
                      "%s. Pass --docx FILE.docx."
                      % (len(encontrados), base.resolve(),
                         ", ".join(p.name for p in encontrados)))


def _exportar_con_word(docx, expected, args, log_dir, step):
    """Export with Microsoft Word. Returns `(code, None)`.

    `--actualizar-campos` is what makes Word recompute the TOC before exporting.
    Without it the fields Word cached at build time are exported as they are,
    which is faster and keeps the pagination stable; with it, the page numbers
    in the index are Word's own. Neither is wrong, so it is the caller's choice
    and never a silent default in one direction or the other.
    """
    actualizar = bool(getattr(args, "actualizar_campos", False))
    step("Running: Microsoft Word through COM%s"
         % (" (fields updated first)" if actualizar else ""))
    resultado = word_motor.exportar(
        docx, expected,
        timeout=getattr(args, "timeout", word_motor.TIMEOUT_EXPORTACION),
        actualizar_campos=actualizar,
        log_dir=str(log_dir),
    )

    if resultado.stdout and resultado.stdout.strip():
        step("stdout: %s" % " ".join(resultado.stdout.split()))
    real = rutas.filter_stderr(resultado.stderr)
    if resultado.stderr:
        if not real:
            step("stderr: only benign Word driver noise (ignored)")
        else:
            step("stderr with %d real line(s):" % len(real), "WARN")
            for line in real:
                step("  " + line, "WARN")

    if resultado.timed_out:
        step("Microsoft Word did not finish within %d s and was stopped (exit code 124)."
             % getattr(args, "timeout", word_motor.TIMEOUT_EXPORTACION), "FAIL")
        step("The conversion was aborted, not failed silently. Raise --timeout and retry.",
             "FAIL")
        return 1, None

    if resultado.pdf_creado and expected.is_file():
        return 0, None

    step("Word did not produce the PDF: %s" % (resultado.motivo or "unknown error"), "FAIL")
    return 1, None


def _exportar_con_libreoffice(docx, expected, args, log_dir, step):
    """Export with headless LibreOffice. Returns `(code, soffice_result)`.

    The SofficeResult is handed back so the caller can delete the isolated
    profile the run created, which is the only part of it worth deleting.
    """
    if getattr(args, "cerrar_libreoffice", False):
        ajenos = rutas.kill_soffice_processes()
        step("--cerrar-libreoffice: closed %d LibreOffice process(es). Any "
             "open document with unsaved changes has been lost." % ajenos, "WARN")
    else:
        # Only our own runs (isolated `lo_profile`), never the user's.
        leftovers = rutas.kill_soffice_processes(rutas.SOFFICE_PROFILE_NAME)
        if leftovers:
            step("Cleaned up %d leftover LibreOffice process(es) from an interrupted run"
                 % leftovers)

    if expected.exists():
        expected.unlink()

    arguments = ["--convert-to", "pdf:writer_pdf_Export", "--outdir", str(expected.parent),
                 str(docx)]
    step("Running: soffice %s" % " ".join(arguments))
    try:
        result = rutas.run_soffice(arguments, timeout=args.timeout, log_dir=str(log_dir))
    except rutas.SofficeNotFound as exc:
        step(str(exc), "FAIL")
        return 1, None

    step("soffice finished with code %d" % result.exit_code)
    if result.stdout and result.stdout.strip():
        step("stdout: %s" % " ".join(result.stdout.split()))
    real = rutas.filter_stderr(result.stderr)
    if result.stderr:
        if not real:
            step("stderr: only benign LibreOffice noise (ignored)")
        else:
            step("stderr with %d real line(s):" % len(real), "WARN")
            for line in real:
                step("  " + line, "WARN")

    if result.timed_out:
        step("LibreOffice did not finish within %d s and was killed (exit code 124)."
             % args.timeout, "FAIL")
        step("The conversion was aborted, not failed silently. Raise --timeout, or "
             "check whether a previous soffice process is stuck, and retry.", "FAIL")
        killed = rutas.kill_soffice_processes(rutas.SOFFICE_PROFILE_NAME)
        if killed:
            step("Killed %d leftover LibreOffice process(es) from the aborted run." % killed,
                 "WARN")
        return 1, result

    if not expected.is_file():
        step("The expected PDF was not produced: %s" % expected, "FAIL")
        return 1, result

    return 0, result


def cmd_export(args):
    """.docx -> .pdf with Microsoft Word or with headless LibreOffice.

    Progress goes to stderr and the resulting PDF path goes to stdout, so the
    caller can use it without having to parse a log. Both are also written to
    the log file.

    `--motor auto|word|libreoffice` picks the engine (see lib/motores.py). Exit
    code 2 means "no usable engine": the environment is missing something, not
    the document, and SKILL.md answers that one by asking the user before
    installing. A named engine that is not usable also exits 2, and never falls
    back to the other one.

    `--docx` may be omitted: the only `.docx` of the working folder (or of the
    current folder) is then exported, and zero or several are reported as
    ambiguous instead of being guessed.

    Four details that are not optional:

    * Word is refused while the USER has Word open. Word is a single-instance COM
      server, so attaching to their session would mean quitting it afterwards,
      closing whatever they had open. The engine therefore never touches a Word
      it did not start.
    * An ISOLATED LibreOffice profile is used on every run (inside run_soffice).
      With the real profile, an already-open LibreOffice makes the conversion
      hang silently. Because of that isolation, an open LibreOffice of the user
      is NOT a problem and is left alone.
    * Only the skill's OWN leftovers are cleaned (the isolated `lo_profile`
      runs). A soffice.bin from a previous interrupted run makes the next
      --convert-to fail or hang, with no message explaining why, but killing the
      user's LibreOffice (with unsaved documents) is not the fix. The explicit
      `--cerrar-libreoffice` does close everything, and says so.
    * A timeout is reported BEFORE asking whether the PDF exists. A wedged engine
      would otherwise be diagnosed as a missing output file, which is a different
      problem with a different fix.
    """
    lines = []

    def step(message, level="INFO"):
        rutas.log("%s %s" % (level, message))
        lines.append("[%s] %s %s" % (datetime.now().strftime("%H:%M:%S"), level, message))

    code = 1
    result = None
    log_file = None
    motor_nombre = None

    try:
        docx, nota, error = _docx_a_exportar(
            args.docx, getattr(args, "carpeta_trabajo", None))
        if error:
            step(error, "FAIL")
            return 1
        if nota:
            step(nota, "INFO")
        if not docx.is_file():
            step("The .docx does not exist: %s" % docx, "FAIL")
            return 1
        docx = docx.resolve()

        out_dir = Path(args.outdir).expanduser() if args.outdir else docx.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        out_dir = out_dir.resolve()

        # An explicit --outdir has always meant "PDF and logs go here", and that
        # is what every existing caller expects, so it keeps its `_logs/`.
        # Without one, the logs belong to the document's own working folder, and
        # a `_logs/` next to the user's document is exactly the clutter this
        # change exists to remove.
        carpeta = getattr(args, "carpeta_trabajo", None)
        if args.outdir and not carpeta:
            log_dir = out_dir / "_logs"
        else:
            log_dir = rutas.rutas_documento(salida=docx, carpeta_trabajo=carpeta).logs
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = Path(args.log).expanduser() if args.log else log_dir / "03-export.log"

        expected = out_dir / (docx.stem + ".pdf")

        # --- The engine --------------------------------------------------------
        pedido = getattr(args, "motor", None) or "auto"
        try:
            pedido = motores_mod.normaliza(pedido)
        except ValueError as exc:
            step(str(exc), "FAIL")
            return 2

        motor, estados = motores_mod.elegir(
            pedido, timeout=word_motor.TIMEOUT_SONDEO)
        for estado in estados:
            if estado.ok:
                step("%s: %s" % (estado.nombre, estado.detalle))
            else:
                step("%s not usable: %s" % (estado.nombre, estado.motivo), "WARN")

        if motor is None:
            if pedido != "auto":
                step("The engine you asked for (--motor %s) cannot convert here: %s"
                     % (pedido, estados[0].motivo if estados else "unknown"), "FAIL")
                step("Nothing else was used on purpose. Rerun without --motor to fall "
                     "back automatically, or install this engine: %s"
                     % rutas.comando_apa7("install", "--only", pedido), "FAIL")
            else:
                step("No usable PDF engine on this machine:", "FAIL")
                for estado in estados:
                    step("  %s: %s" % (estado.nombre, estado.motivo), "FAIL")
                step("Run: %s" % rutas.comando_apa7("install"), "FAIL")
                step("Word is installed by hand; `install` only sets up LibreOffice. "
                     "Ask the user which one to use.", "FAIL")
            return 2

        step("=== PHASE 3: export to PDF with %s ===" % motor.etiqueta())
        step("Skill root      : %s" % rutas.skill_root())
        step("Workdir         : %s" % rutas.workdir())
        step("Engine          : %s" % motor.nombre)
        step("Input (.docx)   : %s" % docx)
        step("Expected output : %s" % expected)

        started = time.time()
        motor_nombre = motor.nombre
        if motor.nombre == motores_mod.WORD:
            code, _ = _exportar_con_word(docx, expected, args, log_dir, step)
        else:
            code, result = _exportar_con_libreoffice(docx, expected, args, log_dir, step)
        elapsed = time.time() - started
        step("Engine finished with code %d in %.1f s" % (code, elapsed))
        if code == 0:
            step("PDF generated: %s (%d bytes)" % (expected, expected.stat().st_size))
            # The only thing on stdout: the caller gets the artifact, not a log.
            sys.stdout.write("%s\n" % expected)
        return code

    except KeyboardInterrupt:
        step("Interrupted.", "WARN")
        return 130
    except Exception as exc:  # noqa: BLE001 - a broken export must not traceback
        step("Unexpected error: %s" % exc, "FAIL")
        return 1
    finally:
        # Only our own isolated-profile run: the whole tree, not just the
        # launcher, because a surviving soffice.bin makes the NEXT --convert-to
        # fail. The user's LibreOffice is not ours to close (see
        # --cerrar-libreoffice).
        #
        # Only when LibreOffice was the engine that ran. On the Word path there
        # is nothing of ours to reap, and querying soffice processes there costs
        # a shell-out per export; it would also kill an in-flight conversion of
        # a *concurrent* apa7 run, which shares the profile name. Leftovers from
        # an interrupted run are reaped by _exportar_con_libreoffice, before the
        # next conversion, which is where they actually do damage.
        if motor_nombre == motores_mod.LIBREOFFICE:
            rutas.kill_soffice_processes(rutas.SOFFICE_PROFILE_NAME)

        # The isolated profile is thousands of files. It is only useful during
        # the conversion, so it is always removed; the .log files are kept
        # because they are the useful diagnostics.
        # getattr, not isinstance: what is needed is only the path of the profile
        # the run created, and the Word path has no profile at all.
        if getattr(result, "profile_path", None):
            profile = Path(result.profile_path)
            if profile.is_dir():
                if rutas.remove_tree(profile):
                    step("Temporary LibreOffice profile deleted: %s" % profile)
                else:
                    step("Could not delete the temporary profile %s" % profile, "WARN")

        if log_file:
            try:
                log_file.write_text("\n".join(lines) + "\n", encoding="utf-8")
                rutas.log("Log written: %s" % log_file)
            except OSError as exc:
                rutas.log("Could not write the log %s: %s" % (log_file, exc))


# ---------------------------------------------------------------------------
# install
# ---------------------------------------------------------------------------
def _confirm(description, assume_yes, dry_run):
    """The single place where an action is allowed to happen.

    Every installing step goes through here, so `--dry-run` cannot leak a real
    installation: there is exactly one gate to get right instead of one per
    step. It also never blocks. An agent or a redirected stdin has nobody to
    answer a password or a y/n prompt, so a non-interactive run without --yes
    refuses instead of hanging until someone kills it.
    """
    if dry_run:
        rutas.log("[dry-run] would: %s" % description)
        return False

    if assume_yes:
        rutas.log("Confirmed by --yes: %s" % description)
        return True

    if not sys.stdin.isatty():
        sys.stderr.write(
            "ERROR: %s needs confirmation, but stdin is not a terminal.\n"
            "Nothing was installed. Re-run with --yes to confirm, or with "
            "--dry-run to see the plan.\n" % description)
        return False

    sys.stderr.write("%s\nProceed? [y/N] " % description)
    sys.stderr.flush()
    answer = sys.stdin.readline().strip().lower()
    return answer in ("y", "yes")


def _run_install(argv, cwd=None, timeout=1800):
    """Run an install command. Output goes to stderr, never to the contract."""
    argv = [str(argument) for argument in argv]
    rutas.log("Running: %s" % " ".join(argv))
    try:
        completed = subprocess.run(
            argv,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            errors="replace",
            # DEVNULL, never inherit: apt and brew both prompt, and an
            # interactive prompt in a redirected run blocks until timeout.
            stdin=subprocess.DEVNULL,
            cwd=str(cwd) if cwd else None,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        # The process tree is left to the platform; these installers are
        # idempotent enough to re-run, and a traceback here would hide which
        # step hung.
        sys.stderr.write("ERROR: the command above did not finish within %s "
                         "seconds and was stopped. Re-run it, or install the "
                         "component by hand.\n" % timeout)
        return 124
    except OSError as exc:
        # The usual cause is a package manager that disappeared between the
        # check and now, or a wrong override in an env var.
        sys.stderr.write("ERROR: could not run %s (%s).\n" % (argv[0], exc))
        return 127

    if completed.returncode != 0:
        for line in (completed.stdout or "").splitlines():
            if line.strip():
                sys.stderr.write("    %s\n" % line)
    return completed.returncode


def cmd_install(args):
    """Run the installer, telling the rest of the code whether it is simulating.

    Kept as a thin wrapper so the dry-run flag is set for the WHOLE body, and
    always cleared afterwards (tests reuse the process). Without it, workdir()
    would create the working directory even for `install --dry-run`.
    """
    rutas.set_dry_run(getattr(args, "dry_run", False))
    try:
        return _cmd_install(args)
    finally:
        rutas.set_dry_run(False)


def _cmd_install(args):
    """Install whatever `check` reported as missing.

    Idempotent: a tool that is already present is not touched. At the end it
    runs the preflight again and returns its exit code, so the caller gets one
    consistent answer about the environment.

    Progress goes to stderr; the final `check` output goes to stdout. That is
    the only thing this command writes to stdout, so `install | check`-style
    parsing keeps working.
    """
    only = set(tool.strip().lower() for tool in (args.only or []))
    unknown = only - set(instalador.TOOLS)
    if unknown:
        sys.stderr.write("ERROR: unknown --only value(s): %s\n"
                         % ", ".join(sorted(unknown)))
        sys.stderr.write("Valid tools: %s\n" % ", ".join(instalador.TOOLS))
        return 2

    def wants(tool):
        return not only or tool in only

    def step(message):
        rutas.log(message)

    step("--- Initial environment check ---")
    if cmd_check(None) == 0:
        step("Nothing to install: environment complete.")
        return 0

    manager = rutas.package_manager()
    if manager is None:
        # `check` reports this as INFO, because the environment can be complete
        # without a manager. `install` is the one command that cannot work
        # without one, so here it is fatal and the manual steps are printed.
        sys.stderr.write("\nERROR: no supported package manager was found "
                         "(winget, brew, apt, dnf, pacman).\n")
        sys.stderr.write("Automatic installation cannot run. Install by hand "
                         "whatever `check` reported as MISSING:\n")
        for tool in instalador.TOOLS:
            if wants(tool):
                for line in instalador.manual_instructions(
                        tool, rutas.is_windows(), rutas.is_macos()):
                    sys.stderr.write("  %s\n" % line)
        sys.stderr.write("Then run: %s\n" % rutas.comando_apa7("check"))
        return 1

    step("Package manager available: %s (%s)" % (manager.name, manager.path))
    step(instalador.python_version_note())

    # --- 1. Node.js + npm ---------------------------------------------------
    node = rutas.node_path()
    if node and wants("node"):
        step("Node.js present: %s" % node)
    elif not node and not wants("node"):
        step("Node.js is missing and --only excludes it: skipping.")
    elif not node:
        argv = instalador.build(manager, "node")
        if _confirm("Install Node.js LTS", args.yes, args.dry_run):
            step("Installing Node.js LTS...")
            if _run_install(argv) != 0:
                sys.stderr.write("FAILED to install Node.js. Try by hand:\n  %s\n"
                                 % instalador.manual_instructions(
                                     "node", rutas.is_windows(), rutas.is_macos())[0])
                return 1
        node = rutas.node_path()

    npm = rutas.npm_path()
    if not npm:
        # Only fatal if this run was supposed to end up with Node in it. With
        # --only libreoffice there is nothing to be done about npm, and bailing
        # out here would silently skip the step the user actually asked for.
        # During a dry run npm is missing because nothing was installed, which
        # is the point of a dry run, not a failure.
        if args.dry_run and (wants("node") or wants("docx")):
            step("[dry-run] npm would be used once Node.js is installed.")
        elif wants("node") or wants("docx"):
            sys.stderr.write("FAILED: npm is not available after installing "
                             "Node.js.\n")
            return 1
        else:
            step("npm is unavailable and --only excludes Node/docx: skipping that step.")

    # --- 2. docx library -----------------------------------------------------
    module_dir = rutas.node_dir() / "node_modules" / "docx"
    if module_dir.is_dir():
        step("docx (npm) present.")
    elif not wants("docx"):
        step("docx is missing and --only excludes it: skipping.")
    elif not npm:
        if args.dry_run:
            step("[dry-run] would install docx@%s once npm is available."
                 % rutas.DEPS["docx"])
        else:
            step("docx cannot be installed without npm: skipping.")
    else:
        # The directory and the npm anchor are only created once the install is
        # actually going ahead. They used to be written before _confirm, so a
        # --dry-run left an empty node_dir and a package.json behind: the one
        # gate that is supposed to make dry runs safe was leaking writes.
        node_dir = rutas.node_dir()
        description = "npm install docx@%s in %s" % (rutas.DEPS["docx"], node_dir)
        if _confirm(description, args.yes, args.dry_run):
            node_dir.mkdir(parents=True, exist_ok=True)
            anchor = node_dir / "package.json"
            if not anchor.exists():
                anchor.write_text(instalador.npm_anchor_json(), encoding="utf-8")
                step("Created npm anchor: %s" % anchor)
            step("Installing docx@%s..." % rutas.DEPS["docx"])
            if _run_install(instalador.npm_install_docx(
                    npm, node_dir, rutas.DEPS["docx"]), cwd=node_dir) != 0:
                sys.stderr.write("FAILED to install docx (npm).\n")
                return 1
            if not module_dir.is_dir():
                sys.stderr.write("FAILED: docx is still not in %s\n" % module_dir)
                return 1
            step("docx (npm) installed.")
        elif not args.dry_run:
            step("docx is missing and the run was not confirmed: skipped.")

    # --- 3. Python interpreter ----------------------------------------------
    # The venv below needs a real base interpreter, so this comes first. It is
    # after Node on purpose: the Windows Python installer adds shims that need
    # to be on PATH, and doing it after npm keeps them visible.
    base = None
    existing_venv = rutas.venv_python()
    if existing_venv and rutas.python_works(existing_venv):
        base = existing_venv
        step("Virtual environment usable: %s" % base)

    if base is None:
        override = os.environ.get("APA7_PYTHON")
        if override and rutas.python_works(override):
            base = override
    if base is None:
        base = rutas.python_path()

    if not base:
        if not wants("python"):
            # Not fatal: the remaining wanted steps (LibreOffice) do not need
            # Python. The final check is what reports the environment as still
            # incomplete, so the exit code stays truthful either way.
            sys.stderr.write("WARNING: no working Python interpreter and --only "
                             "excludes installing one.\n")
            step("Skipping Python, the virtual environment and pymupdf.")
        elif _confirm("Install Python 3.12", args.yes, args.dry_run):
            step("Installing Python 3.12...")
            if _run_install(instalador.build(manager, "python")) != 0:
                sys.stderr.write("FAILED to install Python. Try by hand:\n  %s\n"
                                 % instalador.manual_instructions(
                                     "python", rutas.is_windows(), rutas.is_macos())[0])
                return 1
            base = rutas.python_path()
        elif not args.dry_run:
            step("No Python interpreter and the run was not confirmed: skipping "
                 "the virtual environment and pymupdf.")
    if base:
        step("Python interpreter available: %s" % base)

    # --- 4. venv + pinned pymupdf -------------------------------------------
    # pymupdf used to be installed with a bare `pip install pymupdf` against
    # whatever interpreter was found: unpinned, and against the SYSTEM
    # interpreter, which on a PEP 668 "externally managed" Python fails with
    # externally-managed-environment. Installing into the skill's own venv fixes
    # both: the version is pinned and pip never touches the system.
    venv_dir = rutas.skill_root() / ".venv"
    venv_exe = _venv_executable(venv_dir)

    # A venv records the absolute path of its base interpreter (pyvenv.cfg), so
    # if that interpreter moved or was upgraded the venv exists but cannot run.
    # It cannot be repaired, only rebuilt.
    venv_broken = venv_exe.is_file() and not rutas.python_works(str(venv_exe))
    if venv_broken:
        if args.dry_run:
            step("[dry-run] would rebuild the virtual environment at %s (its "
                 "base interpreter moved or was upgraded)." % venv_dir)
        else:
            step("The virtual environment does not run (its base interpreter "
                 "moved or was upgraded). Recreating it...")
            rutas.remove_tree(venv_dir)

    if venv_broken and args.dry_run:
        # It would have been rebuilt; do not report the broken one as present,
        # and do not try to install into it.
        pass
    elif venv_exe.is_file():
        step("Virtual environment present: %s" % venv_exe)
    elif not base:
        # Nothing to build a venv with; step 3 already said so.
        pass
    elif not wants("pymupdf"):
        # Not fatal: --only libreoffice still has LibreOffice to install, and the
        # final check reports pymupdf as MISSING on its own.
        sys.stderr.write("WARNING: no virtual environment at %s and --only "
                         "excludes creating it; `verify` will not work.\n"
                         % venv_dir)
    elif _confirm("Create the virtual environment at %s" % venv_dir,
                  args.yes, args.dry_run):
        step("Creating the virtual environment: %s" % venv_dir)
        if _run_install([base, "-m", "venv", str(venv_dir)]) != 0 or not venv_exe.is_file():
            sys.stderr.write("FAILED to create the virtual environment.\n")
            return 1
        step("Virtual environment created.")
    elif not args.dry_run:
        step("No virtual environment and the run was not confirmed: skipped.")

    # From here on the venv interpreter is the one that matters, whatever the
    # override or the discovery said.
    if venv_exe.is_file():
        pin = rutas.DEPS["pymupdf"]
        probed = rutas.run([str(venv_exe), "-c", "import pymupdf; print(pymupdf.__version__)"])
        if probed.exit_code != 0 or probed.first_line != pin:
            description = "pip install pymupdf==%s into %s" % (pin, venv_dir)
            if _confirm(description, args.yes, args.dry_run):
                step("Installing pymupdf==%s..." % pin)
                if _run_install([str(venv_exe), "-m", "pip", "install",
                                 "--disable-pip-version-check", "pymupdf==%s" % pin]) != 0:
                    sys.stderr.write("FAILED to install pymupdf==%s.\n" % pin)
                    return 1
                step("pymupdf installed.")
            elif not args.dry_run:
                step("pymupdf missing or wrong version and the run was not "
                     "confirmed: skipping.")
        else:
            step("pymupdf present: %s" % probed.first_line)

    # --- 5. LibreOffice (only if nothing else can make a PDF) -----------------
    # Word is not installable from here (it is a licensed desktop product the
    # user installs by hand), so when Word is proven to work, the engine exists
    # and installing a second, much larger one is work nobody asked for.
    #
    # `--only libreoffice` still installs it: asking for it by name is exactly
    # the case where the user wants it whatever else answers.
    if rutas.soffice_path():
        step("LibreOffice present.")
    elif not wants("libreoffice"):
        step("LibreOffice is missing and --only excludes it: skipping.")
    else:
        # --only libreoffice is an explicit order, and it outranks the shortcut
        # below: if the answer to an explicit request were "you do not need it",
        # the flag would be unusable.
        pedido_explicito = "libreoffice" in (getattr(args, "only", None) or [])
        estado_word = motores_mod.estado_word(probar=not args.dry_run)
        if estado_word.ok and not pedido_explicito:
            step("LibreOffice missing, but %s already converts .docx to .pdf: skipping "
                 "the LibreOffice installation. Use --only libreoffice to install it anyway."
                 % estado_word.etiqueta())
        else:
            if estado_word.sondeado and estado_word.motivo:
                step("Microsoft Word is installed but not usable (%s), so LibreOffice "
                     "is the engine that will be used." % estado_word.motivo)
            if _confirm("Install LibreOffice", args.yes, args.dry_run):
                step("Installing LibreOffice (this can take several minutes)...")
                if _run_install(instalador.build(manager, "libreoffice"),
                                timeout=3600) != 0 or not rutas.soffice_path():
                    sys.stderr.write("FAILED: LibreOffice not found after installation.\n")
                    sys.stderr.write("Paths examined: %s\n"
                                     % " | ".join(p for p in rutas.soffice_candidates() if p))
                    sys.stderr.write("If it is installed elsewhere, set APA7_SOFFICE "
                                     "to the full path.\n")
                    return 1
                step("LibreOffice installed: %s" % rutas.soffice_path())

    # --- 6. Final check ------------------------------------------------------
    step("--- Final environment check ---")
    return cmd_check(None)


def _venv_executable(venv_dir):
    """The interpreter inside a venv, in the layout of THIS platform."""
    if rutas.is_windows():
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python3"


# ---------------------------------------------------------------------------
# parse / build / verify
#
# Thin dispatchers. Each one knows only three things the caller would otherwise
# have to know per platform: which interpreter runs the script, where the script
# is, and how the workdir is wired. Every argument after the subcommand is
# forwarded verbatim, so each underlying script stays the single source of truth
# for its own options and --help still works.
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# The subcommands that hand their arguments to another script instead of
# parsing them. A dict, not a tuple: main() looks the handler up by name before
# argparse runs, so the arguments reach the target script untouched.
# ---------------------------------------------------------------------------
def _forwarded(args):
    """Every argument after the subcommand, in the order the user wrote it.

    parse_known_args is required, not a shortcut: the dispatcher cannot know the
    target script's options, so anything it does not recognise has to be
    forwarded rather than rejected. That is also what makes `--help` reach the
    real script.
    """
    return [str(argument) for argument in (list(getattr(args, "rest", []))
                                          + list(getattr(args, "unknown", [])))]


def _forward(rest, interpreter, script, label, usage):
    if not rest:
        sys.stderr.write("%s: no arguments.\nUsage: %s %s\n"
                         % (label, rutas.comando_apa7(), usage))
        return 2
    if not Path(script).is_file():
        sys.stderr.write("ERROR: %s not found at %s\n" % (label, script))
        return 1
    return subprocess.call([str(interpreter), str(script)] + list(rest))


def cmd_parse(args):
    """Markdown -> MANIFEST.json.

    Needs no third-party package, so the current interpreter is enough.

    `--out` is MANDATORY in md-a-manifiesto.py, so the default is computed here
    and passed explicitly instead of relaxing that script: its own contract is
    left alone, and a caller who gave `--out` still gets exactly what they asked
    for.
    """
    rest = _forwarded(args)
    carpeta, rest = _extrae(rest, "--carpeta-trabajo")
    md = _valor_de(rest, "--md")
    # crear=True: the documented layout is datos/ AND logs/, and it should be
    # there from the first run instead of appearing later, one folder at a time.
    documento = _rutas_de({"md": md}, carpeta, crear=True)
    if documento is None:
        return _forward(rest, sys.executable,
                        rutas.skill_script("md-a-manifiesto.py"), "parse",
                        "parse --md FILE.md [--out MANIFEST.json] [options]")

    if md:
        aviso = rutas.anota_fuente(documento.trabajo, md)
        if aviso:
            sys.stderr.write(aviso + "\n")

    rest = _inyecta(rest, "--out", documento.manifiesto)
    rest = _inyecta(rest, "--log", documento.log_parse)
    # Only when it is really there: md-a-manifiesto.py ignores a portada that
    # does not exist, and naming a file that is not there helps nobody.
    if documento.portada.is_file():
        rest = _inyecta(rest, "--portada", documento.portada)

    return _forward(rest, sys.executable,
                    rutas.skill_script("md-a-manifiesto.py"), "parse",
                    "parse --md FILE.md [--out MANIFEST.json] [options]")


def _valor_de(rest, nombre):
    """Value of `--nombre` in a forwarded argument list, or None."""
    for i, argumento in enumerate(rest):
        if argumento == nombre and i + 1 < len(rest):
            return rest[i + 1]
        if argumento.startswith(nombre + "="):
            return argumento.split("=", 1)[1]
    return None


def _extrae(rest, nombre):
    """(value of `--nombre`, the rest of the list). The option is REMOVED.

    Needed for the options the dispatcher owns and the target scripts do not
    know: forwarded verbatim, `md-a-manifiesto.py` would fail on an argument it
    has never heard of instead of reading the document.
    """
    valor = None
    resto = []
    elementos = list(rest)
    i = 0
    while i < len(elementos):
        argumento = elementos[i]
        if argumento == nombre and i + 1 < len(elementos):
            valor = elementos[i + 1]
            i += 2
            continue
        if argumento.startswith(nombre + "="):
            valor = argumento.split("=", 1)[1]
            i += 1
            continue
        resto.append(argumento)
        i += 1
    return valor, resto


def _inyecta(rest, nombre, valor):
    """Append `--nombre valor` unless the caller already passed the option.

    Precedence, everywhere in this dispatcher: what the user wrote wins. The
    default is only ever what the user did NOT ask for.
    """
    if valor is None or _valor_de(rest, nombre) is not None:
        return list(rest)
    return list(rest) + [nombre, str(valor)]


def _rutas_de(anclas, carpeta=None, **opciones):
    """rutas_documento() for a subcommand, or None when nothing is given.

    None means "let the target script complain": `parse` with no --md has to stay
    the target script's clear error about the missing .md, not a ValueError with
    a traceback coming out of the dispatcher.
    """
    if not carpeta and not any(anclas.values()):
        return None
    return rutas.rutas_documento(carpeta_trabajo=carpeta, **opciones, **anclas)


def _manifiesto_con_toc(rest):
    """The manifest if it asks for real TOC fields, else None.

    Only `toc_campos: false` turns them off; any other value (including the
    option being absent) keeps the default, which is to build them.
    """
    ruta = _valor_de(rest, "--manifiesto")
    if not ruta or not Path(ruta).is_file():
        return None
    try:
        with open(ruta, encoding="utf-8") as fh:
            manifiesto = json.load(fh)
    except (OSError, ValueError):
        return None
    if (manifiesto.get("opciones") or {}).get("toc_campos") is False:
        return None
    return manifiesto


def _segunda_pasada_indices(node, build_js, rest, documento, motor=None):
    """Fill the indexes' page numbers by exporting once and rebuilding.

    A TOC field's result is computed by the rendering engine, never by
    build-docx.js, so the page numbers are only known after a real export. The
    sequence is: export the first .docx to a throwaway PDF, read the page of
    every entry with paginas-de-pdf.py, and rebuild the .docx passing that map
    with --paginas-json. Its numbers become the fields' cached result.

    `motor` is the ALREADY RESOLVED engine name, and it is resolved by the
    caller and passed in rather than chosen here. That is the whole point: the
    throwaway export and the final export must paginate identically, and two
    engines paginate differently. Letting each of them decide could measure the
    pages with LibreOffice and then deliver a Word-rendered PDF whose index does
    not match it.

    Everything temporary goes into the document's logs/ folder, never into the
    folder the user sees: the throwaway PDF and paginas.json used to be written
    to `<docx folder>/_logs/`, which meant a `_logs/` directory beside the
    document the user was trying to tidy up. paginas.json is KEPT, because it is
    the evidence that explains an index with wrong pages; the throwaway PDF is
    deleted once the run succeeds, and kept when it does not, since a PDF that
    could not be converted is exactly what one needs to look at.

    Any missing piece (no venv, no engine, an unreadable PDF) leaves the
    first .docx as the final one: the indexes then ship without numbers and Word
    fills them in when the document is opened. It is a degradation, not a
    failure, so the exit code of the first build is kept.
    """
    docx = _valor_de(rest, "--out")
    ruta_manifiesto = _valor_de(rest, "--manifiesto")
    if not docx or not ruta_manifiesto or not Path(docx).is_file():
        return 0

    interpreter = rutas.venv_python()
    if not (interpreter and rutas.python_works(interpreter)):
        return 0

    logs = documento.logs
    logs.mkdir(parents=True, exist_ok=True)
    mapa = documento.paginas_json

    # Pass 1: export with the fields unresolved. cmd_export is reused so there is
    # one conversion code path, not two. Its stdout (the PDF path) is swallowed:
    # build's own contract must not gain stray lines.
    with contextlib.redirect_stdout(io.StringIO()):
        codigo = cmd_export(argparse.Namespace(
            docx=str(docx), outdir=str(logs),
            log=str(documento.log_paginas), timeout=300,
            carpeta_trabajo=str(documento.trabajo),
            motor=motor or "auto"))
    if codigo != 0:
        return 0
    pdf = documento.pdf_auxiliar
    if not pdf.is_file():
        return 0

    generador = rutas.skill_script("paginas-de-pdf.py")
    if not Path(generador).is_file():
        return 0
    codigo = subprocess.call([str(interpreter), str(generador),
                              "--pdf", str(pdf),
                              "--manifiesto", str(ruta_manifiesto),
                              "--out", str(mapa)])
    if codigo != 0 or not mapa.is_file():
        return 0

    # Pass 2: rebuild with the measured pages cached inside the TOC fields.
    codigo = subprocess.call([str(node), str(build_js)] + list(rest)
                             + ["--paginas-json", str(mapa)])

    if codigo == 0:
        try:
            pdf.unlink()
        except OSError:
            pass
    return codigo


def cmd_build(args):
    """MANIFEST.json -> .docx.

    Needs Node.js. The script resolves the docx package from the workdir by
    absolute path, so the working directory does not matter here.

    When the manifest asks for real TOC fields, this runs build-docx.js twice
    (see _segunda_pasada_indices): the page numbers of the indexes only exist
    after a real export, which needs the .docx from the first run.

    `--motor` is accepted here (and removed before forwarding) for one reason:
    both passes must use the SAME engine, and the engine is decided once, here.
    """
    node = rutas.node_path()
    if not node:
        sys.stderr.write("ERROR: Node.js not found.\nRun: %s\n"
                         % rutas.comando_apa7("install", "--only", "node"))
        return 1
    rest = _forwarded(args)
    carpeta, rest = _extrae(rest, "--carpeta-trabajo")
    motor, rest = _extrae(rest, "--motor")
    manifiesto = _valor_de(rest, "--manifiesto")
    destino = _valor_de(rest, "--out")

    # The NAME is checked here, before anything is built, and the engine is only
    # resolved further down. The two are split on purpose: a typo must fail
    # immediately instead of being silently ignored (build-docx.js does not know
    # this option, and `export` is a different command that would only fail much
    # later), while resolving means probing Microsoft Word for real, which is
    # seconds of work that a build without TOC fields never needs.
    try:
        pedido = motores_mod.normaliza(motor)
    except ValueError as exc:
        sys.stderr.write("ERROR: %s\n" % exc)
        return 2

    documento = _rutas_de({"salida": destino, "manifiesto": manifiesto}, carpeta)
    if documento is None:
        return _forward(rest, node, rutas.skill_script("build-docx.js"), "build",
                        "build --manifiesto MANIFEST.json [--out salida.docx] [--log log.txt]")

    # Two passes over rutas_documento, and both are needed. The first only knows
    # the manifest, which is enough to find the working folder but not to name
    # the deliverable; the second runs with the .docx path settled so the name of
    # the document is the one the user gave, never "MANIFEST".
    if _valor_de(rest, "--out") is None:
        rest = _inyecta(rest, "--out", documento.docx)
    documento = rutas.rutas_documento(salida=destino or documento.docx,
                                      manifiesto=manifiesto,
                                      carpeta_trabajo=carpeta)
    rest = _inyecta(rest, "--log", documento.log_build)

    build_js = rutas.skill_script("build-docx.js")
    codigo = _forward(rest, node, build_js, "build",
                      "build --manifiesto MANIFEST.json [--out salida.docx] [--log log.txt]")
    if codigo != 0 or _manifiesto_con_toc(rest) is None:
        return codigo

    # Resolve the engine ONCE, before the measuring export, and hand the name to
    # both passes. An engine that was asked for by name and is not usable is a
    # hard error; `auto` with nothing usable only means the indexes ship without
    # numbers, which _segunda_pasada_indices already handles. The name was
    # normalised before the first pass; only the resolution is left.
    elegido, estados = motores_mod.elegir(pedido)
    if elegido is not None:
        motor = elegido.nombre
    elif pedido != "auto":
        sys.stderr.write("ERROR: --motor %s cannot convert here: %s\n"
                         % (pedido, estados[0].motivo if estados else "unknown"))
        sys.stderr.write("Nothing else was used on purpose. Rerun without --motor to "
                         "fall back automatically, or install this engine.\n")
        return 2

    return _segunda_pasada_indices(node, build_js, rest, documento, motor)


def cmd_verify(args):
    """PDF + MANIFEST.json -> verification report.

    This is the one that must NOT run on the caller's interpreter: pymupdf is
    installed into the skill's .venv, so a system Python without it either
    fails to import or, worse, imports a different version. The venv interpreter
    is selected here so `verify` behaves the same however it is invoked.

    The report goes to the document's datos/ folder by default. It used to have
    no default at all, which is why a failed verification often left nothing
    behind to explain it.
    """
    interpreter = rutas.venv_python()
    if not (interpreter and rutas.python_works(interpreter)):
        sys.stderr.write(
            "WARNING: no usable virtual environment at %s; running verify with "
            "%s instead.\nIf pymupdf turns out to be missing, run: %s\n"
            % (rutas.skill_root() / ".venv", sys.executable,
               rutas.comando_apa7("install", "--only", "pymupdf")))
        interpreter = sys.executable

    rest = _forwarded(args)
    carpeta, rest = _extrae(rest, "--carpeta-trabajo")
    documento = _rutas_de({"salida": _valor_de(rest, "--pdf"),
                           "manifiesto": _valor_de(rest, "--manifiesto")}, carpeta)
    if documento is not None:
        rest = _inyecta(rest, "--log", documento.log_verify)
        rest = _inyecta(rest, "--json", documento.json_verificacion)

    return _forward(rest, interpreter,
                    rutas.skill_script("verificar-pdf.py"), "verify",
                    "verify --pdf salida.pdf --manifiesto MANIFEST.json [options]")


# Filled in here rather than next to main(), because it has to name the
# handlers, and those are not defined until here.
FORWARDING = {
    "parse": cmd_parse,
    "build": cmd_build,
    "verify": cmd_verify,
}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def _run(handler, args):
    """Run a handler and turn the two expected interrupts into exit codes."""
    try:
        return handler(args)
    except KeyboardInterrupt:
        sys.stderr.write("Interrupted.\n")
        return 130
    except rutas.SofficeNotFound as exc:
        sys.stderr.write("ERROR: %s\n" % exc)
        return 1


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)

    # The forwarding subcommands take their arguments verbatim, and they are
    # sliced off here rather than through argparse. argparse.REMAINDER splits
    # them into two buckets -- the first option it recognises and the tail after
    # it -- so putting them back together REORDERS the command line and turns
    # `--md a.md --out b.json` into `a.md --out b.json --md`. Nothing warns about
    # that; the target script just complains about a missing value.
    if argv and argv[0] in FORWARDING:
        handler = FORWARDING[argv[0]]
        return _run(handler, argparse.Namespace(rest=argv[1:], unknown=[]))

    parser = argparse.ArgumentParser(
        prog="apa7.py",
        description="generate-apa-document: build an APA 7 .docx and .pdf.",
    )
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    subparsers.required = True

    check_parser = subparsers.add_parser(
        "check", help="preflight (STEP 0): verify every required tool")
    check_parser.add_argument(
        "--motor", default="auto", metavar="<engine>",
        help="which PDF engine to report on: one of %s (default: auto)"
             % ", ".join(motores_mod.PREFIJOS))
    check_parser.add_argument(
        "--sin-sondeo", dest="sin_sondeo", action="store_true",
        help="do not probe Microsoft Word, only report what is installed. Faster, "
             "but 'installed' is not the same answer as 'can export'.")
    check_parser.add_argument(
        "--recheck-motor", dest="recheck_motor", action="store_true",
        help="ignore the cached engine probe and run it again. Needed after Word is "
             "installed, moved, repaired or upgraded.")
    check_parser.set_defaults(handler=cmd_check)

    export_parser = subparsers.add_parser(
        "export", help=".docx -> .pdf with Microsoft Word or headless LibreOffice")
    export_parser.add_argument(
        "--motor", default="auto", metavar="<engine>",
        help="PDF engine: one of %s (default: auto = Word when a probe proves it "
             "works here, LibreOffice otherwise). An engine named here that is not "
             "usable is an error, never a silent fallback."
             % ", ".join(motores_mod.PREFIJOS))
    export_parser.add_argument(
        "--docx", default=None,
        help="input .docx (default: the only .docx of the working folder)")
    export_parser.add_argument("--outdir", default=None,
                               help="output folder (default: the .docx folder)")
    export_parser.add_argument("--log", default=None,
                               help="log file (default: <work folder>/logs/03-export.log, "
                                    "or <outdir>/_logs/03-export.log when --outdir is given)")
    export_parser.add_argument(
        "--carpeta-trabajo", dest="carpeta_trabajo", default=None, metavar="DIR",
        help="folder that holds this document's datos/ and logs/. Defaults to "
             "<name>_apa/ next to the .docx.")
    export_parser.add_argument("--timeout", type=int, default=300,
                               help="maximum wait in seconds (default: 300)")
    export_parser.add_argument(
        "--actualizar-campos", dest="actualizar_campos", action="store_true",
        help="with Microsoft Word: update the document fields (the table of contents) "
             "before exporting, so the index carries Word's own page numbers. Off by "
             "default: Word exports the fields as they were cached when the .docx was "
             "built. Ignored by LibreOffice, which recalculates them on its own.")
    export_parser.add_argument(
        "--cerrar-libreoffice", dest="cerrar_libreoffice", action="store_true",
        help="before converting with LibreOffice, close EVERY LibreOffice process, "
             "including the user's open documents (unsaved work is lost). Off by "
             "default: the conversion does not need it and only the skill's own "
             "leftovers are cleaned.")
    export_parser.set_defaults(handler=cmd_export)

    install_parser = subparsers.add_parser(
        "install", help="install whatever is missing")
    install_parser.add_argument(
        "--only", action="append", metavar="<tool>",
        help="install only this tool; repeatable. One of: %s"
             % ", ".join(instalador.TOOLS))
    install_parser.add_argument(
        "--yes", action="store_true",
        help="do not ask for confirmation (required when stdin is not a terminal)")
    install_parser.add_argument(
        "--dry-run", action="store_true",
        help="print what would be installed and change nothing")
    install_parser.set_defaults(handler=cmd_install)

    # Registered only so that `apa7.py --help` lists them; the branch above
    # handles them before argparse ever sees them.
    subparsers.add_parser(
        "parse", help="PHASE 1: .md -> MANIFEST.json (options are forwarded, "
                      "plus --carpeta-trabajo)",
        add_help=False)
    subparsers.add_parser(
        "build", help="PHASE 2: MANIFEST.json -> .docx (options are forwarded, "
                      "plus --carpeta-trabajo)",
        add_help=False)
    subparsers.add_parser(
        "verify", help="PHASE 4: check the PDF against the manifest (options are "
                       "forwarded, plus --carpeta-trabajo)",
        add_help=False)

    args = parser.parse_args(argv)
    return _run(args.handler, args)


if __name__ == "__main__":
    sys.exit(main())