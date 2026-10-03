#!/usr/bin/env python3
"""apa7.py - single entry point of the generate-apa-document skill.

Cross-platform replacement for the four PowerShell scripts. Standard library
only, and no syntax newer than Python 3.9, because `check` is the command that
runs BEFORE anything is installed: it is what tells the user what to install,
so it cannot depend on a third-party package or on a newer interpreter.

Subcommands
    check     preflight (STEP 0): every required tool, one line per tool
    export    .docx -> .pdf with headless LibreOffice
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

Optional environment variables
    APA7_SOFFICE, APA7_PYTHON, APA7_NODEDIR, APA7_WORKDIR, APA7_SKILL_ROOT
"""

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib import instalador  # noqa: E402
from lib import rutas  # noqa: E402

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
def cmd_check(_args):
    """Preflight (STEP 0).

    Every line goes to stdout because this output is the contract the skill
    parses; nothing human-facing is mixed in, so `RESULT:` can be read without
    having to filter noise out of the stream.
    """
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
                   "it with: python scripts/apa7.py install" % broken_venv)
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
               "not installed in %s (run: python scripts/apa7.py install)" % node_dir)
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

    # --- LibreOffice (the single PDF engine) ---------------------------------
    # The version is always queried through run_soffice, never by calling the
    # launcher directly: soffice.exe detaches, the child inherits the output pipe
    # and the caller waits forever. That hangs the whole preflight.
    soffice = rutas.soffice_path()
    if not soffice:
        searched = " | ".join(p for p in rutas.soffice_candidates() if p)
        report(False, "LibreOffice", "not found. Searched: %s"
               % (searched or "(no candidates: set APA7_SOFFICE)"))
    else:
        try:
            probed = rutas.run_soffice(["--version"], timeout=60)
        except rutas.SofficeNotFound as exc:
            report(False, "LibreOffice", str(exc))
        else:
            first = ""
            for line in (probed.stdout or "").replace("\r\n", "\n").split("\n"):
                if line.strip():
                    first = line.strip()
                    break
            ok = probed.exit_code == 0 and bool(first)
            detail = first if ok else "could not read the version (exit code %d)" % probed.exit_code
            report(ok, "LibreOffice", "%s  (%s)" % (detail, soffice))

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
    sys.stdout.write("Run: python scripts/apa7.py install\n")
    sys.stdout.write("If that is not possible, stop the pipeline and tell the user.\n")
    return 1


# ---------------------------------------------------------------------------
# export
# ---------------------------------------------------------------------------
def cmd_export(args):
    """.docx -> .pdf with headless LibreOffice (the single engine).

    Progress goes to stderr and the resulting PDF path goes to stdout, so the
    caller can use it without having to parse a log. Both are also written to
    the log file.

    Three details that are not optional:

    * An ISOLATED LibreOffice profile is used on every run (inside run_soffice).
      With the real profile, an already-open LibreOffice makes the conversion
      hang silently.
    * Leftover soffice processes are killed first. A soffice.bin that survived
      a previous run makes the next --convert-to fail or hang, with no message
      explaining why.
    * A timeout is reported BEFORE asking whether the PDF exists. A wedged
      LibreOffice would otherwise be diagnosed as a missing output file, which
      is a different problem with a different fix.
    """
    lines = []

    def step(message, level="INFO"):
        rutas.log("%s %s" % (level, message))
        lines.append("[%s] %s %s" % (datetime.now().strftime("%H:%M:%S"), level, message))

    code = 1
    result = None
    log_file = None

    try:
        docx = Path(args.docx).expanduser()
        if not docx.is_file():
            step("The .docx does not exist: %s" % docx, "FAIL")
            return 1
        docx = docx.resolve()

        out_dir = Path(args.outdir).expanduser() if args.outdir else docx.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        out_dir = out_dir.resolve()

        log_dir = out_dir / "_logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = Path(args.log).expanduser() if args.log else log_dir / "03-export.log"

        expected = out_dir / (docx.stem + ".pdf")

        step("=== PHASE 3: export to PDF with LibreOffice ===")
        step("Skill root      : %s" % rutas.skill_root())
        step("Workdir         : %s" % rutas.workdir())
        step("Input (.docx)   : %s" % docx)
        step("Expected output : %s" % expected)

        if not rutas.soffice_console():
            step("LibreOffice is not installed. Run: python scripts/apa7.py install", "FAIL")
            return 1

        leftovers = rutas.kill_soffice_processes()
        if leftovers:
            step("Cleaned up %d leftover LibreOffice process(es)" % leftovers)

        if expected.exists():
            expected.unlink()

        arguments = ["--convert-to", "pdf:writer_pdf_Export", "--outdir", str(out_dir), str(docx)]
        step("Running: soffice %s" % " ".join(arguments))

        started = time.time()
        try:
            result = rutas.run_soffice(arguments, timeout=args.timeout, log_dir=str(log_dir))
        except rutas.SofficeNotFound as exc:
            step(str(exc), "FAIL")
            return 1
        elapsed = time.time() - started

        step("soffice finished with code %d in %.1f s" % (result.exit_code, elapsed))

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
            killed = rutas.kill_soffice_processes()
            if killed:
                step("Killed %d leftover LibreOffice process(es) from the aborted run." % killed,
                     "WARN")
            return 1

        if not expected.is_file():
            step("The expected PDF was not produced: %s" % expected, "FAIL")
            return 1

        step("PDF generated: %s (%d bytes)" % (expected, expected.stat().st_size))
        # The only thing on stdout: the caller gets the artifact, not a log.
        sys.stdout.write("%s\n" % expected)
        code = 0
        return 0

    except KeyboardInterrupt:
        step("Interrupted.", "WARN")
        return 130
    except Exception as exc:  # noqa: BLE001 - a broken export must not traceback
        step("Unexpected error: %s" % exc, "FAIL")
        return 1
    finally:
        # The whole tree, not just the launcher: a soffice.bin child that
        # survives makes the NEXT --convert-to fail.
        rutas.kill_soffice_processes()

        # The isolated profile is thousands of files. It is only useful during
        # the conversion, so it is always removed; the .log files are kept
        # because they are the useful diagnostics.
        if result is not None and result.profile_path:
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
        sys.stderr.write("Then run: python scripts/apa7.py check\n")
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
        if wants("node") or wants("docx"):
            sys.stderr.write("FAILED: npm is not available after installing "
                             "Node.js.\n")
            return 1
        step("npm is unavailable and --only excludes Node/docx: skipping that step.")

    # --- 2. docx library -----------------------------------------------------
    module_dir = rutas.node_dir() / "node_modules" / "docx"
    if module_dir.is_dir():
        step("docx (npm) present.")
    elif not wants("docx"):
        step("docx is missing and --only excludes it: skipping.")
    elif not npm:
        step("docx cannot be installed without npm: skipping.")
    else:
        node_dir = rutas.node_dir()
        node_dir.mkdir(parents=True, exist_ok=True)
        anchor = node_dir / "package.json"
        if not anchor.exists():
            anchor.write_text(instalador.npm_anchor_json(), encoding="utf-8")
            step("Created npm anchor: %s" % anchor)

        description = "npm install docx@%s in %s" % (rutas.DEPS["docx"], node_dir)
        if _confirm(description, args.yes, args.dry_run):
            step("Installing docx@%s..." % rutas.DEPS["docx"])
            if _run_install(instalador.npm_install_docx(
                    npm, node_dir, rutas.DEPS["docx"]), cwd=node_dir) != 0:
                sys.stderr.write("FAILED to install docx (npm).\n")
                return 1
            if not module_dir.is_dir():
                sys.stderr.write("FAILED: docx is still not in %s\n" % module_dir)
                return 1
            step("docx (npm) installed.")
        else:
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
        else:
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

    if venv_exe.is_file() and not rutas.python_works(str(venv_exe)):
        # A venv records the absolute path of its base interpreter (pyvenv.cfg),
        # so if that interpreter moved or was upgraded the venv exists but cannot
        # run. It cannot be repaired, only rebuilt.
        step("The virtual environment exists but does not run (its base "
             "interpreter moved or was upgraded). Recreating it...")
        if not args.dry_run:
            rutas.remove_tree(venv_dir)

    if venv_exe.is_file():
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
    else:
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
            else:
                step("pymupdf missing or wrong version and the run was not "
                     "confirmed: skipping.")
        else:
            step("pymupdf present: %s" % probed.first_line)

    # --- 5. LibreOffice ------------------------------------------------------
    if rutas.soffice_path():
        step("LibreOffice present.")
    elif not wants("libreoffice"):
        step("LibreOffice is missing and --only excludes it: skipping.")
    else:
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
# Entry point
# ---------------------------------------------------------------------------
def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="apa7.py",
        description="generate-apa-document: build an APA 7 .docx and .pdf.",
    )
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    subparsers.required = True

    check_parser = subparsers.add_parser(
        "check", help="preflight (STEP 0): verify every required tool")
    check_parser.set_defaults(handler=cmd_check)

    export_parser = subparsers.add_parser(
        "export", help=".docx -> .pdf with headless LibreOffice")
    export_parser.add_argument("--docx", required=True, help="input .docx (required)")
    export_parser.add_argument("--outdir", default=None,
                               help="output folder (default: the .docx folder)")
    export_parser.add_argument("--log", default=None,
                               help="log file (default: <outdir>/_logs/03-export.log)")
    export_parser.add_argument("--timeout", type=int, default=300,
                               help="maximum wait in seconds (default: 300)")
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

    args = parser.parse_args(argv)
    try:
        return args.handler(args)
    except KeyboardInterrupt:
        sys.stderr.write("Interrupted.\n")
        return 130
    except rutas.SofficeNotFound as exc:
        sys.stderr.write("ERROR: %s\n" % exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())