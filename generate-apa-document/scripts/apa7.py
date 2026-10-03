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
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

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