"""Tests for the `check` output contract of apa7.py.

The contract is what SKILL.md parses, so it is tested as carefully as the code
behind it: which stream it goes to, the exact `STATE|tool|detail` shape, and
the fact that an INFO line can never turn a working environment into RESULT:
MISSING.

Standard library only, on purpose.
"""
import argparse
import ast
import contextlib
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import apa7  # noqa: E402
from lib import rutas  # noqa: E402


class _Line:
    """Minimal stand-in for rutas.NativeResult and rutas.SofficeResult."""

    def __init__(self, first_line="", exit_code=0, stdout="", timed_out=False):
        self.first_line = first_line
        self.exit_code = exit_code
        self.stdout = stdout
        self.stderr = ""
        self.timed_out = timed_out


def _fake_run(argv, timeout=60, cwd=None):
    """Answer the five probes cmd_check makes, by looking at the argv."""
    argv = [str(a) for a in argv]
    joined = " ".join(argv)
    if "pymupdf" in joined:
        return _Line("1.28.2")
    if "sys.version_info" in joined:
        return _Line("3.12.10 OK")
    if argv[-1:] == ["--version"]:
        # node prints a leading v, npm does not.
        return _Line("11.12.1" if "npm" in argv[0] else "v24.11.0")
    if "require(" in joined:
        return _Line("ok")
    return _Line()


class _Env:
    """A complete, healthy environment. Individual tests break one thing.

    docx is really installed in a temporary node dir, because a healthy
    environment that reports docx as MISSING is not healthy and the other tests
    would be asserting against noise.
    """

    def __init__(self, **overrides):
        self.overrides = overrides
        self._tmp = None

    def __enter__(self):
        import tempfile

        self._tmp = tempfile.TemporaryDirectory()
        workdir = Path(self._tmp.name)
        package_json = workdir / "node_modules" / "docx" / "package.json"
        package_json.parent.mkdir(parents=True)
        package_json.write_text('{"version": "9.7.1"}', encoding="utf-8")

        # Values are patched as return_value and callables as new, because
        # patch.object's third positional argument REPLACES the attribute: a
        # plain string would make `rutas.python_path()` a call on a str.
        defaults = {
            "python_path": "/usr/bin/python3",
            "venv_python": None,
            "node_path": "/usr/bin/node",
            "npm_path": "/usr/bin/npm",
            "node_dir": workdir,
            "soffice_path": "/usr/bin/soffice",
            "soffice_candidates": ["/usr/bin/soffice"],
            "package_manager": rutas.PackageManager("apt", "/usr/bin/apt-get"),
            "run": _fake_run,
            "run_soffice": lambda *a, **k: _Line(stdout="LibreOffice 26.8.0.3"),
        }
        defaults.update(self.overrides)
        patches = []
        for key, value in defaults.items():
            if callable(value):
                patches.append(mock.patch.object(rutas, key, value))
            else:
                patches.append(mock.patch.object(rutas, key, return_value=value))
        for patcher in patches:
            patcher.start()
        self._patches = patches
        return self

    def __exit__(self, *_exc):
        for patcher in reversed(self._patches):
            patcher.stop()
        if self._tmp is not None:
            self._tmp.cleanup()
        return False


class CheckHarness(unittest.TestCase):
    def run_check(self, env=None):
        """Run cmd_check and return (exit_code, stdout, stderr)."""
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            with env or _Env():
                code = apa7.cmd_check(None)
        return code, out.getvalue(), err.getvalue()


class TestContract(CheckHarness):
    def test_every_tool_line_has_the_documented_shape(self):
        _, out, _ = self.run_check()
        states = set()
        for line in out.splitlines():
            # The contract is about lines that OPEN with a state. The trailing
            # prose (RESULT:, ENVIRONMENT OK:, MISSING:, Run:) is guidance for
            # the agent and is not part of the parseable block.
            if not line.startswith(("OK|", "MISSING|", "INFO|")):
                continue
            parts = line.split("|")
            self.assertEqual(len(parts), 3, "not STATE|tool|detail: %r" % line)
            self.assertTrue(parts[1].strip(), "empty tool name: %r" % line)
            states.add(parts[0])
        self.assertEqual(states, {"OK", "INFO"})

    def test_a_healthy_environment_is_ok_with_exit_code_zero(self):
        code, out, _ = self.run_check()
        self.assertIn("RESULT: OK", out)
        self.assertEqual(code, 0)

    def test_nothing_is_written_to_stderr(self):
        # stdout is the contract; a stray line there breaks the parser, and
        # check has no progress to report anyway.
        _, _, err = self.run_check()
        self.assertEqual(err, "")

    def test_the_result_line_comes_last_and_is_unambiguous(self):
        _, out, _ = self.run_check()
        self.assertIn("RESULT: OK", out)
        self.assertNotIn("RESULT: MISSING", out)


class TestPackageManagerIsInformative(CheckHarness):
    """The regression this whole INFO state exists for."""

    def test_a_missing_package_manager_does_not_break_the_result(self):
        env = _Env(package_manager=None)
        code, out, _ = self.run_check(env)
        self.assertIn("INFO|Package manager|none found", out)
        self.assertIn("RESULT: OK", out)
        self.assertEqual(code, 0)

    def test_a_present_package_manager_is_reported_as_info_too(self):
        _, out, _ = self.run_check(_Env(package_manager=rutas.PackageManager("brew", "/opt/homebrew/bin/brew")))
        self.assertIn("INFO|Package manager|brew  (/opt/homebrew/bin/brew)", out)
        self.assertNotIn("MISSING|Package manager", out)

    def test_info_never_appears_in_the_missing_summary(self):
        code, out, _ = self.run_check(_Env(package_manager=None, node_path=None))
        summary = [l for l in out.splitlines() if l.startswith("MISSING:")][0]
        self.assertNotIn("Package manager", summary)
        self.assertEqual(code, 1)


class TestMissingTools(CheckHarness):
    def test_a_missing_tool_turns_the_result_into_missing(self):
        code, out, _ = self.run_check(_Env(node_path=None))
        self.assertIn("MISSING|Node.js|not found on PATH", out)
        self.assertIn("RESULT: MISSING", out)
        self.assertIn("MISSING: Node.js", out)
        self.assertEqual(code, 1)

    def test_the_summary_names_every_missing_tool(self):
        env = _Env(node_path=None, npm_path=None)
        _, out, _ = self.run_check(env)
        summary = [l for l in out.splitlines() if l.startswith("MISSING:")][0]
        self.assertIn("Node.js", summary)
        self.assertIn("npm", summary)

    def test_the_failure_tells_the_agent_what_to_do_next(self):
        _, out, _ = self.run_check(_Env(node_path=None))
        self.assertIn(rutas.comando_apa7("install"), out)


class TestPythonDiagnostics(CheckHarness):
    """A broken venv and a missing Python need different instructions."""

    def test_no_python_at_all_mentions_the_override(self):
        _, out, _ = self.run_check(_Env(python_path=None, venv_python=None))
        self.assertIn("MISSING|Python|not found (set APA7_PYTHON", out)
        self.assertIn("MISSING|pymupdf|not verifiable: Python missing", out)

    def test_a_broken_venv_is_told_to_be_recreated_not_reinstalled(self):
        _, out, _ = self.run_check(_Env(python_path=None, venv_python="/skill/.venv/bin/python3"))
        self.assertIn("exists but does not run", out)
        self.assertIn("Recreate it with: " + rutas.comando_apa7("install"), out)

    def test_a_too_old_interpreter_is_reported_with_the_floor(self):
        old = lambda *a, **k: _Line("3.8.10 TOO-OLD", exit_code=3)
        _, out, _ = self.run_check(_Env(run=old))
        self.assertIn("MISSING|Python|3.8.10 is below the 3.9 floor", out)

    def test_a_version_mismatch_stays_ok_and_is_only_a_note(self):
        def run_with_other_pymupdf(argv, timeout=60, cwd=None):
            if "pymupdf" in " ".join(str(a) for a in argv):
                return _Line("1.24.0")
            return _fake_run(argv, timeout, cwd)

        code, out, _ = self.run_check(_Env(run=run_with_other_pymupdf))
        self.assertIn("MISMATCH: pinned", out)
        self.assertIn("RESULT: OK", out)
        self.assertEqual(code, 0)


class TestDocxResolution(CheckHarness):
    def test_docx_is_not_verifiable_without_node(self):
        _, out, _ = self.run_check(_Env(node_path=None))
        self.assertIn("MISSING|docx (npm)|not verifiable: Node.js missing", out)

    def test_a_missing_package_json_points_at_install(self):
        with tempfile_workdir() as empty:
            _, out, _ = self.run_check(_Env(node_dir=empty))
        self.assertIn("not installed in", out)
        self.assertIn(rutas.comando_apa7("install"), out)


class TestLibreOffice(CheckHarness):
    def test_not_found_lists_the_places_that_were_searched(self):
        _, out, _ = self.run_check(_Env(soffice_path=None,
                                         soffice_candidates=["/usr/bin/soffice", "/opt/bin/soffice"]))
        self.assertIn("MISSING|LibreOffice|not found. Searched: /usr/bin/soffice | /opt/bin/soffice", out)

    def test_an_unreadable_version_is_a_failure_with_the_exit_code(self):
        env = _Env(run_soffice=lambda *a, **k: _Line(stdout="", exit_code=1))
        _, out, _ = self.run_check(env)
        self.assertIn("MISSING|LibreOffice|could not read the version (exit code 1)", out)

    def test_a_wrong_version_of_the_package_is_only_a_note(self):
        def run_with_json(argv, timeout=60, cwd=None):
            return _Line("ok")

        with tempfile_workdir() as workdir:
            pkg = Path(workdir) / "node_modules" / "docx" / "package.json"
            pkg.parent.mkdir(parents=True)
            pkg.write_text('{"version": "9.0.0"}', encoding="utf-8")
            _, out, _ = self.run_check(_Env(node_dir=Path(workdir), run=run_with_json))
        self.assertIn("MISMATCH: pinned 9.7.1", out)


@contextlib.contextmanager
def tempfile_workdir():
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        yield Path(tmp)


class TestExport(CheckHarness):
    """stdout carries the artifact and nothing else, so the caller can use it."""

    def setUp(self):
        import tempfile
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.dir = Path(self._tmp.name)
        self.docx = self.dir / "doc.docx"
        self.docx.write_bytes(b"PK\x03\x04fake")

    def _args(self, **overrides):
        values = {"docx": str(self.docx), "outdir": str(self.dir / "out"),
                  "log": None, "timeout": 300, "carpeta_trabajo": None}
        values.update(overrides)
        return argparse.Namespace(**values)

    def _export(self, soffice_result, **overrides):
        import argparse

        def fake_run_soffice(arguments, timeout=300, log_dir=None):
            # Create what LibreOffice would have created.
            destino = Path(values.outdir) if values.outdir else self.docx.parent
            expected = destino / (self.docx.stem + ".pdf")
            if getattr(soffice_result, "exit_code", 0) == 0 and not getattr(
                    soffice_result, "timed_out", False):
                expected.parent.mkdir(parents=True, exist_ok=True)
                expected.write_bytes(b"%PDF-1.4 fake")
            return soffice_result

        values = self._args(**overrides)
        out, err = io.StringIO(), io.StringIO()
        with mock.patch.object(rutas, "soffice_console", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "kill_soffice_processes", return_value=0), \
                mock.patch.object(rutas, "run_soffice", fake_run_soffice):
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = apa7.cmd_export(values)
        return code, out.getvalue(), err.getvalue()

    def _ok_result(self, **overrides):
        profile = self.dir / "lo_profile"
        profile.mkdir(parents=True, exist_ok=True)
        values = {"exit_code": 0, "stdout": "convert ... -> doc.pdf", "stderr": "",
                  "timed_out": False, "profile_path": str(profile)}
        values.update(overrides)
        return _ExportResult(values)

    def test_a_successful_export_prints_only_the_pdf_path(self):
        code, out, _ = self._export(self._ok_result())
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), str(self.dir / "out" / "doc.pdf"))
        self.assertEqual(len(out.strip().splitlines()), 1)

    def test_the_log_goes_to_stderr_and_to_a_file(self):
        _, out, err = self._export(self._ok_result())
        self.assertNotIn("PHASE 3", out)
        self.assertIn("PHASE 3", err)
        # --outdir was explicit, so it keeps meaning "PDF and logs here", which
        # is the behavior every existing caller already depends on.
        log_file = self.dir / "out" / "_logs" / "03-export.log"
        self.assertTrue(log_file.is_file())
        self.assertIn("PDF generated", log_file.read_text(encoding="utf-8"))

    def test_without_outdir_the_log_goes_to_the_documents_own_folder(self):
        # No _logs/ next to the user's document: that pile is what this change
        # exists to remove.
        self._export(self._ok_result(), outdir=None)
        log_file = self.dir / "doc_apa" / "logs" / "03-export.log"
        self.assertTrue(log_file.is_file(), "03-export.log is not in %s" % log_file)
        self.assertFalse((self.dir / "_logs").exists())

    def test_carpeta_trabajo_redirects_the_logs_without_moving_the_pdf(self):
        self._export(self._ok_result(), outdir=None,
                     carpeta_trabajo=str(self.dir / "trabajo"))
        self.assertTrue((self.dir / "trabajo" / "logs" / "03-export.log").is_file())
        self.assertTrue((self.dir / "doc.pdf").is_file())

    def test_an_explicit_log_is_respected(self):
        _, _, _ = self._export(self._ok_result(), log=str(self.dir / "mio.log"))
        self.assertTrue((self.dir / "mio.log").is_file())

    def test_a_missing_docx_fails_without_writing_to_stdout(self):
        code, out, err = self._export(self._ok_result(), docx=str(self.dir / "nope.docx"))
        self.assertEqual(code, 1)
        self.assertEqual(out, "")
        self.assertIn("does not exist", err)

    def test_a_timeout_is_reported_as_a_timeout_not_as_a_missing_file(self):
        # The order matters: a wedged LibreOffice diagnosed as "no PDF" sends
        # the user to look for a path problem instead of a stuck process.
        code, out, err = self._export(self._ok_result(timed_out=True, exit_code=124,
                                                      stdout=""))
        self.assertEqual(code, 1)
        self.assertEqual(out, "")
        self.assertIn("did not finish within", err)
        self.assertIn("Raise --timeout", err)
        self.assertLess(err.index("did not finish within"), err.index("was killed"))

    def test_the_isolated_profile_is_removed_even_on_failure(self):
        profile = self.dir / "lo_profile"
        result = self._ok_result(timed_out=True, exit_code=124, stdout="")
        self._export(result)
        self.assertFalse(profile.exists(), "the temporary profile leaked")

    def test_benign_libreoffice_noise_is_reported_as_ignored(self):
        noise = "Could not find platform independent libraries C:\\Python\\312\n"
        _, _, err = self._export(self._ok_result(stderr=noise))
        self.assertIn("only benign LibreOffice noise (ignored)", err)

    def test_a_real_stderr_line_is_surfaced(self):
        noise = ("Warning: failed to launch javaldx\n"
                 "Error: source file could not be loaded\n")
        _, _, err = self._export(self._ok_result(stderr=noise))
        self.assertIn("Error: source file could not be loaded", err)

    def test_our_own_leftovers_are_cleaned_before_and_after(self):
        calls = []
        profile = self.dir / "lo_profile"
        profile.mkdir(parents=True, exist_ok=True)
        expected = self.dir / "out" / "doc.pdf"

        def fake_run_soffice(arguments, timeout=300, log_dir=None):
            expected.parent.mkdir(parents=True, exist_ok=True)
            expected.write_bytes(b"%PDF")
            calls.append("convert")
            return self._ok_result()

        def fake_clean(perfil=None):
            calls.append(("clean", perfil))
            return 0

        with mock.patch.object(rutas, "soffice_console", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "kill_soffice_processes", side_effect=fake_clean), \
                mock.patch.object(rutas, "run_soffice", fake_run_soffice):
            with contextlib.redirect_stdout(io.StringIO()), \
                    contextlib.redirect_stderr(io.StringIO()):
                apa7.cmd_export(self._args())
        # Only our isolated profile is cleaned, before and after.
        self.assertEqual(calls, [("clean", "lo_profile"), "convert",
                                 ("clean", "lo_profile")])

    def test_the_users_libreoffice_is_not_touched_by_default(self):
        perfiles = []
        expected = self.dir / "out" / "doc.pdf"

        def fake_run_soffice(arguments, timeout=300, log_dir=None):
            expected.parent.mkdir(parents=True, exist_ok=True)
            expected.write_bytes(b"%PDF")
            return self._ok_result()

        def fake_clean(perfil=None):
            perfiles.append(perfil)
            return 0

        with mock.patch.object(rutas, "soffice_console", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "kill_soffice_processes", side_effect=fake_clean), \
                mock.patch.object(rutas, "run_soffice", fake_run_soffice):
            with contextlib.redirect_stdout(io.StringIO()), \
                    contextlib.redirect_stderr(io.StringIO()):
                apa7.cmd_export(self._args())
        self.assertEqual(perfiles, ["lo_profile", "lo_profile"])
        self.assertNotIn(None, perfiles)

    def test_cerrar_libreoffice_closes_every_process_first(self):
        perfiles = []
        expected = self.dir / "out" / "doc.pdf"

        def fake_run_soffice(arguments, timeout=300, log_dir=None):
            expected.parent.mkdir(parents=True, exist_ok=True)
            expected.write_bytes(b"%PDF")
            return self._ok_result()

        def fake_clean(perfil=None):
            perfiles.append(perfil)
            return 2

        with mock.patch.object(rutas, "soffice_console", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "kill_soffice_processes", side_effect=fake_clean), \
                mock.patch.object(rutas, "run_soffice", fake_run_soffice):
            out, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = apa7.cmd_export(self._args(cerrar_libreoffice=True))
        self.assertEqual(code, 0)
        self.assertEqual(perfiles, [None, "lo_profile"])
        self.assertIn("unsaved", err.getvalue())

    def test_a_missing_libreoffice_fails_before_running_anything(self):
        out, err = io.StringIO(), io.StringIO()
        with mock.patch.object(rutas, "soffice_console", return_value=None), \
                mock.patch.object(rutas, "run_soffice") as run_soffice:
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = apa7.cmd_export(self._args())
        self.assertEqual(code, 1)
        self.assertEqual(out.getvalue(), "")
        run_soffice.assert_not_called()
        self.assertIn("apa7.py install", err.getvalue())

    def test_without_docx_the_only_one_of_the_folder_is_exported(self):
        code, out, err = self._export(self._ok_result(), docx=None, outdir=None,
                                      carpeta_trabajo=str(self.dir))
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), str(self.dir / "doc.pdf"))
        self.assertIn("using the only .docx", err)

    def test_several_docx_files_are_ambiguous_and_nothing_is_exported(self):
        (self.dir / "otro.docx").write_bytes(b"PK\x03\x04fake")
        code, out, err = self._export(self._ok_result(), docx=None, outdir=None,
                                      carpeta_trabajo=str(self.dir))
        self.assertEqual(code, 1)
        self.assertEqual(out.strip(), "")
        self.assertIn("ambiguous", err)
        self.assertIn("otro.docx", err)
        self.assertFalse((self.dir / "doc.pdf").exists())

    def test_a_folder_without_any_docx_is_reported(self):
        vacio = self.dir / "vacio"
        vacio.mkdir()
        code, _, err = self._export(self._ok_result(), docx=None, outdir=None,
                                    carpeta_trabajo=str(vacio))
        self.assertEqual(code, 1)
        self.assertIn("No .docx in", err)

    def test_an_explicit_docx_wins_over_the_ones_in_the_folder(self):
        (self.dir / "otro.docx").write_bytes(b"PK\x03\x04fake")
        code, out, err = self._export(self._ok_result(), outdir=None,
                                      carpeta_trabajo=str(self.dir))
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), str(self.dir / "doc.pdf"))
        self.assertNotIn("ambiguous", err)


class _ExportResult:
    def __init__(self, values):
        self.__dict__.update(values)
        self.stderr = self.__dict__.get("stderr", "")
        self.stdout = self.__dict__.get("stdout", "")


class TestForwarding(CheckHarness):
    """parse/build/verify must hand their arguments over untouched.

    The dispatcher does not know the target script's options, so anything that
    reshapes the argument list corrupts the command line. The regression this
    guards against is silent: argparse.REMAINDER splits the list in two and
    re-concatenating reorders it, so `--md a.md --out b.json` arrives as
    `a.md --out b.json --md` and the target reports a missing value.
    """

    def setUp(self):
        # Real files: cmd_forward checks is_file() before running anything, so
        # imaginary paths would short-circuit every case.
        scratch = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, scratch, True)
        self.root = Path(scratch)
        (self.root / "scripts").mkdir()
        for name in ("md-a-manifiesto.py", "build-docx.js", "verificar-pdf.py"):
            (self.root / "scripts" / name).write_text("", encoding="utf-8")
        self.venv = str(self.root / ".venv" / "bin" / "python3")

        # Relative paths are resolved against the current directory, and these
        # subcommands now derive a working folder from the document they are
        # handed, so a test must not drop `a_apa/` into the repository.
        previous = os.getcwd()
        self.addCleanup(os.chdir, previous)
        os.chdir(str(self.root))

    def call(self, argv, **probes):
        self.calls = []
        root = self.root

        defaults = {
            "skill_script": lambda name: str(root / "scripts" / name),
            "venv_python": str(root / ".venv" / "bin" / "python3"),
            "python_works": True,
            "node_path": "/usr/bin/node",
            "skill_root": root,
            "is_windows": False,
            "is_macos": False,
        }
        defaults.update(probes)

        def record(command):
            self.calls.append([str(part) for part in command])
            return 0

        with contextlib.ExitStack() as stack:
            for name, value in defaults.items():
                # rutas exposes these as functions, so a plain literal would be
                # stored where a callable is expected.
                if callable(value):
                    stack.enter_context(mock.patch.object(rutas, name, value))
                else:
                    stack.enter_context(
                        mock.patch.object(rutas, name, mock.Mock(return_value=value)))
            stack.enter_context(mock.patch.object(apa7.subprocess, "call", record))
            out, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = apa7.main(argv)
        return code, self.calls, out.getvalue(), err.getvalue()

    def test_the_argument_order_survives_intact(self):
        # The regression guarded here is a REORDER, not an addition: defaults we
        # compute are appended at the end, but everything the user wrote has to
        # arrive in the order they wrote it.
        _, calls, _, _ = self.call(
            ["parse", "--md", "a.md", "--out", "b.json", "--base-dir", "img"])
        recibidos = calls[0][2:]
        self.assertEqual(recibidos[:6],
                         ["--md", "a.md", "--out", "b.json", "--base-dir", "img"])
        # --out was given, so the log is the only thing we added, and it is added
        # at the end rather than in the middle of what the user wrote.
        self.assertEqual(
            recibidos[6:],
            ["--log", str(self.root / "a_apa" / "logs" / "01-analisis.log")])

    def test_nothing_is_appended_when_nothing_is_missing(self):
        _, calls, _, _ = self.call(["parse", "--md", "a.md", "--out", "b.json",
                                    "--log", "c.log"])
        self.assertEqual(calls[0][2:],
                         ["--md", "a.md", "--out", "b.json", "--log", "c.log"])

    def test_an_option_whose_value_looks_like_a_flag_stays_attached(self):
        _, calls, _, _ = self.call(["verify", "--pdf", "--weird.pdf",
                                    "--manifiesto", "m.json"])
        self.assertEqual(calls[0][2:6],
                         ["--pdf", "--weird.pdf", "--manifiesto", "m.json"])

    def test_help_reaches_the_target_script(self):
        # The agent needs the REAL script's options; apa7.py's own help would
        # list none of them.
        _, calls, _, _ = self.call(["parse", "--help"])
        self.assertIn("--help", calls[0])
        self.assertTrue(calls[0][1].endswith("md-a-manifiesto.py"))

    def test_parse_uses_the_current_interpreter(self):
        _, calls, _, _ = self.call(["parse", "--md", "a.md"])
        self.assertEqual(calls[0][0], sys.executable)

    def test_build_runs_node_on_the_js_script(self):
        _, calls, _, _ = self.call(["build", "--manifiesto", "m.json"])
        self.assertEqual(calls[0][0], "/usr/bin/node")
        self.assertTrue(calls[0][1].endswith("build-docx.js"))

    def test_build_without_node_fails_before_running_anything(self):
        code, calls, out, err = self.call(["build", "--manifiesto", "m.json"],
                                          node_path=None)
        self.assertEqual(code, 1)
        self.assertEqual(calls, [])
        self.assertEqual(out, "")
        self.assertIn("install --only node", err)

    def test_verify_prefers_the_virtual_environment_interpreter(self):
        # pymupdf lives in the venv, so running verify on any other interpreter
        # either fails to import or silently uses another version.
        code, calls, _, _ = self.call(["verify", "--pdf", "p.pdf"])
        self.assertEqual(code, 0)
        self.assertEqual(calls[0][0], self.venv)
        self.assertTrue(calls[0][1].endswith("verificar-pdf.py"))

    def test_verify_warns_and_falls_back_when_the_venv_is_unusable(self):
        code, calls, _, err = self.call(["verify", "--pdf", "p.pdf"],
                                        venv_python=None)
        self.assertEqual(code, 0)
        self.assertEqual(calls[0][0], sys.executable)
        self.assertIn("WARNING", err)
        self.assertIn("install --only pymupdf", err)

    def test_verify_falls_back_when_the_venv_cannot_run(self):
        code, calls, _, err = self.call(["verify", "--pdf", "p.pdf"],
                                        python_works=False)
        self.assertEqual(calls[0][0], sys.executable)
        self.assertIn("WARNING", err)

    def test_the_exit_code_of_the_target_script_is_propagated(self):
        with mock.patch.object(apa7.subprocess, "call", return_value=3):
            with contextlib.redirect_stdout(io.StringIO()), \
                    contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(apa7.main(["verify", "--pdf", "p.pdf"]), 3)

    def test_no_arguments_prints_the_usage_and_exits_two(self):
        for command in ("parse", "build", "verify"):
            code, calls, out, err = self.call([command])
            self.assertEqual(code, 2, command)
            self.assertEqual(calls, [], command)
            self.assertIn("Usage: ", err)
            self.assertIn("apa7.py", err)
            self.assertIn(command, err)
            self.assertNotIn("python scripts/apa7.py", err)


class TestRutasPorDefecto(TestForwarding):
    """What lands where when the caller passes no paths at all.

    This is the change the whole task is about: no more MANIFEST.json, document.docx
    and three .log files dropped next to the user's .md.
    """

    def _flags(self, llamada, nombre):
        partes = llamada
        for i, parte in enumerate(partes):
            if parte == nombre:
                return partes[i + 1]
        self.fail("%s not present in %s" % (nombre, partes))

    def test_parse_writes_the_manifest_into_the_working_folder(self):
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        _, calls, _, _ = self.call(["parse", "--md", "informe.md"])
        salida = self._flags(calls[0], "--out")
        self.assertEqual(Path(salida), self.root / "informe_apa" / "datos" / "MANIFEST.json")
        log = self._flags(calls[0], "--log")
        self.assertEqual(Path(log), self.root / "informe_apa" / "logs" / "01-analisis.log")

    def test_parse_records_which_md_the_folder_belongs_to(self):
        # Without this anchor, build/export/verify are handed a manifest and a
        # .docx with no idea which document they belong to.
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        self.call(["parse", "--md", "informe.md"])
        fuente = self.root / "informe_apa" / "datos" / "fuente.json"
        self.assertTrue(fuente.is_file())
        self.assertEqual(json.loads(fuente.read_text(encoding="utf-8"))["nombre"], "informe.md")

    def test_the_working_folder_is_the_only_thing_created_next_to_the_md(self):
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        self.call(["parse", "--md", "informe.md"])
        self.assertEqual(sorted(p.name for p in self.root.iterdir()
                                if p.name != "scripts"),
                         ["informe.md", "informe_apa"])
        self.assertEqual(sorted(p.name for p in
                                (self.root / "informe_apa").iterdir()),
                         ["datos", "logs"])

    def test_build_writes_the_docx_next_to_the_md_not_inside_the_folder(self):
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        datos = self.root / "informe_apa" / "datos"
        datos.mkdir(parents=True)
        (datos / "MANIFEST.json").write_text("{}", encoding="utf-8")
        (datos / "fuente.json").write_text(
            json.dumps({"md": str(self.root / "informe.md"), "nombre": "informe.md"}),
            encoding="utf-8")
        _, calls, _, _ = self.call(["build", "--manifiesto",
                                    str(datos / "MANIFEST.json")])
        self.assertEqual(Path(self._flags(calls[0], "--out")),
                         self.root / "informe.docx")
        self.assertEqual(Path(self._flags(calls[0], "--log")),
                         self.root / "informe_apa" / "logs" / "02-build.log")

    def test_the_docx_is_never_called_manifest_docx(self):
        # The regression this catches: naming the deliverable after the artifact
        # that happened to be passed instead of after the document.
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        datos = self.root / "informe_apa" / "datos"
        datos.mkdir(parents=True)
        (datos / "MANIFEST.json").write_text("{}", encoding="utf-8")
        rutas.anota_fuente(self.root / "informe_apa", self.root / "informe.md")
        _, calls, _, _ = self.call(["build", "--manifiesto",
                                    str(datos / "MANIFEST.json")])
        self.assertEqual(Path(self._flags(calls[0], "--out")).name, "informe.docx")

    def test_verify_gets_a_log_and_a_report_by_default(self):
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        _, calls, _, _ = self.call(["verify", "--pdf", "p.pdf", "--manifiesto", "m.json"])
        self.assertEqual(Path(self._flags(calls[0], "--log")),
                         self.root / "p_apa" / "logs" / "04-verificacion.log")
        self.assertEqual(Path(self._flags(calls[0], "--json")),
                         self.root / "p_apa" / "datos" / "verificacion.json")

    def test_an_explicit_flag_always_wins(self):
        # The one rule the whole defaulting scheme rests on.
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        (self.root / "mio.json").write_text("{}", encoding="utf-8")
        (self.root / "mio.log").write_text("", encoding="utf-8")
        _, calls, _, _ = self.call(["parse", "--md", "informe.md",
                                    "--out", "mio.json", "--log", "mio.log"])
        self.assertEqual(self._flags(calls[0], "--out"), "mio.json")
        self.assertEqual(self._flags(calls[0], "--log"), "mio.log")

    def test_carpeta_trabajo_moves_everything_and_never_reaches_the_script(self):
        # The target scripts have never heard of this option; forwarded verbatim
        # they would abort on an unknown argument instead of reading the document.
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        fuera = self.root / "trabajo"
        _, calls, _, _ = self.call(["parse", "--md", "informe.md",
                                    "--carpeta-trabajo", str(fuera)])
        self.assertNotIn("--carpeta-trabajo", calls[0])
        self.assertEqual(Path(self._flags(calls[0], "--out")),
                         fuera / "datos" / "MANIFEST.json")

    def test_carpeta_trabajo_also_works_with_the_equals_form(self):
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        fuera = self.root / "trabajo"
        _, calls, _, _ = self.call(["parse", "--md", "informe.md",
                                    "--carpeta-trabajo=%s" % fuera])
        self.assertNotIn("--carpeta-trabajo=%s" % fuera, calls[0])
        self.assertEqual(Path(self._flags(calls[0], "--out")),
                         fuera / "datos" / "MANIFEST.json")

    def test_a_missing_md_is_left_to_the_target_script(self):
        # No anchor, no defaults, and no traceback out of the dispatcher either.
        _, calls, out, _ = self.call(["parse", "--base-dir", "img"])
        self.assertEqual(calls[0][2:], ["--base-dir", "img"])
        self.assertEqual(out, "")

    def test_rerunning_replaces_the_files_and_adds_nothing(self):
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        self.call(["parse", "--md", "informe.md"])
        self.call(["parse", "--md", "informe.md"])
        trabajo = self.root / "informe_apa"
        self.assertEqual(sorted(p.name for p in trabajo.iterdir()), ["datos", "logs"])
        # The manifest is written by the target script, which is mocked here; what
        # the dispatcher itself must not accumulate is the anchor.
        self.assertEqual(sorted(p.name for p in trabajo.joinpath("datos").iterdir()),
                         ["fuente.json"])
        self.assertEqual(sorted(p.name for p in self.root.iterdir()
                                if p.name != "scripts"),
                         ["informe.md", "informe_apa"])

    def test_a_warning_about_a_reused_folder_goes_to_stderr(self):
        # A folder pointed at a second document is a warning, not an error: the
        # .docx is a function of the .md just passed, and a run with no terminal
        # cannot ask anything.
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        (self.root / "otro.md").write_text("# y", encoding="utf-8")
        trabajo = self.root / "trabajo"
        self.call(["parse", "--md", "informe.md", "--carpeta-trabajo", str(trabajo)])
        _, _, out, err = self.call(["parse", "--md", "otro.md",
                                    "--carpeta-trabajo", str(trabajo)])
        self.assertIn("WARNING", err)
        self.assertIn("otro.md", err)
        self.assertEqual(out, "")
        self.assertEqual(json.loads(
            (trabajo / "datos" / "fuente.json").read_text(encoding="utf-8"))["nombre"],
            "otro.md")

    def test_two_documents_get_two_working_folders(self):
        # The folder is named after the document, so two documents never collide.
        (self.root / "informe.md").write_text("# x", encoding="utf-8")
        (self.root / "otro.md").write_text("# y", encoding="utf-8")
        self.call(["parse", "--md", "informe.md"])
        _, _, _, err = self.call(["parse", "--md", "otro.md"])
        self.assertEqual(err, "")
        self.assertTrue((self.root / "informe_apa").is_dir())
        self.assertTrue((self.root / "otro_apa").is_dir())

    def test_inyecta_never_overrides_what_the_user_wrote(self):
        self.assertEqual(apa7._inyecta(["--out", "x"], "--out", "y"), ["--out", "x"])
        self.assertEqual(apa7._inyecta(["--out=x"], "--out", "y"), ["--out=x"])
        self.assertEqual(apa7._inyecta(["--a", "1"], "--out", "y"), ["--a", "1", "--out", "y"])
        self.assertEqual(apa7._inyecta(["--a", "1"], "--out", None), ["--a", "1"])

    def test_extrae_removes_the_option_and_leaves_the_order_intact(self):
        self.assertEqual(apa7._extrae(["--a", "1", "--carpeta-trabajo", "x", "--b", "2"],
                                      "--carpeta-trabajo"), ("x", ["--a", "1", "--b", "2"]))
        self.assertEqual(apa7._extrae(["--carpeta-trabajo=x", "--a"], "--carpeta-trabajo"),
                         ("x", ["--a"]))
        self.assertEqual(apa7._extrae(["--a", "1"], "--carpeta-trabajo"), (None, ["--a", "1"]))

    def test_a_missing_target_script_is_reported_not_raised(self):
        code, calls, out, err = self.call(["parse", "--md", "a.md"],
                                          skill_script=lambda name: "/nope/" + name)
        self.assertEqual(code, 1)
        self.assertEqual(calls, [])
        self.assertIn("not found at", err)

    def test_the_strict_subcommands_still_reject_unknown_arguments(self):
        for command in (["check", "--bogus"], ["export", "--nope"],
                        ["install", "--nope"]):
            with contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as raised:
                    apa7.main(command)
            self.assertEqual(raised.exception.code, 2, command)


class TestNoHardcodedPaths(CheckHarness):
    """Every machine-specific path has to be resolved, never written down."""

    def test_the_source_contains_no_absolute_machine_paths(self):
        source = Path(apa7.__file__).read_text(encoding="utf-8")
        for needle in ("C:\\", "C:/", "/Users/", "\\\\Users\\\\"):
            self.assertNotIn(needle, source, "hardcoded path in apa7.py: %s" % needle)

    def test_it_imports_nothing_outside_the_standard_library(self):
        # Checked on the parsed tree, not with a substring search: `check` runs
        # a subprocess probe that contains the text "import pymupdf", and that
        # is a string handed to another interpreter, not an import of ours.
        tree = ast.parse(Path(apa7.__file__).read_text(encoding="utf-8"))
        imported = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imported.update(a.name.split(".")[0] for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.level == 0:
                imported.add((node.module or "").split(".")[0])
        imported.discard("")
        self.assertTrue(imported)
        allowed = {"argparse", "contextlib", "io", "json", "os", "sys", "time",
           "datetime", "pathlib", "subprocess", "lib"}
        self.assertEqual(imported - allowed, set(),
                         "third-party import in apa7.py: %s" % (imported - allowed))


class TestDoblePasadaDeIndices(unittest.TestCase):
    """The two-pass index build: the second run must carry --paginas-json."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="apa7_doble_"))
        self.addCleanup(shutil.rmtree, str(self.tmp), True)

    def _documento(self, md="salida"):
        return rutas.rutas_documento(md=self.tmp / (md + ".md"), crear=True)

    def test_valor_de_lee_las_dos_formas_del_flag(self):
        self.assertEqual(apa7._valor_de(["--out", "a.docx"], "--out"), "a.docx")
        self.assertEqual(apa7._valor_de(["--out=a.docx"], "--out"), "a.docx")
        self.assertIsNone(apa7._valor_de(["--otro", "x"], "--out"))

    def test_manifiesto_con_toc_respeta_toc_campos_false(self):
        tmp = Path(tempfile.mkdtemp(prefix="apa7_toc_"))
        self.addCleanup(shutil.rmtree, str(tmp), True)
        ruta = tmp / "m.json"
        ruta.write_text(json.dumps({"opciones": {"toc_campos": False}}), encoding="utf-8")
        self.assertIsNone(apa7._manifiesto_con_toc(["--manifiesto", str(ruta)]))
        ruta.write_text(json.dumps({"opciones": {}}), encoding="utf-8")
        self.assertIsNotNone(apa7._manifiesto_con_toc(["--manifiesto", str(ruta)]))
        self.assertIsNone(apa7._manifiesto_con_toc(["--manifiesto", str(tmp / "nope.json")]))

    def test_sin_out_no_hay_segunda_pasada(self):
        llamadas = []
        with mock.patch.object(apa7.subprocess, "call",
                               side_effect=lambda c: llamadas.append(c) or 0):
            codigo = apa7._segunda_pasada_indices("/usr/bin/node", "build.js",
                                                  ["--manifiesto", "m.json"],
                                                  self._documento())
        self.assertEqual(codigo, 0)
        self.assertEqual(llamadas, [])

    def _correr(self, codigo_rebuild=0, docx_nombre="salida"):
        docx = self.tmp / (docx_nombre + ".docx")
        docx.write_text("x", encoding="utf-8")
        manifest = self.tmp / "m.json"
        manifest.write_text("{}", encoding="utf-8")
        (self.tmp / "paginas-de-pdf.py").write_text("", encoding="utf-8")
        rest = ["--manifiesto", str(manifest), "--out", str(docx)]
        documento = rutas.rutas_documento(md=self.tmp / (docx_nombre + ".md"))
        llamadas = []

        def fake_call(comando):
            partes = [str(c) for c in comando]
            llamadas.append(partes)
            # The first call is paginas-de-pdf.py: emulate the JSON it writes.
            destino = Path(partes[partes.index("--out") + 1])
            destino.parent.mkdir(parents=True, exist_ok=True)
            destino.write_text('{"secciones":{},"tablas":{},"figuras":{}}', encoding="utf-8")
            return 0

        exports = []

        def fake_export(ns):
            exports.append(ns)
            (Path(ns.outdir) / (Path(ns.docx).stem + ".pdf")).write_bytes(b"%PDF-1.4")
            return 0

        with mock.patch.object(apa7, "cmd_export", fake_export), \
             mock.patch.object(rutas, "venv_python", return_value="C:/py"), \
             mock.patch.object(rutas, "python_works", return_value=True), \
             mock.patch.object(rutas, "skill_script", side_effect=lambda n: str(self.tmp / n)), \
             mock.patch.object(apa7.subprocess, "call", fake_call):
            codigo = apa7._segunda_pasada_indices("/usr/bin/node",
                                                  str(self.tmp / "build-docx.js"),
                                                  rest, documento)
        return codigo, llamadas, exports, documento

    def test_la_segunda_pasada_rebuilds_con_paginas_json(self):
        codigo, llamadas, _, _ = self._correr()
        self.assertEqual(codigo, 0)
        self.assertEqual(len(llamadas), 2)          # measure + rebuild
        self.assertTrue(llamadas[0][1].endswith("paginas-de-pdf.py"))
        self.assertTrue(llamadas[1][1].endswith("build-docx.js"))
        self.assertIn("--paginas-json", llamadas[1])

    def test_los_temporales_van_a_logs_y_no_a_la_carpeta_del_usuario(self):
        # The old code wrote <docx folder>/_logs/, which meant a _logs/ directory
        # beside the document the user was trying to tidy up.
        _, _, exports, documento = self._correr()
        self.assertEqual(Path(exports[0].outdir), documento.logs)
        self.assertEqual(Path(exports[0].log), documento.log_paginas)
        self.assertEqual(documento.paginas_json.parent, documento.logs)
        self.assertFalse((self.tmp / "_logs").exists())
        self.assertTrue(documento.paginas_json.is_file())

    def test_el_pdf_desechable_se_borra_al_terminar(self):
        _, _, _, documento = self._correr()
        self.assertTrue(str(documento.pdf_auxiliar).startswith(str(documento.logs)))
        self.assertFalse(documento.pdf_auxiliar.exists(),
                         "the throwaway PDF was left in logs/")
        # The evidence of the measured pages is what stays.
        self.assertTrue(documento.paginas_json.is_file())

    def test_el_pdf_desechable_no_se_borra_si_la_segunda_pasada_falla(self):
        docx = self.tmp / "salida.docx"
        docx.write_text("x", encoding="utf-8")
        manifest = self.tmp / "m.json"
        manifest.write_text("{}", encoding="utf-8")
        (self.tmp / "paginas-de-pdf.py").write_text("", encoding="utf-8")
        documento = rutas.rutas_documento(md=self.tmp / "salida.md")
        documento.logs.mkdir(parents=True, exist_ok=True)
        aux = documento.pdf_auxiliar
        aux.write_bytes(b"%PDF-1.4")

        def fake_call(comando):
            partes = [str(c) for c in comando]
            destino = Path(partes[partes.index("--out") + 1])
            destino.parent.mkdir(parents=True, exist_ok=True)
            destino.write_text("{}", encoding="utf-8")
            return 3 if partes[1].endswith("build-docx.js") else 0

        def fake_export(ns):
            (Path(ns.outdir) / (Path(ns.docx).stem + ".pdf")).write_bytes(b"%PDF-1.4")
            return 0

        with mock.patch.object(apa7, "cmd_export", fake_export), \
             mock.patch.object(rutas, "venv_python", return_value="C:/py"), \
             mock.patch.object(rutas, "python_works", return_value=True), \
             mock.patch.object(rutas, "skill_script", side_effect=lambda n: str(self.tmp / n)), \
             mock.patch.object(apa7.subprocess, "call", fake_call):
            codigo = apa7._segunda_pasada_indices("/usr/bin/node",
                                                  str(self.tmp / "build-docx.js"),
                                                  ["--manifiesto", str(manifest),
                                                   "--out", str(docx)], documento)
        self.assertEqual(codigo, 3)
        self.assertTrue(aux.exists(), "the PDF that failed to convert was deleted")


class TestCompatibleWithPython39(unittest.TestCase):
    """The scripts must PARSE under Python 3.9.

    ast.parse(feature_version=(3, 9)) rejects grammar that 3.9 did not have
    (the match statement, for instance). It is not a full compatibility check:
    `int | None` in an annotation parses fine under 3.9 and only fails when
    evaluated, and no standard-library API is inspected. So this proves syntax
    only, not that every API used exists in 3.9.
    """

    def test_every_python_file_parses_as_python_39(self):
        scripts = Path(__file__).resolve().parents[1]
        files = sorted(scripts.rglob("*.py"))
        self.assertTrue(files, "no .py files found under %s" % scripts)
        for path in files:
            try:
                ast.parse(path.read_text(encoding="utf-8"),
                          filename=str(path), feature_version=(3, 9))
            except SyntaxError as exc:
                self.fail("not valid Python 3.9 syntax: %s (%s)" % (path, exc))


if __name__ == "__main__":
    unittest.main()