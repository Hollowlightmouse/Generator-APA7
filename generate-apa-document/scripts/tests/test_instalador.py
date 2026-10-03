"""Tests for the install path: the plan tables, the elevation rules, and the
single confirmation gate.

The expensive mistake to avoid is an installer that half-runs: it prompts, then
installs something it was not allowed to install. So the tests below assert on
negative space -- which commands were NOT run -- as much as on what was.
"""

import argparse
import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import apa7  # noqa: E402
from lib import instalador  # noqa: E402
from lib import rutas  # noqa: E402


def manager(name):
    return rutas.PackageManager(name, {"winget": "winget.exe", "brew": "/opt/homebrew/bin/brew",
                                       "apt": "/usr/bin/apt-get", "dnf": "/usr/bin/dnf",
                                       "pacman": "/usr/bin/pacman"}[name])


class TestPlans(unittest.TestCase):
    def test_every_tool_declares_a_linux_command(self):
        # apt is the documented Linux path, so its row must never be missing.
        for tool in ("node", "python", "libreoffice"):
            self.assertIn("apt", instalador.PLANS[tool], "%s has no apt plan" % tool)
            self.assertTrue(instalador.PLANS[tool]["apt"])

    def test_the_python_plan_installs_venv_support(self):
        # Without python3-venv, `python3 -m venv` fails and the pinned pymupdf
        # is never installed -- so this row is load-bearing.
        self.assertIn("python3-venv", instalador.PLANS["python"]["apt"])
        self.assertTrue(any("pip" in str(argument)
                            for argument in instalador.PLANS["python"]["apt"]))

    def test_the_libreoffice_plan_asks_for_writer_only(self):
        self.assertIn("libreoffice-writer", instalador.PLANS["libreoffice"]["apt"])

    def test_pymupdf_and_docx_are_never_left_to_the_package_manager(self):
        # They are pinned by the skill and installed into its own dirs, so no
        # distro package must ever be trusted with them.
        for tool in ("docx", "pymupdf"):
            self.assertNotIn(tool, instalador.PLANS)

    def test_the_docx_version_is_pinned(self):
        self.assertEqual(rutas.DEPS["docx"], "9.7.1")
        self.assertEqual(rutas.DEPS["pymupdf"], "1.28.2")


class TestElevation(unittest.TestCase):
    def test_apt_is_elevated_with_a_non_interactive_sudo(self):
        argv = instalador.build(manager("apt"), "node")
        # -n matters more than sudo: without it sudo PROMPTS and an agent hangs.
        self.assertEqual(argv[:2], ["sudo", "-n"])

    def test_dnf_and_pacman_are_elevated_too(self):
        for name in ("dnf", "pacman"):
            self.assertEqual(instalador.build(manager(name), "node")[:2], ["sudo", "-n"])

    def test_winget_and_brew_are_never_elevated(self):
        for name in ("winget", "brew"):
            argv = instalador.build(manager(name), "node")
            self.assertNotIn("sudo", argv)
            self.assertEqual(argv[0], manager(name).path)

    def test_no_sudo_when_already_root(self):
        with mock.patch.object(instalador, "is_root", return_value=True):
            self.assertNotIn("sudo", instalador.build(manager("apt"), "node"))

    def test_an_unsupported_pair_returns_none_instead_of_a_guess(self):
        # None means "print manual instructions"; a wrong argv would install
        # the wrong thing.
        self.assertIsNone(instalador.build(manager("pacman"), "docx"))


class TestManualInstructions(unittest.TestCase):
    def test_each_tool_has_a_command_for_every_platform(self):
        for tool in instalador.TOOLS:
            for platform in ("windows", "macos", "linux"):
                self.assertTrue(
                    instalador.manual_instructions(tool, is_windows=platform == "windows",
                                                   is_macos=platform == "macos"),
                    "%s has no %s instructions" % (tool, platform))

    def test_the_platform_selects_the_command(self):
        windows = instalador.manual_instructions("node", is_windows=True)[0]
        macos = instalador.manual_instructions("node", is_macos=True)[0]
        self.assertIn("winget", windows)
        self.assertIn("brew", macos)
        self.assertNotEqual(windows, macos)


class TestNpmCommand(unittest.TestCase):
    def test_the_pin_and_the_anchor_flags_are_present(self):
        argv = instalador.npm_install_docx("/usr/bin/npm", Path("/tmp/n"), "9.7.1")
        self.assertIn("docx@9.7.1", argv)
        self.assertIn("--no-save", argv)
        self.assertIn("--no-package-lock", argv)

    def test_the_anchor_is_valid_json_and_private(self):
        import json
        parsed = json.loads(instalador.npm_anchor_json())
        self.assertTrue(parsed["private"])


class InstallHarness(unittest.TestCase):
    """Runs cmd_install with every environment probe pinned down.

    The fake installer also performs the filesystem effects the real command
    would have: otherwise cmd_install's own post-conditions ("is docx there
    now?") fail and the test asserts the wrong thing.
    """

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)
        self.state = {"node": False, "docx": False, "soffice": False, "pymupdf": False}
        self.runner = mock.Mock(side_effect=self._perform)

    def _perform(self, argv, **overrides):
        joined = [str(argument) for argument in argv]
        if joined[:1] == ["sudo"]:
            joined = joined[3:]
        text = " ".join(joined)

        if "npm" in text and "docx@" in text:
            (self.tmp / "node" / "node_modules" / "docx").mkdir(parents=True, exist_ok=True)
            self.state["docx"] = True
        if "-m" in joined and "venv" in joined:
            folder = "Scripts" if self.windows else "bin"
            name = "python.exe" if self.windows else "python3"
            exe = self.tmp / ".venv" / folder / name
            exe.parent.mkdir(parents=True, exist_ok=True)
            exe.write_text("", encoding="utf-8")
        if "pip" in text and "pymupdf" in text:
            self.state["pymupdf"] = True
        if "libreoffice" in text.lower():
            self.state["soffice"] = True
        if "nodejs" in text or "NodeJS" in text:
            self.state["node"] = True
        return 0

    def run_install(self, *, complete=False, manager_found=True, yes=False,
                    dry_run=False, only=None, stdin_tty=False, python=True,
                    windows=True):
        args = argparse.Namespace(only=only, yes=yes, dry_run=dry_run)
        self.windows = windows
        calls = {"n": 0}
        if complete:
            for key in self.state:
                self.state[key] = True

        def fake_check(_args):
            # Faithful model: the environment is complete only once every
            # tool has actually been installed. A refused confirmation leaves
            # a state entry False, so the final check still reports MISSING.
            calls["n"] += 1
            if complete:
                return 0
            return 0 if all(self.state.values()) else 1

        def fake_run(argv, **overrides):
            # The pymupdf version probe: only reports the pin once pip ran.
            if self.state["pymupdf"]:
                return rutas.NativeResult([rutas.DEPS["pymupdf"]], "", 0)
            return rutas.NativeResult(["ModuleNotFoundError"], "", 1)

        # A real file object: cmd_install asks isatty() before reading, and
        # patching that method on the real stdin would be global state.
        stdin = io.StringIO()
        stdin.isatty = lambda: stdin_tty

        out, err = io.StringIO(), io.StringIO()
        with contextlib.ExitStack() as stack:
            for patcher in (
                mock.patch.object(apa7, "cmd_check", side_effect=fake_check),
                mock.patch.object(apa7, "_run_install", self.runner),
                mock.patch.object(rutas, "package_manager",
                                  return_value=manager("apt") if manager_found else None),
                mock.patch.object(rutas, "skill_root", return_value=self.tmp),
                mock.patch.object(rutas, "is_windows", return_value=windows),
                mock.patch.object(rutas, "is_macos", return_value=False),
                mock.patch.object(rutas, "is_linux", return_value=not windows),
                mock.patch.object(
                    rutas, "node_path",
                    side_effect=lambda: "/usr/bin/node"
                    if self.state["node"] else None),
                mock.patch.object(rutas, "npm_path", return_value="/usr/bin/npm"),
                mock.patch.object(rutas, "node_dir", return_value=self.tmp / "node"),
                mock.patch.object(rutas, "venv_python", return_value=None),
                mock.patch.object(rutas, "python_works", return_value=python),
                mock.patch.object(
                    rutas, "python_path",
                    side_effect=lambda: "/usr/bin/python3" if python else None),
                mock.patch.object(
                    rutas, "soffice_path",
                    side_effect=lambda: "/usr/bin/soffice"
                    if self.state["soffice"] else None),
                mock.patch.object(rutas, "remove_tree", return_value=True),
                mock.patch.object(rutas, "run", side_effect=fake_run),
                mock.patch.object(sys, "stdin", stdin),
            ):
                stack.enter_context(patcher)
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = apa7.cmd_install(args)
        return code, out.getvalue(), err.getvalue()

    @property
    def commands(self):
        """Every command as a flat string list, for substring assertions."""
        return [" ".join(str(argument) for argument in call[0][0])
                for call in self.runner.call_args_list]


class TestCompleteEnvironment(InstallHarness):
    def test_a_complete_environment_installs_nothing(self):
        code, _, err = self.run_install(complete=True)
        self.assertEqual(code, 0)
        self.assertEqual(self.commands, [])
        self.assertIn("Nothing to install", err)

    def test_only_a_package_manager_must_not_make_a_complete_machine_fail(self):
        # The exact regression from the port: winget absent on a Linux box with
        # a working LibreOffice used to be reported as MISSING.
        code, _, _ = self.run_install(complete=True, manager_found=False)
        self.assertEqual(code, 0)


class TestDryRun(InstallHarness):
    def test_dry_run_runs_no_command_at_all(self):
        code, _, err = self.run_install(dry_run=True)
        self.assertEqual(self.commands, [], "dry-run executed %r" % (self.commands,))
        self.assertIn("[dry-run] would:", err)

    def test_dry_run_creates_no_virtual_environment(self):
        self.run_install(dry_run=True)
        self.assertFalse((self.tmp / ".venv").exists())

    def test_dry_run_prints_every_installing_step(self):
        # python=False so the Python step is part of the plan; a machine that
        # already has Python must not be told it would install it.
        _, _, err = self.run_install(dry_run=True, python=False)
        for tool in ("Node.js", "docx", "Python 3.12", "LibreOffice"):
            self.assertIn(tool, err)

    def test_dry_run_leaves_a_present_tool_out_of_the_plan(self):
        _, _, err = self.run_install(dry_run=True, complete=True)
        self.assertNotIn("[dry-run]", err)

    def test_dry_run_still_reports_the_environment_through_stdout(self):
        _, out, _ = self.run_install(dry_run=True)
        self.assertNotIn("[dry-run]", out)


class TestConfirmation(InstallHarness):
    def test_a_non_interactive_run_refuses_instead_of_hanging(self):
        # Nobody is there to answer y/n; blocking forever is the bug.
        code, _, err = self.run_install(stdin_tty=False)
        self.assertEqual(code, 1)
        self.assertEqual(self.commands, [])
        self.assertIn("needs confirmation", err)
        self.assertIn("--yes", err)

    def test_yes_installs_without_asking(self):
        code, _, _ = self.run_install(yes=True)
        self.assertEqual(code, 0)
        self.assertTrue(self.commands)

    def test_a_yes_run_never_shells_out_to_bare_sudo(self):
        self.run_install(yes=True)
        for command in self.commands:
            if "sudo" in command.split():
                self.assertIn("-n", command, "a bare sudo would prompt: %r" % command)

    def test_an_interactive_yes_answer_is_accepted(self):
        # The prompt path, not just --yes.
        stdin = io.StringIO("y\n")

        def run_with_answer():
            args = argparse.Namespace(only=None, yes=False, dry_run=False)
            self.windows = True
            stdin.isatty = lambda: True
            out, err = io.StringIO(), io.StringIO()
            with mock.patch.object(sys, "stdin", stdin), \
                    mock.patch.object(sys, "stdout", out), \
                    mock.patch.object(sys, "stderr", err), \
                    mock.patch.object(apa7, "cmd_check",
                                      side_effect=[1, 0]), \
                    mock.patch.object(apa7, "_run_install", self.runner), \
                    mock.patch.object(rutas, "package_manager", return_value=manager("brew")), \
                    mock.patch.object(rutas, "skill_root", return_value=self.tmp), \
                    mock.patch.object(rutas, "is_windows", return_value=False), \
                    mock.patch.object(rutas, "is_macos", return_value=True), \
                    mock.patch.object(rutas, "is_linux", return_value=False), \
                    mock.patch.object(rutas, "node_path", return_value=None), \
                    mock.patch.object(rutas, "npm_path", return_value="/opt/npm"), \
                    mock.patch.object(rutas, "node_dir", return_value=self.tmp / "node"), \
                    mock.patch.object(rutas, "venv_python", return_value=None), \
                    mock.patch.object(rutas, "python_works", return_value=True), \
                    mock.patch.object(rutas, "python_path", return_value="/usr/bin/python3"), \
                    mock.patch.object(rutas, "soffice_path", return_value=None), \
                    mock.patch.object(rutas, "run",
                                      return_value=rutas.NativeResult([""], "", 1)), \
                    contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                return apa7.cmd_install(args), err.getvalue()

        code, err = run_with_answer()
        self.assertEqual(code, 0)
        self.assertIn("Proceed?", err)
        # macOS: brew is per-user, so no sudo anywhere.
        self.assertTrue(any("brew" in command for command in self.commands))
        self.assertFalse(any("sudo" in command for command in self.commands))


class TestOnlyFilter(InstallHarness):
    def test_an_unknown_tool_is_rejected_with_the_valid_list(self):
        code, out, err = self.run_install(only=["libre"], yes=True)
        self.assertEqual(code, 2)
        self.assertEqual(out, "")
        self.assertIn("unknown --only", err)
        self.assertIn("libreoffice", err)

    def test_excluding_python_refuses_to_continue_without_it(self):
        code, _, err = self.run_install(only=["node"], yes=True, python=False)
        self.assertEqual(code, 1)
        self.assertIn("excludes", err)

    def test_only_node_leaves_the_other_tools_alone(self):
        _, _, _ = self.run_install(only=["node"], yes=True)
        installed = self.commands
        self.assertTrue(any("nodejs" in command for command in installed))
        self.assertFalse(any("libreoffice" in command.lower() for command in installed))
        # The apt node plan lists npm as a package, so the discriminator for
        # "docx was not installed" is the pinned library, not the word npm.
        self.assertFalse(any("docx@" in command for command in installed))

    def test_a_repeated_only_flag_accumulates(self):
        _, _, _ = self.run_install(only=["node", "libreoffice"], yes=True)
        installed = self.commands
        self.assertTrue(any("nodejs" in command for command in installed))
        self.assertTrue(any("libreoffice" in command.lower() for command in installed))
        self.assertFalse(any("docx@" in command for command in installed))
        self.assertFalse(any("-m venv" in command for command in installed))


class TestNoPackageManager(InstallHarness):
    def test_it_exits_nonzero_with_manual_steps_and_nothing_on_stdout(self):
        code, out, err = self.run_install(manager_found=False)
        self.assertEqual(code, 1)
        self.assertEqual(out, "")
        self.assertIn("no supported package manager", err)
        self.assertIn("Then run:", err)
        self.assertEqual(self.commands, [])

    def test_it_prints_one_copy_pasteable_command_per_tool(self):
        _, _, err = self.run_install(manager_found=False)
        for tool in instalador.TOOLS:
            self.assertIn(tool, instalador.MANUAL)
        # Every manual line is indented under the explanation, so it can be
        # copied straight out of the terminal.
        commands = [line.strip() for line in err.splitlines()
                    if line.startswith("  ") and line.strip()]
        self.assertGreaterEqual(len(commands), len(instalador.TOOLS))


class TestStreamContract(InstallHarness):
    def test_progress_never_reaches_stdout(self):
        _, out, err = self.run_install(yes=True)
        for noise in ("would:", "Installing", "Confirmed by", "Running:"):
            self.assertNotIn(noise, out)
        self.assertIn("Installing", err)

    def test_the_final_check_is_what_stdout_carries(self):
        def check_output(_args):
            sys.stdout.write("RESULT: OK\n")
            return 0

        out, err = io.StringIO(), io.StringIO()
        self.windows = True
        with mock.patch.object(apa7, "cmd_check", side_effect=check_output):
            with mock.patch.object(apa7, "_run_install", self.runner), \
                    mock.patch.object(rutas, "package_manager", return_value=manager("brew")), \
                    mock.patch.object(rutas, "skill_root", return_value=self.tmp), \
                    mock.patch.object(rutas, "node_path", return_value=None), \
                    mock.patch.object(rutas, "npm_path", return_value="/opt/npm"), \
                    mock.patch.object(rutas, "node_dir", return_value=self.tmp / "node"), \
                    mock.patch.object(rutas, "venv_python", return_value=None), \
                    mock.patch.object(rutas, "python_works", return_value=True), \
                    mock.patch.object(rutas, "python_path", return_value="/usr/bin/python3"), \
                    mock.patch.object(rutas, "soffice_path", return_value="/usr/bin/soffice"), \
                    mock.patch.object(rutas, "run",
                                      return_value=rutas.NativeResult([rutas.DEPS["pymupdf"]], "", 0)), \
                    contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = apa7.cmd_install(argparse.Namespace(only=None, yes=True, dry_run=False))
        self.assertEqual(code, 0)
        self.assertEqual(out.getvalue(), "RESULT: OK\n")


class TestRunInstall(unittest.TestCase):
    """The real _run_install, against real subprocesses.

    It is the only part of install that touches stdout at all, and it touches
    it only by accident: npm and apt both print progress. So its logging has to
    be verified for real rather than through a mock.
    """

    def test_it_logs_the_command_and_hides_stdout(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = apa7._run_install(
                [sys.executable, "-c", "print('npm chatter')"])
        self.assertEqual(code, 0)
        self.assertIn("Running:", err.getvalue())
        self.assertEqual(out.getvalue(), "", "a subprocess spoke to stdout")

    def test_it_forwards_the_failure_output_to_stderr(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = apa7._run_install(
                [sys.executable, "-c",
                 "import sys; sys.stderr.write('E: cannot find package\\n');"
                 " sys.exit(2)"])
        self.assertEqual(code, 2)
        self.assertIn("E: cannot find package", err.getvalue())
        self.assertEqual(out.getvalue(), "")

    def test_a_missing_binary_is_a_failure_not_a_traceback(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = apa7._run_install(["definitely-not-a-real-binary-xyz"])
        self.assertNotEqual(code, 0)
        self.assertEqual(out.getvalue(), "")

    def test_a_timeout_is_a_failure_not_a_traceback(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                code = apa7._run_install([sys.executable, "-c", "import time;"
                                                            "time.sleep(30)"],
                                         timeout=1)
            except subprocess.TimeoutExpired:
                self.fail("a hung installer must not raise at the caller")
        self.assertNotEqual(code, 0)
        self.assertEqual(out.getvalue(), "")


if __name__ == "__main__":
    unittest.main()