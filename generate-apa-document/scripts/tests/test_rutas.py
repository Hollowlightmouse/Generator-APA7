"""Tests for scripts/lib/rutas.py.

Standard library only, on purpose: `apa7.py check` has to be runnable before
anything is installed, and a test suite that needs pytest is one more thing to
install first.

Run them with:
    python -m unittest discover -s scripts/tests -v
"""
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import lib.rutas as rutas  # noqa: E402

_ABSENT = object()


@contextmanager
def env(**overrides):
    """Set environment variables for the block. A value of None REMOVES one.

    mock.patch.dict cannot do that: passing None raises TypeError, so a variable
    that has to be absent during the test has to be unset explicitly.
    """
    saved = {}
    for key, value in overrides.items():
        saved[key] = os.environ.get(key, _ABSENT)
        if value is None:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value
    try:
        yield
    finally:
        for key, value in saved.items():
            if value is _ABSENT:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


class PlatformMixin:
    """Pins the three platform predicates so one test machine can test all three."""

    @contextmanager
    def as_platform(self, system):
        with mock.patch.object(rutas, "is_windows", return_value=system == "Windows"), \
                mock.patch.object(rutas, "is_macos", return_value=system == "Darwin"), \
                mock.patch.object(rutas, "is_linux", return_value=system == "Linux"):
            yield

    def as_windows(self):
        return self.as_platform("Windows")

    def as_macos(self):
        return self.as_platform("Darwin")

    def as_linux(self):
        return self.as_platform("Linux")


class TestSofficeCandidates(PlatformMixin, unittest.TestCase):
    """Every OS gets its own default locations.

    rutas.ps1 branched on macOS and treated everything else as Windows, so on
    Linux it searched Program Files and never the Unix locations.
    """

    def test_windows_looks_in_program_files(self):
        with env(APA7_SOFFICE=None, ProgramFiles=r"C:\PF",
                 **{"ProgramFiles(x86)": r"C:\PF86", "LOCALAPPDATA": r"C:\LAD"}), \
                self.as_windows(), \
                mock.patch.object(rutas.shutil, "which", return_value=None):
            found = rutas.soffice_candidates()
        self.assertIn(str(Path(r"C:\PF") / "LibreOffice" / "program" / "soffice.exe"), found)
        self.assertIn(str(Path(r"C:\LAD") / "Programs" / "LibreOffice" / "program" / "soffice.exe"), found)
        self.assertNotIn("/usr/bin/soffice", found)

    def test_macos_looks_in_the_app_bundle_and_both_brew_prefixes(self):
        with env(APA7_SOFFICE=None), \
                self.as_macos(), \
                mock.patch.object(rutas.shutil, "which", return_value=None):
            found = rutas.soffice_candidates()
        self.assertIn("/Applications/LibreOffice.app/Contents/MacOS/soffice", found)
        self.assertIn("/opt/homebrew/bin/soffice", found)   # Apple Silicon
        self.assertIn("/usr/local/bin/soffice", found)      # Intel

    def test_linux_looks_in_unix_locations_and_snap_and_flatpak(self):
        with env(APA7_SOFFICE=None), \
                self.as_linux(), \
                mock.patch.object(rutas.shutil, "which", return_value=None):
            found = rutas.soffice_candidates()
        self.assertIn("/usr/bin/soffice", found)
        self.assertIn("/snap/bin/libreoffice.current/usr/bin/soffice", found)
        self.assertIn("/var/lib/flatpak/exports/bin/soffice", found)
        self.assertFalse([p for p in found if "Program Files" in p])

    def test_environment_override_comes_first(self):
        with env(APA7_SOFFICE="/custom/soffice"):
            self.assertEqual(rutas.soffice_candidates()[0], "/custom/soffice")

    def test_path_comes_before_the_hardcoded_defaults(self):
        with env(APA7_SOFFICE=None), \
                self.as_linux(), \
                mock.patch.object(rutas.shutil, "which", return_value="/from/path/soffice"):
            self.assertIn("/from/path/soffice", rutas.soffice_candidates())


class TestSofficeConsole(unittest.TestCase):
    """On Windows the .com launcher is mandatory, on Unix there is only one."""

    def test_windows_prefers_the_com_sibling(self):
        with tempfile.TemporaryDirectory() as tmp:
            launcher = Path(tmp) / "soffice.exe"
            launcher.write_bytes(b"")
            console = Path(tmp) / "soffice.com"
            console.write_bytes(b"")
            with mock.patch.object(rutas, "soffice_path", return_value=str(launcher)):
                self.assertEqual(rutas.soffice_console(), str(console))

    def test_windows_falls_back_to_the_exe_without_a_com(self):
        with tempfile.TemporaryDirectory() as tmp:
            launcher = Path(tmp) / "soffice.exe"
            launcher.write_bytes(b"")
            with mock.patch.object(rutas, "soffice_path", return_value=str(launcher)):
                self.assertEqual(rutas.soffice_console(), str(launcher))

    def test_returns_none_when_nothing_is_installed(self):
        with mock.patch.object(rutas, "soffice_path", return_value=None):
            self.assertIsNone(rutas.soffice_console())


class _FakeProcess:
    def __init__(self, returncode=0, hang=False):
        self._returncode = returncode
        self._hang = hang
        self.pid = 4242

    def wait(self, timeout=None):
        if self._hang:
            raise subprocess.TimeoutExpired("soffice", timeout)
        return self._returncode


class TestRunSoffice(unittest.TestCase):
    """The plumbing that keeps a wedged LibreOffice from hanging the pipeline.

    Popen is faked so the test is deterministic and needs no LibreOffice.
    """

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.fake = Path(self.tmp.name) / "soffice"
        self.fake.write_bytes(b"")

    def test_arguments_carry_headless_and_an_isolated_profile(self):
        with mock.patch.object(rutas, "soffice_console", return_value=str(self.fake)), \
                mock.patch.object(rutas.subprocess, "Popen", return_value=_FakeProcess()) as popen:
            result = rutas.run_soffice(["--version"], log_dir=self.tmp.name)
        argv = popen.call_args[0][0]
        for flag in ("--headless", "--norestore", "--nolockcheck", "--nofirststartwizard"):
            self.assertIn(flag, argv)
        self.assertIn("--version", argv)
        # The isolated profile is what keeps an already-open LibreOffice from
        # making the conversion hang silently, so it must always be there.
        profile = [a for a in argv if a.startswith("-env:UserInstallation=")]
        self.assertEqual(len(profile), 1)
        self.assertTrue(profile[0].startswith("-env:UserInstallation=file:///"), profile[0])
        self.assertTrue(profile[0].endswith("/lo_profile"), profile[0])
        self.assertEqual(Path(result.profile_path).name, "lo_profile")

    def test_a_wedged_libreoffice_is_killed_and_reported_as_124(self):
        with mock.patch.object(rutas, "soffice_console", return_value=str(self.fake)), \
                mock.patch.object(rutas.subprocess, "Popen", return_value=_FakeProcess(hang=True)), \
                mock.patch.object(rutas, "kill_process_tree") as killer:
            result = rutas.run_soffice(["--convert-to", "pdf"], log_dir=self.tmp.name)
        self.assertTrue(result.timed_out)
        self.assertEqual(result.exit_code, 124)
        killer.assert_called_once_with(4242)

    def test_a_failure_to_launch_is_127_and_not_a_crash(self):
        with mock.patch.object(rutas, "soffice_console", return_value=str(self.fake)), \
                mock.patch.object(rutas.subprocess, "Popen", side_effect=OSError("boom")):
            result = rutas.run_soffice(["--version"], log_dir=self.tmp.name)
        self.assertEqual(result.exit_code, 127)
        self.assertFalse(result.timed_out)
        self.assertIn("boom", result.launch_error)

    def test_output_is_written_to_files_not_pipes(self):
        with mock.patch.object(rutas, "soffice_console", return_value=str(self.fake)), \
                mock.patch.object(rutas.subprocess, "Popen", return_value=_FakeProcess()) as popen:
            result = rutas.run_soffice(["--version"], log_dir=self.tmp.name)
        kwargs = popen.call_args[1]
        self.assertNotEqual(kwargs.get("stdout"), subprocess.PIPE)
        self.assertEqual(kwargs.get("stdin"), subprocess.DEVNULL)
        self.assertEqual(Path(result.stdout_path).name, "lo.out.log")
        self.assertEqual(Path(result.stderr_path).name, "lo.err.log")

    def test_a_directory_we_own_is_removed_again(self):
        with mock.patch.object(rutas, "soffice_console", return_value=str(self.fake)), \
                mock.patch.object(rutas.subprocess, "Popen", return_value=_FakeProcess()):
            result = rutas.run_soffice(["--version"])
        self.assertFalse(Path(result.log_dir).exists())

    def test_a_caller_directory_is_left_in_place(self):
        with mock.patch.object(rutas, "soffice_console", return_value=str(self.fake)), \
                mock.patch.object(rutas.subprocess, "Popen", return_value=_FakeProcess()):
            result = rutas.run_soffice(["--version"], log_dir=self.tmp.name)
        self.assertTrue(Path(result.log_dir).exists())

    def test_missing_libreoffice_raises_a_named_error(self):
        with mock.patch.object(rutas, "soffice_console", return_value=None):
            with self.assertRaises(rutas.SofficeNotFound):
                rutas.run_soffice(["--version"])


class TestKillProcessTree(PlatformMixin, unittest.TestCase):
    def test_windows_delegates_to_taskkill_with_the_whole_tree(self):
        with self.as_windows(), \
                mock.patch.object(rutas.shutil, "which", return_value="taskkill"), \
                mock.patch.object(rutas.subprocess, "run") as runner:
            rutas.kill_process_tree(1234)
        argv = runner.call_args[0][0]
        self.assertIn("/T", argv)
        self.assertIn("/F", argv)
        self.assertIn("1234", argv)

    def test_posix_kills_the_whole_process_group(self):
        # os.killpg and os.getpgid do not exist on Windows, hence create=True.
        with self.as_linux(), \
                mock.patch.object(rutas.os, "getpgid", return_value=999, create=True), \
                mock.patch.object(rutas.os, "killpg", create=True) as killpg:
            rutas.kill_process_tree(1234)
        killpg.assert_called_once()
        self.assertEqual(killpg.call_args[0][0], 999)

    def test_a_missing_pid_is_a_no_op(self):
        with mock.patch.object(rutas.os, "killpg", create=True) as killpg:
            rutas.kill_process_tree(0)
        killpg.assert_not_called()


class TestFilterStderr(unittest.TestCase):
    """Only the known LibreOffice headless noise is discarded."""

    def test_drops_the_known_noise(self):
        noise = "\n".join([
            "Could not find platform independent libraries C:\\Python",
            "Warning: failed to launch javaldx",
            "libpng warning: iCCP: known incorrect sRGB profile",
        ])
        self.assertEqual(rutas.filter_stderr(noise), [])

    def test_keeps_lines_that_deserve_attention(self):
        mixed = "\n".join([
            "Warning: failed to launch javaldx",
            "Error: source file could not be loaded",
        ])
        self.assertEqual(rutas.filter_stderr(mixed), ["Error: source file could not be loaded"])

    def test_empty_input_gives_empty_output(self):
        self.assertEqual(rutas.filter_stderr(""), [])
        self.assertEqual(rutas.filter_stderr(None), [])


class TestSkillRoot(unittest.TestCase):
    """APA7_SKILL_ROOT was documented for a long time and never read."""

    def setUp(self):
        rutas._SKILL_ROOT = None
        self.addCleanup(setattr, rutas, "_SKILL_ROOT", None)

    def test_defaults_to_the_folder_containing_this_file(self):
        with env(APA7_SKILL_ROOT=None):
            root = rutas.skill_root()
        self.assertTrue((root / "SKILL.md").is_file())
        self.assertEqual(root, Path(rutas.__file__).resolve().parents[2])

    def test_a_valid_override_wins(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "SKILL.md").write_text("x", encoding="utf-8")
            with env(APA7_SKILL_ROOT=tmp):
                self.assertEqual(rutas.skill_root(), Path(tmp).resolve())

    def test_an_override_without_skill_md_is_reported_and_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            with env(APA7_SKILL_ROOT=tmp), \
                    mock.patch.object(rutas.sys, "stderr") as stderr:
                root = rutas.skill_root()
            self.assertTrue((root / "SKILL.md").is_file())
            self.assertIn("APA7_SKILL_ROOT", stderr.write.call_args[0][0])


class TestWorkdir(unittest.TestCase):
    def setUp(self):
        rutas._SKILL_ROOT = None
        self.addCleanup(setattr, rutas, "_SKILL_ROOT", None)

    def test_it_is_created_even_though_it_is_only_scratch(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "work"
            with env(APA7_WORKDIR=str(target)):
                found = rutas.workdir()
            self.assertEqual(found, target)
            self.assertTrue(target.is_dir())

    def test_node_dir_defaults_to_the_workdir_and_honours_its_override(self):
        with tempfile.TemporaryDirectory() as tmp:
            with env(APA7_WORKDIR=tmp, APA7_NODEDIR=None):
                self.assertEqual(rutas.node_dir(), Path(tmp))
            with env(APA7_WORKDIR=tmp, APA7_NODEDIR=str(Path(tmp) / "other")):
                self.assertEqual(rutas.node_dir(), Path(tmp) / "other")


class TestVenvCandidates(unittest.TestCase):
    """The Linux bug: rutas.ps1 looked for `venv/Scripts/python.exe` there too."""

    def test_both_layouts_are_probed_on_every_platform(self):
        found = [p.as_posix() for p in rutas._venv_candidates(Path("/skill"))]
        self.assertIn("/skill/venv/bin/python3", found)
        self.assertIn("/skill/.venv/bin/python3", found)
        self.assertIn("/skill/venv/Scripts/python.exe", found)
        self.assertIn("/skill/.venv/Scripts/python.exe", found)


class TestPythonDiscovery(unittest.TestCase):
    def test_a_windows_apps_stub_is_rejected_even_though_it_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            stub = Path(tmp, "WindowsApps", "python.exe")
            stub.parent.mkdir()
            stub.write_bytes(b"")
            self.assertFalse(rutas.python_works(str(stub)))

    def test_a_real_interpreter_is_accepted(self):
        self.assertTrue(rutas.python_works(sys.executable))

    def test_a_path_that_does_not_exist_is_not_a_python(self):
        self.assertFalse(rutas.python_works("/no/existe/python"))
        self.assertFalse(rutas.python_works(""))

    def test_the_override_comes_first(self):
        with env(APA7_PYTHON="/custom/python"):
            self.assertEqual(rutas.python_candidates()[0], "/custom/python")

    def test_venv_python_returns_none_when_there_is_no_venv(self):
        with tempfile.TemporaryDirectory() as tmp:
            rutas._SKILL_ROOT = Path(tmp)
            self.addCleanup(setattr, rutas, "_SKILL_ROOT", None)
            self.assertIsNone(rutas.venv_python())


class TestRun(unittest.TestCase):
    def test_a_timeout_is_124(self):
        self.assertEqual(rutas.run([sys.executable, "-c", "import time; time.sleep(30)"],
                                   timeout=1).exit_code, 124)

    def test_a_missing_binary_is_127_and_not_an_exception(self):
        result = rutas.run(["definitely-not-a-real-binary-xyz"])
        self.assertEqual(result.exit_code, 127)

    def test_first_line_is_the_version_and_stderr_is_merged(self):
        result = rutas.run([sys.executable, "-c",
                            "import sys; print(sys.version.split()[0]); print('noise', file=sys.stderr)"])
        self.assertEqual(result.exit_code, 0)
        self.assertTrue(result.first_line[0].isdigit())
        self.assertIn("noise", result.output)

    def test_the_output_is_never_truncated(self):
        result = rutas.run([sys.executable, "-c",
                            "print('x'); print('y'); print('z')"])
        self.assertEqual([line for line in result.output if line], ["x", "y", "z"])


class TestPackageManager(unittest.TestCase):
    """winget and brew first, then apt/dnf/pacman so Linux is not a dead end."""

    def _which(self, available):
        return lambda name: ("/bin/" + name) if name in available else None

    def test_detection_order_prefers_the_platform_tools(self):
        with mock.patch.object(rutas.shutil, "which", self._which({"winget", "brew", "apt-get"})):
            self.assertEqual(rutas.package_manager().name, "winget")
        with mock.patch.object(rutas.shutil, "which", self._which({"brew", "apt-get"})):
            self.assertEqual(rutas.package_manager().name, "brew")

    def test_a_plain_linux_box_finds_apt(self):
        with mock.patch.object(rutas.shutil, "which", self._which({"apt-get", "dpkg"})):
            found = rutas.package_manager()
            self.assertEqual(found.name, "apt")
            self.assertEqual(found.path, "/bin/apt-get")

    def test_dnf_and_pacman_are_recognised(self):
        with mock.patch.object(rutas.shutil, "which", self._which({"dnf"})):
            self.assertEqual(rutas.package_manager().name, "dnf")
        with mock.patch.object(rutas.shutil, "which", self._which({"pacman"})):
            self.assertEqual(rutas.package_manager().name, "pacman")

    def test_no_manager_is_none_and_not_a_crash(self):
        with mock.patch.object(rutas.shutil, "which", return_value=None):
            self.assertIsNone(rutas.package_manager())


class TestDeps(unittest.TestCase):
    def test_the_pins_are_the_ones_the_docs_promise(self):
        self.assertEqual(rutas.DEPS["docx"], "9.7.1")
        self.assertEqual(rutas.DEPS["pymupdf"], "1.28.2")

    def test_the_python_floor_matches_what_is_installed_and_what_is_required(self):
        self.assertEqual(rutas.MIN_PYTHON, (3, 9))
        self.assertGreaterEqual(sys.version_info[:2], rutas.MIN_PYTHON)


if __name__ == "__main__":
    unittest.main()