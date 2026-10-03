"""Portable path and tool discovery for the generate-apa-document skill.

STANDARD LIBRARY ONLY. This module has to stay importable with nothing
installed, because `apa7.py check` and `apa7.py install` are the scripts that
install everything else. A third-party import here would make the bootstrap
impossible.

Replaces scripts/lib/rutas.ps1. The three PowerShell traps that library had to
work around do not exist in Python and are documented where they mattered, so
they are not reproduced:
  * soffice.exe detaching and leaving the reader waiting forever -> here the
    output goes to real files, so no handle is ever inherited.
  * Start-Process -PassThru returning a Process whose ExitCode is $null unless
    the .Handle is touched first -> Popen.returncode always works.
  * RedirectStandardOutput leaving the streams open in the child forever ->
    there is no such API here.

GOLDEN RULE: no absolute path of any specific machine lives in this file.
Everything is resolved at runtime from the environment, from this file's own
location, from PATH, or from the default install locations of the detected OS.

Supported environment variables (all optional):
    APA7_SKILL_ROOT  skill root
    APA7_SOFFICE     path to the LibreOffice executable
    APA7_PYTHON      path to the Python interpreter that has pymupdf
    APA7_WORKDIR     working directory (where node_modules/docx lives)
    APA7_NODEDIR     directory that contains node_modules/docx
"""
import os
import platform
import shutil
import signal
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Pinned versions of the libraries this skill's own code talks to
# ---------------------------------------------------------------------------
# ONE place for both. They used to be literal strings inside instalar-entorno.ps1
# with nothing ever checking them, so the pin was documentation rather than a
# restriction. The preflight compares what is installed against these values and
# says so in its detail.
#
# pymupdf is pinned because verificar-pdf.py calls its API directly. Above 1.24.3
# the top-level module is `pymupdf` and the `fitz` alias no longer exists.
DEPS = {
    "docx": "9.7.1",
    "pymupdf": "1.28.2",
}

# The oldest interpreter this skill supports. Verified against the parser and
# the verifier: both use only the standard library and no syntax newer than 3.9.
# It is the floor, not the target: install still puts 3.12 in place.
MIN_PYTHON = (3, 9)

_PY_NAMES = ("python3", "python", "python.exe", "python3.exe")
_PY_DIRS = ("Scripts", "bin")

# Windows has no SIGKILL. Resolved once, at import, so that kill_process_tree
# stays callable (and testable) on the machine that does not have it.
_KILL_SIGNAL = getattr(signal, "SIGKILL", signal.SIGTERM)


class SofficeNotFound(RuntimeError):
    """LibreOffice could not be located. Never raised by discovery itself."""


# ---------------------------------------------------------------------------
# Platform detection
# ---------------------------------------------------------------------------
# Three-way, not two-way. rutas.ps1 branched on macOS and treated "everything
# else" as Windows, so on Linux it looked for `venv/Scripts/python.exe` and
# python.exe: the PowerShell version never worked on Linux either.
def is_windows():
    return os.name == "nt"


def is_macos():
    return platform.system() == "Darwin"


def is_linux():
    return platform.system() == "Linux"


# ---------------------------------------------------------------------------
# Skill root
# ---------------------------------------------------------------------------
_SKILL_ROOT = None


def skill_root():
    """Root folder of the skill, i.e. the folder that contains SKILL.md.

    APA7_SKILL_ROOT overrides it. The README and system-requirements.md have
    documented this variable for a long time while rutas.ps1 overwrote it
    without ever reading the environment, so the documented override did not
    work. An override that does not look like a skill root is reported and
    ignored rather than silently accepted, which would break every later path.
    """
    global _SKILL_ROOT
    if _SKILL_ROOT is not None:
        return _SKILL_ROOT

    override = os.environ.get("APA7_SKILL_ROOT")
    if override:
        candidate = Path(override).expanduser()
        if (candidate / "SKILL.md").is_file():
            _SKILL_ROOT = candidate.resolve()
            return _SKILL_ROOT
        sys.stderr.write(
            "WARNING: APA7_SKILL_ROOT=%s has no SKILL.md; ignoring it and using "
            "the folder this script lives in.\n" % override
        )

    _SKILL_ROOT = Path(__file__).resolve().parents[2]
    return _SKILL_ROOT


def skill_script(name):
    """Absolute path of a script inside scripts/."""
    path = skill_root() / "scripts" / name
    if not path.is_file():
        raise FileNotFoundError("Script not found: %s" % path)
    return path


def workdir():
    """Generated working directory. Created even in a dry run.

    An empty working directory is scratch state, not an installation, so a
    dry run must not suppress it: the probes that follow expect it to exist.
    """
    override = os.environ.get("APA7_WORKDIR")
    path = Path(override).expanduser() if override else skill_root() / ".work"
    path.mkdir(parents=True, exist_ok=True)
    return path


def node_dir():
    """Directory that MUST contain node_modules/docx."""
    override = os.environ.get("APA7_NODEDIR")
    return Path(override).expanduser() if override else workdir()


# ---------------------------------------------------------------------------
# Shared utilities
# ---------------------------------------------------------------------------
@dataclass
class NativeResult:
    output: list = field(default_factory=list)
    first_line: str = ""
    exit_code: int = 0
    stderr: str = ""


def run(argv, timeout=60, cwd=None):
    """Run a native command and return its output and exit code.

    TimeoutExpired is reported as 124 and a failure to launch as 127, which are
    the same codes the PowerShell version produced.

    Nothing truncates the output. The old implementation reached for
    `Select-Object -First 1`, which cuts the pipeline early and can leave the
    child hung or with a wrong exit code, producing false missing-tool results.
    """
    argv = [str(a) for a in argv]
    try:
        completed = subprocess.run(
            argv,
            capture_output=True,
            text=True,
            errors="replace",
            timeout=timeout,
            cwd=str(cwd) if cwd else None,
            stdin=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired:
        return NativeResult([], "", 124)
    except OSError as exc:
        return NativeResult([str(exc)], "", 127)

    merged = "%s\n%s" % (completed.stdout or "", completed.stderr or "")
    lines = [line for line in merged.replace("\r\n", "\n").split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    first = lines[0].strip() if lines else ""
    return NativeResult(lines, first, completed.returncode, completed.stderr or "")


def log(message):
    """Timestamped line, on stderr.

    stderr, not stdout, because stdout carries the machine-readable contract
    that SKILL.md parses. A human running the skill by hand still sees it.
    """
    sys.stderr.write("[%s] %s\n" % (datetime.now().strftime("%H:%M:%S"), message))
    sys.stderr.flush()


def _remove_tree(path):
    if path.exists():
        shutil.rmtree(str(path), ignore_errors=True)


def kill_process_tree(pid):
    """Kill a process and its children (soffice -> soffice.bin).

    The tree, never just the launcher: the real work happens in the soffice.bin
    child, which survives a flat kill and then holds the output files and makes
    the NEXT conversion fail.
    """
    if not pid:
        return
    if is_windows():
        taskkill = shutil.which("taskkill")
        if taskkill:
            try:
                subprocess.run(
                    [taskkill, "/PID", str(pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=30,
                )
                return
            except (OSError, subprocess.SubprocessError):
                pass
        try:
            os.kill(pid, signal.SIGTERM)
        except OSError:
            pass
        return

    # The child was started with start_new_session=True, so it leads its own
    # process group and the whole group dies with one call.
    try:
        os.killpg(os.getpgid(pid), _KILL_SIGNAL)
    except (OSError, AttributeError):
        try:
            os.kill(pid, _KILL_SIGNAL)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# LibreOffice: environment -> PATH -> default locations per OS
# ---------------------------------------------------------------------------
def soffice_candidates():
    """Every path that might hold the LibreOffice launcher, best first."""
    out = []

    override = os.environ.get("APA7_SOFFICE")
    if override:
        out.append(override)

    for name in ("soffice", "soffice.com"):
        found = shutil.which(name)
        if found:
            out.append(found)

    if is_macos():
        out.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        # Homebrew's prefix differs by architecture; shutil.which("brew") in
        # the installer resolves the real one, so these are only a last resort.
        out.append("/opt/homebrew/bin/soffice")   # Apple Silicon
        out.append("/usr/local/bin/soffice")      # Intel
    elif is_windows():
        program_files = os.environ.get("ProgramFiles")
        program_files_x86 = os.environ.get("ProgramFiles(x86)")
        local_app_data = os.environ.get("LOCALAPPDATA")
        if program_files:
            out.append(str(Path(program_files) / "LibreOffice" / "program" / "soffice.exe"))
        if program_files_x86:
            out.append(str(Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe"))
        if local_app_data:
            out.append(str(Path(local_app_data) / "Programs" / "LibreOffice" / "program" / "soffice.exe"))
    else:
        out.append("/usr/bin/soffice")
        out.append("/usr/lib/libreoffice/program/soffice")
        # Distributions that ship LibreOffice through Snap or Flatpak keep the
        # launcher out of the usual places, so a successful install can still
        # look like a missing tool without these.
        out.append("/snap/bin/libreoffice.current/usr/bin/soffice")
        out.append("/var/lib/flatpak/exports/bin/soffice")

    return out


def soffice_path():
    """Path of the LibreOffice executable found, or None.

    It does not raise: the preflight decides whether a missing engine is fatal.
    """
    for candidate in soffice_candidates():
        if candidate and Path(candidate).exists():
            return candidate
    return None


def soffice_console():
    """On Windows, the CONSOLE launcher (soffice.com) if it exists.

    Using soffice.exe instead makes the conversion unreliable: it is a launcher
    that returns control before the work is done, so the caller can report
    success while nothing was written. On macOS and Linux this is the same path
    as soffice_path.
    """
    found = soffice_path()
    if not found:
        return None
    path = Path(found)
    if path.suffix.lower() == ".exe":
        console = path.with_suffix(".com")
        if console.exists():
            return str(console)
    return found


@dataclass
class SofficeResult:
    exit_code: int = 0
    stdout: str = ""
    stderr: str = ""
    timed_out: bool = False
    launch_error: str = None
    log_dir: str = None
    stdout_path: str = None
    stderr_path: str = None
    profile_path: str = None


def run_soffice(arguments=(), timeout=300, log_dir=None):
    """Run LibreOffice headless with a REAL time bound.

    Returns a SofficeResult. Exit code 124 means "timed out" and 127 means
    "could not be launched"; both are conventions the PowerShell version already
    used and the preflight and the exporter already interpret.

    An ISOLATED user profile is passed on every run. If the real profile were
    used, an already-open LibreOffice makes the conversion hang silently.

    Output goes to files, not to pipes. A child that inherits the caller's
    redirected handle can keep LibreOffice alive after we are done, which is
    what used to leave the pipeline hanging; with files there is nothing to
    inherit and the contents are complete the moment the process ends.

    On timeout the whole TREE is killed, not just the launcher.
    """
    binary = soffice_console()
    if not binary:
        raise SofficeNotFound(
            "LibreOffice not found. Set APA7_SOFFICE or run: python scripts/apa7.py install"
        )

    own_dir = log_dir is None
    if own_dir:
        # Internal scratch, NOT an installation: created and deleted with plain
        # filesystem calls so a dry run never suppresses it. A missing directory
        # means LibreOffice gets a -env:UserInstallation pointing at nothing,
        # prints nothing on --version and the preflight wrongly reports
        # LibreOffice as MISSING with exit code 0 and an empty version.
        run_dir = Path(tempfile.gettempdir()) / ("apa7-lo-%d" % os.getpid())
    else:
        run_dir = Path(log_dir).expanduser().resolve()
    run_dir.mkdir(parents=True, exist_ok=True)

    # FIXED names: a run overwrites the previous one instead of leaving a new
    # pair of files behind. Leftovers from a crashed run go first.
    out_file = run_dir / "lo.out.log"
    err_file = run_dir / "lo.err.log"
    profile = run_dir / "lo_profile"
    for stale in (out_file, err_file):
        try:
            stale.unlink()
        except OSError:
            pass
    _remove_tree(profile)

    full_args = [
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nofirststartwizard",
        "-env:UserInstallation=" + profile.as_uri(),
    ] + [str(a) for a in arguments]

    exit_code = 0
    timed_out = False
    launch_error = None

    with open(str(out_file), "w", encoding="utf-8", errors="replace") as out_handle, \
            open(str(err_file), "w", encoding="utf-8", errors="replace") as err_handle:
        try:
            process = subprocess.Popen(
                [binary] + full_args,
                stdout=out_handle,
                stderr=err_handle,
                stdin=subprocess.DEVNULL,
                start_new_session=(not is_windows()),
            )
        except OSError as exc:
            launch_error = str(exc)
            exit_code = 127
        else:
            try:
                exit_code = process.wait(timeout=max(1, int(timeout)))
            except subprocess.TimeoutExpired:
                timed_out = True
                kill_process_tree(process.pid)
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    pass
                exit_code = 124

    def _read(path):
        try:
            return path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            return ""

    result = SofficeResult(
        exit_code=exit_code,
        stdout=_read(out_file),
        stderr=_read(err_file),
        timed_out=timed_out,
        launch_error=launch_error,
        log_dir=str(run_dir),
        stdout_path=str(out_file),
        stderr_path=str(err_file),
        profile_path=str(profile),
    )

    # Everything we needed is already read. A directory we own is removed here
    # so a dry run does not leak apa7-lo-* folders into TEMP.
    if own_dir:
        _remove_tree(run_dir)

    return result


def filter_stderr(stderr):
    """Keep only the stderr lines that deserve attention.

    Discarded, because they are known LibreOffice headless noise and none of them
    affect the conversion:
      * "Could not find platform independent libraries <prefix>": the process
        inherited PYTHONHOME or PYTHONPATH.
      * "Warning: failed to launch javaldx".
      * libpng warnings.

    The PowerShell version also filtered its own wrapper noise (CategoryInfo,
    FullyQualifiedErrorId, ".ps1: N Car"). Those lines were never process output,
    so there is nothing to filter here.
    """
    if not stderr:
        return []
    noise = (
        "Could not find platform independent libraries",
        "Warning: failed to launch javaldx",
        "libpng warning",
    )
    real = []
    for line in stderr.replace("\r\n", "\n").split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        if any(marker in stripped for marker in noise):
            continue
        real.append(stripped)
    return real


# ---------------------------------------------------------------------------
# Python with pymupdf: environment -> own venv -> PATH -> usual locations
# ---------------------------------------------------------------------------
def _venv_candidates(root):
    """Every plausible interpreter inside a virtual environment at `root`.

    Both layouts are probed on every platform instead of branching per OS. That
    is what rutas.ps1 got wrong on Linux: it branched on macOS and then assumed
    Windows everywhere else, so `venv/Scripts/python.exe` was the only candidate
    on Linux and a Unix venv was never found.
    """
    for sub in ("venv", ".venv"):
        for directory in _PY_DIRS:
            for name in _PY_NAMES:
                yield root / sub / directory / name


def venv_python():
    """Interpreter inside the skill's own virtual environment, or None.

    It deliberately does NOT check that the interpreter RUNS: the preflight has
    to tell "no venv" apart from "venv present but broken", because the fix is
    different in each case (install Python vs recreate the venv).
    """
    for candidate in _venv_candidates(skill_root()):
        if candidate.is_file():
            return str(candidate)
    return None


def python_candidates():
    out = []

    # 1) explicit override, always wins
    override = os.environ.get("APA7_PYTHON")
    if override:
        out.append(override)

    # 2) the skill's own venv, BEFORE any interpreter on PATH. The venv carries
    #    the pinned pymupdf, so it has to win over a system-wide one that may
    #    hold a different version. Order in this list decides which one is used.
    own = venv_python()
    if own:
        out.append(own)
    out.extend(str(c) for c in _venv_candidates(skill_root()))

    # 3) interpreters on PATH
    for name in ("python3", "python"):
        found = shutil.which(name)
        if found:
            out.append(found)

    if is_windows():
        # Typical install locations when Python is not on PATH.
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data and Path(local_app_data).is_dir():
            out.extend(str(p) for p in Path(local_app_data).glob("Programs/Python/Python*/python.exe"))
        program_files = os.environ.get("ProgramFiles")
        if program_files and Path(program_files).is_dir():
            out.extend(str(p) for p in Path(program_files).glob("Python*/python.exe"))
        # The py launcher goes LAST: it still works when a Microsoft Store stub
        # shadows python.exe on PATH.
        launcher = shutil.which("py")
        if launcher:
            out.append(launcher)

    return out


def python_works(executable):
    """True only if the path is a WORKING Python interpreter.

    The Microsoft Store stub lives under WindowsApps and exists but is not a
    real interpreter, so it is rejected by path and, as belt and suspenders, by
    actually running it.
    """
    if not executable:
        return False
    path = Path(executable)
    if "WindowsApps" in path.parts:
        return False
    if not path.is_file():
        return False
    result = run([str(path), "-c", "import sys; print(sys.version.split()[0])"], timeout=20)
    if result.exit_code != 0:
        return False
    return bool(result.first_line.strip())


def python_path():
    """First candidate that actually runs, or None.

    Interpreters that exist but are broken (Microsoft Store stubs, dangling
    venvs) are skipped. Whether pymupdf is installed is checked afterwards by the
    preflight with the interpreter returned here.
    """
    for candidate in python_candidates():
        if python_works(candidate):
            return candidate
    return None


# ---------------------------------------------------------------------------
# Node and npm
# ---------------------------------------------------------------------------
def node_path():
    """Path to the Node executable, or None."""
    return shutil.which("node") or shutil.which("node.exe")


def npm_path():
    """Path to npm, or None.

    On Windows npm.cmd is preferred: it is a native executable, while npm.ps1 is
    a PowerShell wrapper that this Python pipeline has no reason to involve.
    """
    names = ("npm.cmd", "npm", "npm.exe") if is_windows() else ("npm", "npm.cmd")
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


# ---------------------------------------------------------------------------
# Package managers
# ---------------------------------------------------------------------------
@dataclass
class PackageManager:
    name: str
    path: str


# winget and brew come first because they are what the Windows and macOS
# instructions name. apt/dnf/pacman follow so that a Linux box without Homebrew
# still gets an automatic installer instead of a dead end.
_MANAGERS = (
    ("winget", "winget"),
    ("brew", "brew"),
    ("apt", "apt-get"),
    ("dnf", "dnf"),
    ("pacman", "pacman"),
)


def package_manager():
    """First available system package manager, or None."""
    for name, executable in _MANAGERS:
        found = shutil.which(executable)
        if found:
            return PackageManager(name, found)
    return None