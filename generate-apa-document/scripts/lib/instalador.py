"""instalador.py - what has to be run to install each tool, per package manager.

Kept apart from apa7.py on purpose: this file is pure data and pure argv
building, with no prompting, no logging and no side effects. That makes the
part of the installer that is easy to get wrong (which command installs what,
and which of them need root) testable without installing anything.

Every plan is idempotent at the package-manager level: the caller only reaches
here for a tool that `check` already reported as MISSING.
"""

import os

# A tool the user has to provide by hand, with the exact command per platform.
# Printed when there is no package manager to automate it with, or when the
# automatic attempt failed. Never a guess: a wrong hint is worse than none.
MANUAL = {
    "node": {
        "windows": "winget install OpenJS.NodeJS.LTS",
        "macos": "brew install node",
        "linux": "sudo apt install nodejs npm",
    },
    "python": {
        "windows": "winget install Python.Python.3.12",
        "macos": "brew install python@3.12",
        "linux": "sudo apt install python3 python3-venv python3-pip",
    },
    "libreoffice": {
        "windows": "winget install TheDocumentFoundation.LibreOffice",
        "macos": "brew install --cask libreoffice",
        "linux": "sudo apt install libreoffice-writer",
    },
    "docx": {
        "windows": "npm install docx@9.7.1 --no-save",
        "macos": "npm install docx@9.7.1 --no-save",
        "linux": "npm install docx@9.7.1 --no-save",
    },
    "pymupdf": {
        "windows": "python -m venv .venv && .venv/Scripts/python -m pip install pymupdf==1.28.2",
        "macos": "python3 -m venv .venv && .venv/bin/python3 -m pip install pymupdf==1.28.2",
        "linux": "python3 -m venv .venv && .venv/bin/python3 -m pip install pymupdf==1.28.2",
    },
}

# Sub-commands only: the manager path and the elevation prefix are added by
# build(), so the same table serves every machine.
#
# linux rows use apt deliberately. dnf and pacman machines are usually not
# Debian derivatives, and the documented fallback is a copy-pasteable command,
# not a per-distro table that would be wrong half the time.
PLANS = {
    "node": {
        "winget": ["install", "--id", "OpenJS.NodeJS.LTS", "-e",
                   "--accept-package-agreements", "--accept-source-agreements",
                   "--silent", "--disable-interactivity"],
        "brew": ["install", "node"],
        "apt": ["install", "-y", "nodejs", "npm"],
        "dnf": ["install", "-y", "nodejs", "npm"],
        "pacman": ["-S", "--noconfirm", "nodejs", "npm"],
    },
    "python": {
        "winget": ["install", "--id", "Python.Python.3.12", "-e",
                   "--accept-package-agreements", "--accept-source-agreements",
                   "--silent", "--disable-interactivity"],
        "brew": ["install", "python@3.12"],
        # python3-venv is not optional: without it `python3 -m venv` fails and
        # the pinned pymupdf never gets installed.
        "apt": ["install", "-y", "python3", "python3-venv", "python3-pip"],
        "dnf": ["install", "-y", "python3", "python3-pip"],
        "pacman": ["-S", "--noconfirm", "python", "python-pip"],
    },
    "libreoffice": {
        "winget": ["install", "--id", "TheDocumentFoundation.LibreOffice", "-e",
                   "--accept-package-agreements", "--accept-source-agreements",
                   "--silent", "--disable-interactivity"],
        "brew": ["install", "--cask", "libreoffice"],
        # libreoffice-writer only: it is the Writer engine this skill needs, and
        # it is a fraction of the size of the full suite.
        "apt": ["install", "-y", "libreoffice-writer"],
        "dnf": ["install", "-y", "libreoffice-writer"],
        "pacman": ["-S", "--noconfirm", "libreoffice-fresh"],
    },
}

# These install into system directories, so they need root. winget and brew
# install under the user and must NOT be prefixed.
ELEVATED = ("apt", "dnf", "pacman")

TOOLS = ("node", "docx", "python", "pymupdf", "libreoffice")


def is_root():
    """True when the process already has root, if the platform has the notion."""
    geteuid = getattr(os, "geteuid", None)
    if geteuid is None:  # Windows
        return False
    return geteuid() == 0


def build(manager, tool):
    """Full argv to install `tool` with `manager` (a rutas.PackageManager).

    None when this manager has no plan for that tool, which the caller has to
    treat as "print manual instructions", not as an error.
    """
    plan = PLANS.get(tool, {}).get(manager.name)
    if plan is None:
        return None

    argv = []
    if manager.name in ELEVATED and not is_root():
        # -n is the whole point: without it sudo PROMPTS for a password and an
        # agent or a redirected stdin blocks forever. With it, the command fails
        # immediately and the manual instructions get printed.
        argv += ["sudo", "-n"]
    argv.append(manager.path)
    argv += [str(argument) for argument in plan]
    return argv


def npm_install_docx(npm, node_dir, pin):
    """argv for the docx install, run inside `node_dir`.

    --no-save and --no-package-lock because the node dir is generated state, not
    a project to keep a manifest of.
    """
    return [str(npm), "install", "docx@%s" % pin, "--no-save", "--no-package-lock",
            "--no-audit", "--no-fund"]


def npm_anchor_json():
    """package.json that anchors npm to the node dir.

    Without it, `npm install` walks UP the directory tree looking for a
    package.json and can drop node_modules in an unexpected parent folder, which
    is exactly the mismatch that broke the pipeline before.
    """
    return ('{"name": "apa7-workdir", "private": true, "version": "1.0.0", '
            '"description": "Anchors npm so that node_modules/docx installs here."}')


def manual_instructions(tool, is_windows=False, is_macos=False):
    """Copy-pasteable commands for a tool the installer cannot handle."""
    table = MANUAL.get(tool)
    if not table:
        return []
    if is_windows:
        key = "windows"
    elif is_macos:
        key = "macos"
    else:
        key = "linux"
    return [table[key]]


def python_version_note():
    """Shown next to the Python step.

    3.9 is the floor the skill's own code supports; install puts 3.12 in place.
    3.9 left upstream support in October 2025, so this is worth saying out loud
    instead of letting someone conclude the floor is a recommendation.
    """
    return ("Python 3.9 is the minimum the scripts support and it is out of "
            "upstream support (October 2025); install puts 3.12 in place.")