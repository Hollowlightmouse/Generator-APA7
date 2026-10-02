# System Requirements

Tools the pipeline needs to convert a `.md` into a `.docx` and then into
a `.pdf`. **Microsoft Word is not a requirement**: the conversion is done by
LibreOffice in headless mode.

Everything needed is resolved automatically, without hand-written paths. This
file explains how it is checked and how to adjust it if detection fails.

## Tools

| Tool | What it is used for | How it is checked | Automatic installation |
|---|---|---|---|
| PowerShell 5.1 or later | run the skill scripts | `$PSVersionTable.PSVersion` | ships with Windows |
| Node.js LTS | generate the `.docx` with the `docx` library | `node --version` | `winget install OpenJS.NodeJS.LTS` or `brew install node` |
| `docx` library (npm) | build the cover page, indexes, tables and figures | searched in the working directory | `npm install docx@9.7.1 --no-save` |
| Python 3.12 | create the virtual environment that runs the checks | searched on `PATH` and in the usual locations | `winget install Python.Python.3.12` or `brew install python@3.12` |
| `pymupdf` **1.28.2** | verify the resulting `.pdf` | run inside the virtual environment | `pip install pymupdf==1.28.2` (into `.venv`) |
| LibreOffice | **only engine** for `.docx` to `.pdf` conversion | `soffice --version` | `winget install TheDocumentFoundation.LibreOffice` or `brew install --cask libreoffice` |

The installer picks the available package manager (`winget` on Windows,
`brew` on macOS/Linux) and then runs the check again.

### Pinned versions

Both libraries are pinned to a single tested version, defined in one place,
`$APA7_DEPS` in `scripts/lib/rutas.ps1`:

```powershell
$script:APA7_DEPS = @{
    'docx'    = '9.7.1'
    'pymupdf' = '1.28.2'
}
```

Changing a version means changing it there, not in the individual scripts.

A **version mismatch is reported, not fatal**: the preflight keeps the line in
`OK` state and adds the warning inside the detail, so a different but working
library does not stop a document from being produced.

```
OK|pymupdf|1.28.2  [MISMATCH: pinned 1.26.0]  (…\.venv\Scripts\python.exe)
RESULT: OK
```

Only a genuinely missing tool produces `MISSING|…` and exit code 1.

## The virtual environment (`.venv`)

`pymupdf` is a compiled library, so it is **not** installed into the system
Python. The installer creates a virtual environment at `<skill>\.venv` and
installs the pinned `pymupdf` there:

```
<raiz de la skill>\.venv\Scripts\python.exe     Windows
<raiz de la skill>/.venv/bin/python             macOS / Linux
```

Everything resolves to that interpreter, in this order:

1. `APA7_PYTHON`, if the user has set it.
2. The skill's own `.venv`.
3. Any Python found on `PATH`, which is only used to *create* the `.venv`.

Consequences worth knowing:

- Each copy of the skill has its own `.venv`, so no two versions of `pymupdf`
  can be mixed.
- `.venv/` is in `.gitignore`; it is regenerated state, never source code.
- If `.venv` is deleted or its interpreter stops working, the installer
  recreates it from scratch. The preflight reports the consequence
  (`MISSING|pymupdf`) rather than crashing.
- Setting `APA7_PYTHON` overrides the `.venv`, for the rare case of needing an
  interpreter that is already provisioned elsewhere.

## Preflight

Checks everything and returns exit code 0 if the environment is complete, 1 if
something is missing. It always runs before transforming a document. It prints
one line per tool (`OK|<tool>|<detail>` or `MISSING|<tool>|<detail>`) and ends
with `RESULT: OK` or `RESULT: MISSING`.

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\comprobar-entorno.ps1"
```

- If it replies `RESULT: OK`, you can continue.
- If it replies `RESULT: MISSING` (the `MISSING|<tool>|...` lines), install what is missing:

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\instalar-entorno.ps1"
```

The installer applies the changes and repeats the check. If after finishing it
still does not return `RESULT: OK`, **stop and warn the user**: a document is
not transformed with an incomplete environment, because the failure appears
later as a malformed PDF and is much harder to diagnose.

## Paths: how they are resolved

No script has an absolute path written inside it. All of them go through
`scripts/lib/rutas.ps1`, which searches in this order:

1. The corresponding environment variable, if it is set.
2. The system's usual installation paths.
3. The `PATH`.

| Variable | What it overrides | Default value |
|---|---|---|
| `APA7_SKILL_ROOT` | skill root | the folder that contains `SKILL.md` |
| `APA7_WORKDIR` | generated working directory | `<raiz de la skill>\.work` |
| `APA7_NODEDIR` | where `node_modules` lives | the working directory |
| `APA7_SOFFICE` | LibreOffice executable | automatic detection |
| `APA7_PYTHON` | Python interpreter with `pymupdf`, bypassing `.venv` | the skill's `.venv` |

Example: if the `docx` library is somewhere else because the project already
had its own dependencies,

```powershell
$env:APA7_NODEDIR = "C:\ruta\al\proyecto\node_modules"
```

### Installer options

`instalar-entorno.ps1` is fully scriptable, so it can be run unattended:

| Option | Effect |
|---|---|
| `-WhatIf` | prints what would be installed and changes nothing (dry run) |
| `-Only <tool>` | touches only that tool: `node`, `docx`, `python`, `pymupdf` or `libreoffice` |

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\instalar-entorno.ps1" -WhatIf
powershell -ExecutionPolicy Bypass -File "scripts\instalar-entorno.ps1" -Only pymupdf
```

### Why the working directory exists

The `docx` library is installed in a directory of its own inside the skill and
not next to the user's document, so that `require('docx')` always resolves no
matter where the generator is run from. By default that directory is `.work/`
and it only contains generated state:

- `node_modules/` with the `docx` library.
- (it used to hold a LibreOffice profile as well; no longer, each run creates
  its own temporary one next to the output PDF).

It is regenerated with `instalar-entorno.ps1` and is in `.gitignore`. **It is
not source code and must not be edited or delivered.**

## Operational notes

- After installing with `winget` the `PATH` needs to be refreshed before
  invoking `node` or `soffice`; the installer already does it, but if you
  install by hand you have to reopen the terminal.
- The `PATH` of the session and the one of the process that launches the agent
  are different: that is why the scripts look for the paths on their own and do
  not assume that `node` is visible.
- If `node` is installed but `docx` is not, that is a different failure from
  the previous one and the preflight reports it separately.
- Python detection tries several interpreters, because there may be more than
  one installed and not all of them have `pymupdf`. The one inside `.venv`
  always wins; the others are only used to create it.
- LibreOffice runs with a hard time limit (`-Timeout`, 300 s by default). When
  it expires the process tree is killed and the run is reported as a **timeout**
  (exit code 124), not as a silent conversion failure: those are different
  problems with different remedies, and confusing them wastes hours.

### Why LibreOffice runs redirected to files

LibreOffice is launched with `Start-Process` and its output sent to temporary
files instead of being piped. On Windows a child that inherits the caller's
redirected handle can keep `soffice` alive after the script has finished, which
is what used to leave the pipeline hanging. The files are read after the process
ends and then deleted.

There is one subtlety that is easy to get wrong: reading `.ExitCode` on the
`Process` object returned by `Start-Process -PassThru` only works if the handle
was touched first (`$proc.Handle`). Without that line the exit code comes back
empty — not zero, empty — and a successful run looks like a failure.