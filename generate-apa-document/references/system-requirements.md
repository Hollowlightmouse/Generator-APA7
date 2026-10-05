# System Requirements

Tools the pipeline needs to convert a `.md` into a `.docx` and then into
a `.pdf`. **Microsoft Word is not a requirement**: the conversion is done by
LibreOffice in headless mode.

Everything needed is resolved automatically, without hand-written paths and
without a particular shell. This file explains how it is checked and how to
adjust it if detection fails.

## Tools

| Tool | What it is used for | How it is checked | Automatic installation |
|---|---|---|---|
| Python 3.9 or later | run the skill scripts | run a probe, not just `--version` | ships with the OS, or the installer provisions 3.12 |
| Node.js LTS | generate the `.docx` with the `docx` library | `node --version` | per package manager, see below |
| `docx` library (npm) | build the cover page, indexes, tables and figures | searched in `node_modules` | `npm install docx@9.7.1 --no-save` |
| Python 3.12 | create the virtual environment that runs the checks | searched on `PATH` and in the usual locations | per package manager, see below |
| `pymupdf` **1.28.2** | verify the resulting `.pdf` | imported inside the virtual environment | `pip install pymupdf==1.28.2` (into `.venv`) |
| LibreOffice | **only engine** for `.docx` to `.pdf` conversion | `soffice --version` | per package manager, see below |

The installer picks the available package manager and then runs the check
again:

| Platform | Manager | Used for |
|---|---|---|
| Windows | `winget` | Node, Python, LibreOffice |
| macOS | `brew` | Node, Python, and `brew install --cask libreoffice` |
| Linux | `apt`, `dnf`, `pacman` | Node, Python, LibreOffice (package name per distribution) |

If none is available, or if the package is not in its repository, the installer
prints the exact command to run by hand instead of failing silently.

### Pinned versions

Both libraries are pinned to a single tested version, defined in one place,
`DEPENDENCIAS` in `scripts/lib/rutas.py`:

```python
DEPENDENCIAS = {
    "docx": "9.7.1",
    "pymupdf": "1.28.2",
}
```

Changing a version means changing it there, not in the individual scripts.

A **version mismatch is reported, not fatal**: the preflight adds the warning
inside the detail and keeps the environment usable, so a different but working
library does not stop a document from being produced.

```
INFO|pymupdf|1.28.2  [MISMATCH: pinned 1.26.0]  (…/.venv/bin/python)
RESULT: OK
```

Only a genuinely missing tool produces `MISSING|…` and exit code 1.

### Why the Python floor is 3.9

`apa7.py` and the scripts it calls must import on Python **3.9 or later**, so
they use the standard library only and avoid syntax newer than that. This is
checked in the tests by parsing every source file with `feature_version=(3, 9)`.

3.9 reached end of life in October 2025, so it is accepted for compatibility
but not recommended. `check` reports an interpreter on `PATH` older than the
floor as `INFO`, not as a failure, because the skill's own `.venv` is what
actually runs the verification. `install` provisions **3.12**.

## The virtual environment (`.venv`)

`pymupdf` is a compiled library, so it is **not** installed into the system
Python. The installer creates a virtual environment at `<skill>/.venv` and
installs the pinned `pymupdf` there:

```
<raiz de la skill>/.venv/Scripts/python.exe     Windows
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
one line per tool (`OK|<tool>|<detail>`, `INFO|<tool>|<detail>` or
`MISSING|<tool>|<detail>`) and ends with `RESULT: OK` or `RESULT: MISSING`.

The examples use `python`; use whichever interpreter exists (`python3` on macOS
and most Linux distributions, `py` on Windows). `apa7.py`'s own messages always
print the exact command with the interpreter it is running.

```bash
python scripts/apa7.py check
```

- If it replies `RESULT: OK`, you can continue.
- If it replies `RESULT: MISSING` (the `MISSING|<tool>|...` lines), install what is missing:

```bash
python scripts/apa7.py install
```

The installer applies the changes and repeats the check. If after finishing it
still does not return `RESULT: OK`, **stop and warn the user**: a document is
not transformed with an incomplete environment, because the failure appears
later as a malformed PDF and is much harder to diagnose.

An `INFO` line is a usable environment with something worth knowing. It never
changes the result, so there is nothing to install because of it.

## Paths: how they are resolved

No script has an absolute path written inside it. All of them go through
`scripts/lib/rutas.py`, which searches in this order:

1. The corresponding environment variable, if it is set.
2. The system's usual installation paths.
3. The `PATH`.

| Variable | What it overrides | Default value |
|---|---|---|
| `APA7_SKILL_ROOT` | skill root | the folder that contains `SKILL.md` |
| `APA7_WORKDIR` | generated working directory | `<raiz de la skill>/.work` |
| `APA7_NODEDIR` | where `node_modules` lives | the working directory |
| `APA7_SOFFICE` | LibreOffice executable | automatic detection |
| `APA7_PYTHON` | Python interpreter with `pymupdf`, bypassing `.venv` | the skill's `.venv` |

An `APA7_SKILL_ROOT` that does not contain `SKILL.md` is **warned about and
ignored**: pointing the scripts at an unrelated directory produces a confusing
cascade of missing files, and the warning points at the cause.

### Where a document's own files go

Tools are resolved relative to the skill; a **document's** files are resolved
relative to that document. `scripts/lib/rutas.py` owns this and is the only
place that decides it, so no caller has to invent a path:

| What | Where it lands by default |
|---|---|
| `parse` manifest (`--out`) | `<md>_apa/datos/MANIFEST.json` |
| `parse` log (`--log`) | `<md>_apa/logs/01-analisis.log` |
| `build` output (`--out`) | next to the `.md`, with the same name |
| `build` log (`--log`) | `<md>_apa/logs/02-build.log` |
| double-pass page map | `<md>_apa/logs/paginas.json` |
| double-pass throwaway PDF | `<md>_apa/logs/`, deleted once pass 2 succeeds |
| `export` PDF (`--outdir`) | next to the `.docx`, with the same name |
| `export` log (`--log`) | `<md>_apa/logs/03-export.log` |
| `verify` report (`--json`) | `<md>_apa/datos/verificacion.json` |

`paginas-de-pdf.py` is called **without** `--out`, so it writes `paginas.json`
into its working folder rather than into the folder the first `.docx` passed to
it; the first pass uses a temporary copy precisely so it has no folder of its own
to write into.

The `<stem>_apa` name is derived from the `.md`: NFKD without diacritics,
invalid/control characters and spaces become `_`, Windows reserved names get a
`_` suffix, capped at 60 characters. `Informe técnico.md` becomes
`Informe_tecnico_apa/`. The `.docx` and the `.pdf` keep their original name.

A phase that is handed no document at all (`export` without `--docx`) looks for
a single `.docx` in the **current folder** — the one that holds the `.md` and the
`.docx`, not the working folder — and refuses to choose between zero or several.
`datos/fuente.json` records which `.md` a folder belongs to, so a later phase can
resolve the document without being told again; pointing a folder that already has
an anchor at a different `.md` warns and updates it.

`--carpeta-trabajo <dir>` overrides the working folder for `parse`, `build`,
`export` and `verify`, and is the only way to put the logs somewhere other than
next to the document. **Explicit flags win**: an `--outdir` on `export` keeps its
historical behaviour (PDF *and* `<outdir>/_logs/03-export.log`), which is what
existing callers depend on.

Example: if the `docx` library is somewhere else because the project already
had its own dependencies,

```bash
export APA7_NODEDIR=/path/to/project/node_modules
```

### Installer options

`apa7.py install` is fully scriptable, so it can be run unattended:

| Option | Effect |
|---|---|
| `--dry-run` | prints what would be installed and changes nothing |
| `--only <tool>` | touches only that tool: `node`, `docx`, `python`, `pymupdf` or `libreoffice` (repeatable) |
| `--yes` | do not ask for confirmation |

```bash
python scripts/apa7.py install --dry-run
python scripts/apa7.py install --only pymupdf
python scripts/apa7.py install --yes
```

`--yes` is **required** when stdin is not a terminal, because the confirmation
would otherwise read from nothing.

`--only` installs just what was asked for and does not fail because something
else is absent; the final `check` is what decides whether the environment is
complete.

Installing a system package needs administrator rights. The installer tries
`sudo -n` first so it never blocks waiting for a password, and when that fails
it prints the command to run by hand.

### Why the working directory exists

The `docx` library is installed in a directory of its own inside the skill and
not next to the user's document, so that `require('docx')` always resolves no
matter where the generator is run from. By default that directory is `.work/`
and it only contains generated state:

- `node_modules/` with the `docx` library.
- `package.json`, the anchor that makes `npm install` write there.
- (it used to hold a LibreOffice profile as well; no longer, each run creates
  its own temporary one next to the output PDF).

It is regenerated with `apa7.py install` and is in `.gitignore`. **It is
not source code and must not be edited or delivered.**

## Operational notes

- After installing with `winget` or `brew` the `PATH` needs to be refreshed
  before invoking `node` or `soffice`; the installer already does it, but if you
  install by hand you have to reopen the terminal.
- The `PATH` of the session and the one of the process that launches the agent
  are different: that is why the scripts look for the paths on their own and do
  not assume that `node` is visible.
- If `node` is installed but `docx` is not, that is a different failure from
  the previous one and the preflight reports it separately.
- Python detection tries several interpreters, because there may be more than
  one installed and not all of them have `pymupdf`. The one inside `.venv`
  always wins; the others are only used to create it.
- LibreOffice runs with a hard time limit (`--timeout`, 300 s by default). When
  it expires the process tree is killed and the run is reported as a **timeout**
  (exit code 124), not as a silent conversion failure: those are different
  problems with different remedies, and confusing them wastes hours.
- `export` closes any LibreOffice left running by a previous attempt before
  starting, because a stale instance holds a lock and the new run then exits
  without converting anything.

### Why LibreOffice runs with its output redirected to files

LibreOffice is launched with its stdout and stderr sent to temporary files
instead of being piped, and those files are read after the process ends and
then deleted.

The reason is not cosmetic. A child that inherits the caller's redirected pipe
handle can keep `soffice` alive after the parent has finished, which is what
used to leave the pipeline hanging forever with no output. Redirecting to files
means the child never holds the caller's pipe.

The same care applies to how the process is started: launch the resolved binary
directly, and do **not** use `Start-Process` or a shell background operator
(`&`, `Start-Job`). See `word-toc-fields.md`.