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
| Python with `pymupdf` | verify the resulting `.pdf` | `python -c "import fitz"` | `pip install pymupdf` |
| LibreOffice | **only engine** for `.docx` to `.pdf` conversion | `soffice --version` | `winget install TheDocumentFoundation.LibreOffice` or `brew install --cask libreoffice` |

The installer picks the available package manager (`winget` on Windows,
`brew` on macOS/Linux) and then runs the check again.

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
| `APA7_PYTHON` | Python interpreter with `pymupdf` | automatic detection |

Example: if the `docx` library is somewhere else because the project already
had its own dependencies,

```powershell
$env:APA7_NODEDIR = "C:\ruta\al\proyecto\node_modules"
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
  one installed and not all of them have `pymupdf`. It can be pinned with
  `APA7_PYTHON`.