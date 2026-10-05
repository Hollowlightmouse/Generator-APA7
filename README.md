# Generate an APA 7 academic document (.docx + PDF)

Skill `generate-apa-document`. Builds a **Word (.docx) and PDF** academic document following the **APA 7th edition** rules, ready to submit: institutional cover page, table of contents and lists of tables/figures with real page numbers, all from a source Markdown `.md` (extracted with MinerU/Docling, for example).

> **Export engine: LibreOffice (headless). Microsoft Word is not required.**

---

## What it is and what it is for

It is an **agent skill** (a `SKILL.md` plus helper scripts) that turns the output — usually imperfect — of a PDF-to-Markdown conversion into a **final, verified and presentable** academic paper.

It is not a PDF-to-Word converter: it starts from a `.md` that already exists. What it does is **structure and formalize** it:

- **Institutional cover page** with a 3-zone layout (title at the top, authors vertically centered, institutional block anchored at the bottom) and an optional logo only if you provide it.
- **Table of contents always**, and **lists of tables and figures only when they exist**, as real Word `TOC` fields. The page numbers are measured by a **double pass** (build → throwaway PDF → `paginas-de-pdf.py` → rebuild) and shipped as the cached field result, so the PDF has them right even without Word; Word recalculates them on open.
- **Automatic correction** of conversion artifacts: glued words (`2:Diferenciaciónentre Bugs` → `2: Diferenciación entre Bugs`), duplicated list markers (`• •`, `1. 1.`), glued numbering (`ACTIVIDAD1:`). Pure UPPERCASE acronyms (`OWASP`, `CVSS`), URLs, DOIs, e-mails and file names are never touched, and **every** change is reported, split into clear corrections and items to review, so a term split by mistake can be protected or fixed in the source `.md`.
- **References normalized to APA 7**: alphabetical order, hanging indent and italics where appropriate, even if the `.md` already brings them as a list.
- **Wide tables on an occasional landscape page** (allowed by APA 7) for comparison matrices that do not fit in portrait.
- **Final verification with `pymupdf`** before delivering: cover, indices, captions, landscape pages, notes and references.
- Delivers **both files** (`.docx` and `.pdf`).

## Requirements (tools it needs)

| Tool | What it is used for | How to get it |
|---|---|---|
| **Python 3.9+** | running the skill scripts | ships with macOS/Linux; `winget install Python.Python.3.12` or `brew install python@3.12` |
| **Node.js LTS + npm** | generating the `.docx` with the `docx` library | `winget install OpenJS.NodeJS.LTS`, `brew install node`, or the distro package |
| **`docx` library (npm)** | building the cover, TOC, indices and tables | `npm install docx@9.7.1 --no-save` |
| **`pymupdf` 1.28.2** | verifying the final PDF | installed into `.venv`, not into the system Python |
| **LibreOffice** | **the only engine** for `.docx → .pdf` conversion | `winget install TheDocumentFoundation.LibreOffice`, `brew install --cask libreoffice`, or the distro package |
| **Internet connection** | package manager, npm and pip downloads | — |

Everything above except the internet is installed by `apa7.py install`, which picks the available package manager (`winget`, `brew`, `apt`, `dnf` or `pacman`) or prints the command to run by hand.

The skill validates these tools automatically before touching the document (**STEP 0 / preflight**). If any is missing, it installs it; if something cannot be installed, it **stops and warns** without transforming the document.

**Which Python command?** Use whichever your machine actually has: `python3` on macOS and most Linux distributions, `python` where it exists, and `py` on Windows. This README writes `python` for brevity, but every message printed by `apa7.py` shows the exact, copyable command using the interpreter that is running. Check the version first (`python3 --version`, `python --version` or `py --version`); it must be **3.9 or later**.

Tool paths overridable by environment variable: `APA7_SOFFICE`, `APA7_PYTHON`, `APA7_NODEDIR`, `APA7_WORKDIR`, `APA7_SKILL_ROOT`. Details in [`generate-apa-document/references/system-requirements.md`](generate-apa-document/references/system-requirements.md).

## Installation

1. **Copy the** `generate-apa-document/` **folder** into your AI agent's skills directory, so that `SKILL.md` ends up at the root of the skills directory. For example: `<your-user>\.agents\skills\generate-apa-document`.
2. **The environment sets itself up.** The first time, run the preflight:

   ```bash
   python scripts/apa7.py check
   ```

   If it returns `RESULT: MISSING`, the automatic installer resolves it and re-checks:

   ```bash
   python scripts/apa7.py install
   ```

   Add `--dry-run` to see the plan without changing anything, and `--yes` when running unattended.

   > Do not install anything by hand and do not transform documents with missing tools.

The `docx` dependency is installed in `generate-apa-document/.work/node_modules`, inside the skill itself: that way `require('docx')` always resolves, no matter where it is run from. That folder is generated state and is in `.gitignore`.

## How to use it

The skill **activates by intent**: it is enough that you ask for an academic document assembled in APA (*"generate the APA PDF"*, *"armá mi trabajo en formato APA"*, *"estructura este informe con norma APA"*, *"pasame esto a APA con portada y referencias"*). It does not activate merely because a `.md` exists, nor when you only ask to read, summarize, translate or spell-check.

### 1. Provide the inputs

| # | Input | What it is for | Mandatory |
|---|--------|----------------|-------------|
| 1 | The source **`.md`** | text, headings, markdown tables, references and image markers `![]()` | yes |
| 2 | The referenced **images** | inserted into the document (`images/` folder or loose files) | no |
| 3 | The **layout JSON** (`*_content_list.json`, `*_content_list_v2.json`, `*_middle.json`, `*_model.json`) | real size and position of each figure and table width | no |

The pipeline accepts three levels of input:

- `.md` only → text document, no figures and no lists of tables/figures.
- `.md` + `images/` → figures at a reasonable fixed size.
- `.md` + images + JSON → figures at their **real size** (from the `bbox`) and tables at their real page width.

> Golden rule: if something essential is missing, the skill **asks you** before continuing. It never assumes or invents content.

### 2. Answer only what is missing

Almost everything comes from your files. The skill will only ask you for the **data that does not appear** in your documents: cover data (instructor and their title/profession, course and its code, date), logo (optional, only if you provide it), table or figure titles/captions that do not come with one and, when in doubt, whether the document really has no tables or figures.

**Hard rule:** while there are questions marked as blocking, the document **is not built**. Nothing is filled with invented text and the `.docx` is not patched by hand.

### 3. Receive the two files

`.docx` and `.pdf` verified page by page, ready to submit, **next to the source `.md`** and with the same name.

### Where the files land

You do not pass `--out`, `--log`, `--outdir` or `--json`: each phase derives its own paths from the document it is handed.

```
<carpeta del .md>/
  informe.md                 ← your source, untouched
  informe.docx               ← deliverable, same name as the .md
  informe.pdf                ← deliverable, same name as the .md
  informe_apa/               ← working folder, name derived and sanitized
    datos/
      fuente.json            ← which .md this folder belongs to
      MANIFEST.json          ← the manifest
      verificacion.json      ← verification report
    logs/
      01-analisis.log        ← parse
      02-build.log           ← build
      paginas.json           ← measured page numbers (double pass)
      03-export.log          ← export
```

- The deliverables keep the **original** name of the `.md`; only the working folder is renamed: no diacritics, invalid characters and spaces become `_`, Windows reserved names get a `_` suffix, capped at 60 characters (`Informe técnico.md` → `Informe_tecnico_apa/`).
- **The working folder is reused, never versioned.** Rerunning the pipeline replaces the files in place instead of producing `informe (2).docx`.
- **The throwaway PDF of the double pass is deleted** when the second pass succeeds; if it fails, that PDF stays, because it is the only record of what was measured. A failed verification likewise keeps its PDF and its whole working folder.
- `--carpeta-trabajo <dir>` moves the working folder when the document lives elsewhere, and works on `parse`, `build`, `export` and `verify`.
- **Explicit flags always win.** Passing `--outdir` to `export` keeps its old behaviour (PDF *and* `_logs/03-export.log` there), so existing scripts and callers are unaffected. That is the only case that still writes `_logs/`.

## Workflow

1. **Activation** — the user asks for an academic document in APA (by intent).
2. **STEP 0 · Preflight** — `apa7.py check` checks the tools. If something is missing, `apa7.py install` installs it and re-checks. If it does not return `RESULT: OK`, it stops and warns.
3. **Inputs** — the `.md` (mandatory), the images and the layout JSON (optional) are received.
4. **Parser** — `apa7.py parse` turns `.md` + layout JSON into `MANIFEST.json` (in `<md>_apa/datos/`): it enriches, deduplicates, splits glued words (protecting URLs, DOIs, e-mails, file names and any term passed with `--terminos-protegidos`), computes real image sizes from the JSON `bbox` and applies the default note. It leaves the pending questions in `diagnostico`.
5. **Mandatory questions (STEP 0.5)** — if `pendientes_bloqueantes` is not empty, the user is asked (cover data, missing titles/captions, "no tables/figures" confirmation) and nothing is built until they are resolved. The cover answers are saved in `portada.json` and passed to the parser with `--portada`.
6. **Build the `.docx`** — `apa7.py build` reads `MANIFEST.json` and builds the cover, TOC, indices, tables, figures and references, writing the `.docx` next to the source `.md`. It runs in **two passes**: it exports the first `.docx` to a throwaway PDF in `<md>_apa/logs/`, measures the real page of every entry with `paginas-de-pdf.py` and rebuilds with `--paginas-json` to cache those numbers in the `TOC` fields. The throwaway PDF is deleted once the second pass succeeds. Aborts with code 4 if a title or caption is missing.
7. **Export to PDF** — `apa7.py export` converts `.docx → .pdf` with headless LibreOffice (temporary copy + isolated profile per run), writing the PDF next to the `.docx`.
8. **Verify** — `apa7.py verify` (with `pymupdf`) checks the cover, indices with the correct page, captions, table notes below, figure notes above their image and hanging indent in references. If a critical check fails, it is fixed and exported again; nothing is deleted, so the failing PDF stays for inspection.
9. **Deliver** — both files are delivered (`.docx` and `.pdf`).

```
Activation
   └─ STEP 0  Environment preflight ──► (something missing) apa7.py install ──► re-check
        └─ Inputs: .md (+ images + layout JSON)
             └─ apa7.py parse ──► MANIFEST.json (+ diagnostico)
                  └─ STEP 0.5  Blocking pending questions? ──► (yes) ask the user
                       └─ apa7.py build ──► document.docx
                            └─ apa7.py export (headless LibreOffice) ──► document.pdf
                                 └─ apa7.py verify ──► (failures) fix ──► re-export
                                      │       └── Delivery: .docx + .pdf next to the .md, plus informe_apa/ to inspect or delete
```

## How it works inside

- **Image size without guessing by eye:** the parser reads each figure's `bbox` in the layout JSON (normalized to 0..1000), keeps its real proportion and clamps the width to the usable content width.
- **Prose and order come from the `.md`**, not the JSON: the JSON files enrich the table width and the figure size, but do not rewrite the text.
- **Genuinely functional TOC:** the indices are real `TOC` fields (`TOC \h \u` for the contents, `TOC \c "Tabla" \h` / `TOC \c "Figura" \h` for the lists), and their cached result carries the real pages measured in the double pass; Word recalculates them when the document is opened. Table and figure captions use real `SEQ` fields, which is what `TOC \c` collects.
- **Tables renderable in LibreOffice:** explicit width (`width` + `columnWidths` + `layout: FIXED`) and horizontal-only borders; without this, LibreOffice does not show them (known bug).
- **Automatic per-table orientation:** a table that does not fit in portrait (6 or more columns, or very long cells) moves on its own to a **landscape page** with its title and note; afterwards the text returns to portrait. The criterion is legibility, not the number of rows.
- **Figure note above the image; table note below the table** (APA 7 distinguishes them by position, not only by text).
- **Clean export:** the resolved `soffice` binary is invoked directly, with its output redirected to files (never to an inherited pipe), an isolated LibreOffice profile per run and a temporary copy. Because the profile is isolated, a LibreOffice the user has open is left alone; only the skill's own leftover runs are closed (opt in to closing everything with `export --cerrar-libreoffice`, which warns that unsaved documents are lost). The run has a hard timeout reported as exit code 124 rather than as a silent failure. The `Could not find platform independent libraries <prefix>` message on stderr is benign and ignored.
- **Terms kept intact only on request:** the pipeline ships no whitelist of untouchable terms, because that data belongs to one concrete document and not to the tool. URLs, DOIs, e-mails and file names are protected automatically; any other term (`NodeJS`, `SHA256`) can be kept intact by listing it in a protected-terms file and passing `--terminos-protegidos <ARCHIVO>`. Otherwise the deglue splits it and reports it under **Review** (ambiguous) instead of among the clear corrections, and the residue report lists every stretch it could not separate.

## Repository structure

```
Generator-APA7/
├── .gitignore                           # single ignore file: generated state, env data
├── README.md
└── generate-apa-document/
    ├── SKILL.md                          # skill definition (activation, workflow, rules)
    ├── references/
    │   ├── apa7-format.md                # APA 7 formatting rules
    │   ├── institutional-cover.md        # 3-zone cover layout
    │   ├── word-toc-fields.md            # functional TOC/indices and docx/LibreOffice pitfalls
    │   └── system-requirements.md        # tools, preflight and path resolution
    ├── scripts/
    │   ├── apa7.py                      # entry point: check/install/export/parse/build/verify
    │   ├── md-a-manifiesto.py            # .md + JSON → MANIFEST.json  (apa7.py parse)
    │   ├── build-docx.js                 # MANIFEST.json → .docx        (apa7.py build)
    │   ├── paginas-de-pdf.py             # real index pages, 2nd pass  (apa7.py build)
    │   ├── verificar-pdf.py              # .pdf verification (pymupdf) (apa7.py verify)
    │   ├── tests/                       # unit tests
    │   └── lib/
    │       ├── rutas.py                  # portable path and tool resolution, and where a document's files go
    │       ├── instalador.py             # installation plans per package manager
    │       └── fuentes.py                # font names accepted by the verifier
    ├── .work/                            # generated state (docx's node_modules)
    └── .venv/                            # generated state (pinned pymupdf)
```

## Limitations

- **Windows, macOS and Linux.** The CLI is cross-platform: Python 3.9+ and no shell-specific syntax. The interpreter name is not hardcoded: use `python`, `python3` or `py`, whichever exists, and `apa7.py` prints the exact command it was run with. LibreOffice is found through `PATH` and the documented install locations; the Snap (`/snap/bin/libreoffice`) and Flatpak (`org.libreoffice.LibreOffice`) launchers and the `apt`/`dnf`/`pacman` plans are **best effort and have not been verified on this machine** — only the Windows locations have been exercised.
- It does not convert directly from PDF to Word: it starts from a `.md` already extracted by MinerU or Docling. The layout JSON is optional, but without it the figures do not keep their real size.
- **APA 7 letter format with 1 in margins**, with no extra institutional rules.
- The cover page carries no visible page number; numbering starts on page 2.
- The `TOC` fields and the double pass are verified against **LibreOffice** (the shipped PDF). Word is not required and has not been tested locally; if a user updates the index fields *in Word*, the lists of tables/figures may keep only the number (Word's `\c` collects the caption paragraph, and the APA caption number and title are separate paragraphs). The delivered PDF already carries the full entries.
- **Automated checks (CI).** `.github/workflows/ci.yml` runs the test suite (`unittest`) plus `apa7.py check` and `install --dry-run` on Windows, macOS and Linux with Python 3.9 and 3.12. It does **not** install LibreOffice, so the end-to-end `.docx → .pdf` export is not exercised there. The workflow has not run yet (there is no CI history in the repository), so the cross-OS claims above are pending that first run.

## License

No license has been chosen yet. Until one is added, the code is **all rights reserved**: no one may reuse, modify or redistribute it. To allow others to use it, add a `LICENSE` file (for example MIT or Apache-2.0) and record the choice here.

---

*Skill `generate-apa-document` — APA 7 academic document (.docx + PDF) without depending on Microsoft Word.*
