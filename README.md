# Skill `generar-pdf-apa`

Builds a **Word (.docx) and PDF** academic document following the **APA 7th edition** rules, ready to submit: institutional cover page, table of contents and lists of tables/figures with real page numbers, all from a source Markdown `.md` (extracted with MinerU/Docling, for example).

> **Export engine: LibreOffice (headless). Microsoft Word is not required.**

---

## What it is and what it is for

It is an **agent skill** (a `SKILL.md` plus helper scripts) that turns the output — usually imperfect — of a PDF-to-Markdown conversion into a **final, verified and presentable** academic paper.

It is not a PDF-to-Word converter: it starts from a `.md` that already exists. What it does is **structure and formalize** it:

- **Institutional cover page** with a 3-zone layout (title at the top, authors vertically centered, institutional block anchored at the bottom) and an optional logo only if you provide it.
- **Table of contents always**, and **lists of tables and figures only when they exist**, with real pages via `PAGEREF` fields that update when the document is opened.
- **Automatic correction** of conversion artifacts: glued words (`2:Diferenciaciónentre Bugs` → `2: Diferenciación entre Bugs`), duplicated list markers (`• •`, `1. 1.`), glued numbering (`ACTIVIDAD1:`), always respecting proper nouns and technical terms.
- **References normalized to APA 7**: alphabetical order, hanging indent and italics where appropriate, even if the `.md` already brings them as a list.
- **Wide tables on an occasional landscape page** (allowed by APA 7) for comparison matrices that do not fit in portrait.
- **Final verification with `pymupdf`** before delivering: cover, indices, captions, landscape pages, notes and references.
- Delivers **both files** (`.docx` and `.pdf`).

## Requirements (tools it needs)

| Tool | What it is used for | How to get it |
|---|---|---|
| **PowerShell 5.1+** | running the skill scripts | included in Windows |
| **Node.js LTS + npm** | generating the `.docx` with the `docx` library | `winget install OpenJS.NodeJS.LTS` or `brew install node` |
| **`docx` library (npm)** | building the cover, TOC, indices and tables | `npm install docx@9.7.1 --no-save` |
| **Python 3 + `pymupdf`** | verifying the final PDF | `pip install pymupdf` |
| **LibreOffice** | **the only engine** for `.docx → .pdf` conversion | `winget install TheDocumentFoundation.LibreOffice` or `brew install --cask libreoffice` |
| **Internet connection** | winget/brew, npm and pip downloads | — |

The skill validates these tools automatically before touching the document (**STEP 0 / preflight**). If any is missing, it installs it; if something cannot be installed, it **stops and warns** without transforming the document.

Tool paths overridable by environment variable: `APA7_SOFFICE`, `APA7_PYTHON`, `APA7_NODEDIR`, `APA7_WORKDIR`, `APA7_SKILL_ROOT`. Details in [`generar-pdf-apa/references/system-requirements.md`](generar-pdf-apa/references/system-requirements.md).

## Installation

1. **Copy the** `generar-pdf-apa/` **folder** into your AI agent's skills directory, so that `SKILL.md` ends up at the root of the skills directory. For example: `<your-user>\.agents\skills\generar-pdf-apa`.
2. **The environment sets itself up.** The first time, run the preflight:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "...\generar-pdf-apa\scripts\comprobar-entorno.ps1"
   ```

   If it returns `RESULT: MISSING`, the automatic installer resolves it and re-checks:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "...\generar-pdf-apa\scripts\instalar-entorno.ps1"
   ```

   > Do not install anything by hand and do not transform documents with missing tools.

The `docx` dependency is installed in `generar-pdf-apa/.work/node_modules`, inside the skill itself: that way `require('docx')` always resolves, no matter where it is run from. That folder is generated state and is in `.gitignore`.

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

`.docx` and `.pdf` verified page by page, ready to submit.

## Workflow

1. **Activation** — the user asks for an academic document in APA (by intent).
2. **STEP 0 · Preflight** — `comprobar-entorno.ps1` checks the tools. If something is missing, `instalar-entorno.ps1` installs it and re-checks. If it does not return `RESULT: OK`, it stops and warns.
3. **Inputs** — the `.md` (mandatory), the images and the layout JSON (optional) are received.
4. **Parser** — `md-a-manifiesto.py` turns `.md` + layout JSON into `MANIFEST.json`: it enriches, deduplicates, splits glued words, computes real image sizes from the JSON `bbox` and applies the default note. It leaves the pending questions in `diagnostico`.
5. **Mandatory questions (STEP 0.5)** — if `pendientes_bloqueantes` is not empty, the user is asked (cover data, missing titles/captions, "no tables/figures" confirmation) and nothing is built until they are resolved. The cover answers are saved in `portada.json` and passed to the parser with `--portada`.
6. **Build the `.docx`** — `build-docx.js` reads `MANIFEST.json` and builds the cover, TOC, indices, tables, figures and references. Aborts with code 4 if a title or caption is missing.
7. **Export to PDF** — `export-pdf.ps1` converts `.docx → .pdf` with headless LibreOffice (temporary copy + isolated profile per run).
8. **Verify** — `verificar-pdf.py` (with `pymupdf`) checks the cover, indices with the correct page, captions, table notes below, figure notes above their image and hanging indent in references. If a critical check fails, it is fixed and exported again.
9. **Deliver** — both files are delivered (`.docx` and `.pdf`).

```
Activation
   └─ STEP 0  Environment preflight ──► (something missing) instalar-entorno.ps1 ──► re-check
        └─ Inputs: .md (+ images + layout JSON)
             └─ md-a-manifiesto.py ──► MANIFEST.json (+ diagnostico)
                  └─ STEP 0.5  Blocking pending questions? ──► (yes) ask the user
                       └─ build-docx.js ──► document.docx
                            └─ export-pdf.ps1 (headless LibreOffice) ──► document.pdf
                                 └─ verificar-pdf.py ──► (failures) fix ──► re-export
                                      └─ Delivery: .docx + .pdf
```

## How it works inside

- **Image size without guessing by eye:** the parser reads each figure's `bbox` in the layout JSON (normalized to 0..1000), keeps its real proportion and clamps the width to the usable content width.
- **Prose and order come from the `.md`**, not the JSON: the JSON files enrich the table width and the figure size, but do not rewrite the text.
- **Genuinely functional TOC:** the entries use `PAGEREF` fields over *bookmarks*, so the page number is recalculated when the document is opened. LibreOffice does not resolve `SEQ` fields, which is why table and figure numbering is literal.
- **Tables renderable in LibreOffice:** explicit width (`width` + `columnWidths` + `layout: FIXED`) and horizontal-only borders; without this, LibreOffice does not show them (known bug).
- **Automatic per-table orientation:** a table that does not fit in portrait (6 or more columns, or very long cells) moves on its own to a **landscape page** with its title and note; afterwards the text returns to portrait. The criterion is legibility, not the number of rows.
- **Figure note above the image; table note below the table** (APA 7 distinguishes them by position, not only by text).
- **Clean export:** the resolved `soffice` binary is invoked directly (not via `Start-Process`), with an isolated LibreOffice profile per run and a temporary copy; the `Could not find platform independent libraries <prefix>` message on stderr is benign and ignored.
- **No embedded whitelists:** the terms the deglue must not split live only in `references/terms-whitelist.txt`.

## Repository structure

```
Generator-APA7/
├── README.md
└── generar-pdf-apa/
    ├── SKILL.md                          # skill definition (activation, workflow, rules)
    ├── .gitignore                        # excludes generated state (.work, __pycache__)
    ├── references/
    │   ├── apa7-format.md                # APA 7 formatting rules
    │   ├── institutional-cover.md        # 3-zone cover layout
    │   ├── word-toc-fields.md            # functional TOC/indices and docx/LibreOffice pitfalls
    │   ├── system-requirements.md        # tools, preflight and path resolution
    │   └── terms-whitelist.txt           # untouchable terms for the deglue
    ├── scripts/
    │   ├── comprobar-entorno.ps1         # environment preflight (STEP 0)
    │   ├── instalar-entorno.ps1          # auto-install of what is missing
    │   ├── md-a-manifiesto.py            # .md + JSON → MANIFEST.json
    │   ├── build-docx.js                 # MANIFEST.json → .docx
    │   ├── export-pdf.ps1                # .docx → .pdf with LibreOffice
    │   ├── verificar-pdf.py              # .pdf verification (pymupdf)
    │   └── lib/
    │       └── rutas.ps1                 # portable path and tool resolution
    └── .work/                            # generated state (docx's node_modules)
```

## Limitations

- **Windows / PowerShell 5.1+** (the scripts are designed for this environment).
- It does not convert directly from PDF to Word: it starts from a `.md` already extracted by MinerU or Docling. The layout JSON is optional, but without it the figures do not keep their real size.
- **APA 7 letter format with 1 in margins**, with no extra institutional rules.
- The cover page carries no visible page number; numbering starts on page 2.

---

*Skill `generar-pdf-apa` — APA 7 academic document (.docx + PDF) without depending on Microsoft Word.*
