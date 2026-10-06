---
name: generate-apa-document
description: Builds or structures an academic Word (.docx) and PDF document following APA 7th edition, from a source Markdown (.md) file, its images and the layout JSON files (content_list, content_list_v2, middle, model) that carry figure size and position. ALWAYS USE when the user asks to assemble or produce an academic APA document or PDF, formalize a paper, set it up with an institutional cover page, table of contents, list of tables, list of figures and APA references, or apply APA formatting to a text. Activation is by INTENT, not keywords: "generate the APA PDF", "armá mi trabajo en formato APA", "estructura este informe con norma APA", "pasame esto a APA con portada y referencias" all apply. Also use when required data is missing and the user must be asked before continuing. Exports the PDF with headless LibreOffice. Microsoft Word is optional and used only when available (auto probe). The default engine auto-detects Word on Windows/macOS; LibreOffice is the fallback. DO NOT USE if the user only wants to read, summarize, translate or spell-check a text.
---

# Generate an APA 7 academic document (Word + PDF)

Turns a source `.md` (plus its images and layout JSON) into a `.docx` with an institutional cover page, a table of contents / list of tables / list of figures generated with the `docx` library as **real `TOC` fields** (their page numbers are measured in a **double pass**: the first `.docx` is exported to a throwaway PDF with the **resolved engine**, `scripts/paginas-de-pdf.py` reads the real page of every entry, and the second build ships those numbers as the cached field result) and APA 7 references — and exports that same document to PDF with **headless LibreOffice** (the fallback engine). The build's double pass uses whichever engine was resolved in STEP 0 (Word when available on Windows/macOS, LibreOffice otherwise). Both files are delivered.

**Mandatory engine: auto** (prefers Word when available on Windows/macOS), **LibreOffice** fallback.

All paths in this document are **relative to the skill root**
(the folder that contains this `SKILL.md`). Every command below is run from
there, on Windows, macOS or Linux:

```bash
cd "<skill path>"
python scripts/apa7.py <command> [options]
```

`apa7.py` is the only entry point. It runs on Python 3.9 or later, uses the
standard library only, and needs no shell-specific syntax.

**The interpreter name is not fixed.** `python` above stands for whichever
command exists on this machine: `python3` on macOS and most Linux
distributions, `python` where it is present, `py` on Windows. When in doubt,
check the version (`python3 --version`, `python --version` or `py --version`)
and use the one that reports **3.9 or later**; every message `apa7.py` prints
already shows the exact command with the interpreter that is running.

## When it activates

By **intent**, not by exact phrase. It activates when the user wants an academic
document *assembled* in APA, whatever the wording:

| The user says something like | Activates |
|---|---|
| "generate the APA PDF", "documento APA" | yes |
| "armá mi trabajo en APA con portada y referencias" | yes |
| "estructura este informe con norma APA" | yes |
| "pasame este .md a APA, que quede presentable" | yes |
| "formatea esto según APA 7" | yes |
| "solo léeme el .md" / "resúmeme" / "corrígeme la ortografía" | **no** |

It does not activate merely because a `.md` exists in the conversation, nor
when the text is only a loose portion to read, summarize, translate or review.
The boundary is the **fully assembled document**, not APA as a style detail.

When the user asks for the document but mandatory information is missing, **ask
before generating** (see "Mandatory questions" below).

## Required inputs

1. The source **`.md`** (text, headings, markdown tables, references, image markers `![]()`).
2. The **images** referenced in the `.md` (`images/` folder or loose files).
3. The **layout JSON** (`*_content_list.json`, `*_content_list_v2.json`, `*_middle.json`, `*_model.json`, or only some of them) — used only to learn the real size and position of each image and table.

If any is missing, ask for it before continuing. **The `.md` is mandatory; the JSON and images are optional.** The pipeline accepts:

- `.md` only → text document. No figures and no lists of tables/figures.
- `.md` + `images/` → figures at a reasonable fixed size.
- `.md` + images + JSON → figures at their **real size** (from `bbox`) and tables at their real page width.

## STEP 0 — Environment check and installation (MANDATORY)

Before touching any document:

```bash
python scripts/apa7.py check
```

The preflight prints one line per tool: `OK|<tool>|<detail>`,
`INFO|<tool>|<detail>` or `MISSING|<tool>|<detail>`, and ends with `RESULT: OK`
or `RESULT: MISSING` (plus an `ENVIRONMENT OK:` / `MISSING: ...` summary).

- `RESULT: OK` → go to STEP 1.
- `RESULT: MISSING` → run `python scripts/apa7.py install`, which installs what
  is missing and re-checks.

`INFO` means the environment is usable but something is worth knowing: the
interpreter on `PATH` is older than the floor, a pin differs, LibreOffice came
from a non-standard place. It is **not** a failure and never changes the result,
so there is nothing to install because of it.

`install` asks for confirmation before changing the machine, which needs a
terminal. Pass `--yes` to skip the prompt. `--only <tool>` installs a single
tool (repeatable) and `--dry-run` prints the plan and changes nothing, which is
the safe way to inspect it.

If something is still missing after installing, **stop and tell the user**: never
transform a document with an incomplete environment. Details in
`references/system-requirements.md`.

A version that differs from the pin is **not** a failure: the line stays `OK` or
becomes `INFO` and the mismatch is written into its detail
(`[MISMATCH: pinned …]`). Only a genuinely missing tool produces `MISSING` and
exit code 1.

`docx` (npm) and `pymupdf` are pinned in `scripts/lib/rutas.py`. `pymupdf` is a
compiled library, so it is **not** installed into the system Python: the
installer creates `<skill>\.venv` and installs it there, and every script uses
that interpreter. If `.venv` is missing or broken it is recreated automatically.

Environment variables, all optional: `APA7_SOFFICE`, `APA7_PYTHON`, `APA7_NODEDIR`,
`APA7_WORKDIR` and `APA7_SKILL_ROOT` override the resolution. `install` may need
administrator rights: it tries `sudo -n` first so it never blocks an agent, and
prints manual instructions if the password cannot be supplied non-interactively.

## STEP 0.5 — Mandatory questions (cannot be skipped)

The parser (`scripts/md-a-manifiesto.py`) is the missing-data detector. When you
run it, it writes to `MANIFEST.json → diagnostico`:

- `preguntas`: list of objects with `id`, `bloqueante`, `campo`, `tipo`,
  `objetivo`, `indice`, `pregunta`, `opciones` and `como_resolver`.
- `pendientes_bloqueantes`: only the `id`s that prevent delivery.
- `preguntas_texto`: the same questions as plain text, for a quick read.

**Hard rule: if `pendientes_bloqueantes` is not empty, ask the user and do NOT
build the `.docx`.** Do not patch the `.docx` by hand and do not fill the gap
with invented text.

What each part guarantees, because they are not the same:

- **Checked by the machine**: `build-docx.js` aborts with code 4 if a table or
  figure title is missing. That is a subset of the blockers: titles and
  captions, the only ones verifiable without asking.
- **Checked by the agent**: the rest (`portada_*`, `sin_tablas`, `sin_figuras`).
  Nothing in the code prevents building the `.docx` if only those remain, so
  whoever reads `pendientes_bloqueantes` and decides to stop is the agent. If the
  user confirms there are no tables or figures, or that a cover field is
  omitted, it is enough to record it: no extra flag is needed to unblock.
**Already-resolved exception:** `portada_logo`, `portada_docente_titulo` and `portada_vicerrectoria` are asked but do **not** block (`bloqueante: false`). The document is delivered without that line and `verificar-pdf.py` flags each one as a warning. `portada_materia_nrc` is now a **blocking** question — it must be answered before delivery.

### Missing titles and captions

The parser asks a **binary choice** for each table or figure without a title:

> Table 3 has no title. Should I write the title from the report context, or
> will you provide it?

Depending on the answer:

1. **The user provides it** → record it in the `.md` (`Tabla 3. <title>`),
   or pass a JSON with `--titulos-tabla-json` / `--titulos-figura-json`
   (format `{"3": "Title..."}`, 1-based index), and re-run the parser.
2. **The user asks for it to be written** → write the title into that same JSON,
   reading the document context so it truly describes what the table or figure
   shows, and re-run the parser.

Both paths end the same way: the parser stops reporting the missing item. What is
never done is deciding it for the user.

### What is NOT asked

The **source** of tables and figures has a resolved default: `Nota. Elaboración
propia`, unless the `.md` or a JSON indicate otherwise. That is not asked. A
`Nota. …` line that the source `.md` already carries under a table or figure
**wins over the default and is taken verbatim** (`nota_origen: "md"`), including
its punctuation: it is assumed the author wrote it; the generator only prepends
the `Nota.` label when the text does not bring it already.

## Output layout

By default every phase derives its own paths from the document it is handed, so
**you no longer pass `--out`, `--log`, `--outdir` or `--json`** unless you
deliberately want a different location. Everything lands like this, next to the
source `.md`:

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

- The `.docx` and the `.pdf` keep the **original** name of the `.md`. Only the
  working folder is renamed: `NFKD` without diacritics, invalid/control
  characters and spaces become `_`, Windows reserved names get a `_` suffix, and
  it is capped at 60 characters. `Informe técnico.md` → `Informe_tecnico_apa/`.
- **The folder is reused, never versioned.** A second run replaces the files in
  place and adds nothing; rerunning the pipeline on the same `.md` must not
  leave `informe (2).docx` behind. If a folder that already holds an anchor is
  pointed at a *different* `.md`, a warning goes to stderr and the anchor is
  updated.
- **The throwaway PDF of the double pass is deleted** once the second pass
  succeeds, so `logs/` keeps only `paginas.json`. If the second pass fails, that
  PDF **stays** — it is the only record of what was measured. Either way, a
  failed `verify` keeps its PDF and its whole working folder.
- `--carpeta-trabajo <dir>` overrides the working folder for `parse`, `build`,
  `export` and `verify` when the document must live somewhere else.
- **Explicit flags always win.** Passing `--log`, `--outdir`, `--out` or `--json`
  suppresses the corresponding default, so any script or caller that already
  passes them keeps working exactly as before. `export --outdir X` without
  `--carpeta-trabajo` keeps its historical meaning, PDF *and* log in `X`
  (`X/_logs/03-export.log`), which is the only case that still writes `_logs/`.

## Workflow

1. **Read the `.md`** and extract: title, authors, course data, sections with their real heading level, tables, figures and references. Make an **explicit inventory of tables and figures** (how many, with which titles and sources). While reading, **fix glued words** (PDF-conversion artifacts: `2:Diferenciaciónentre Bugs`, `ACTIVIDAD1:MapeodeControles`). The full rule and the exceptions file are in `references/apa7-format.md`.

2. **Build `MANIFEST.json`** by running the parser on the `.md` (add the layout JSON with `--docling-json`, or let it auto-detect a neighbouring `*_content_list.json`):

   ```bash
   python scripts/apa7.py parse --md "<document>.md" [--portada "portada.json"] [--carpeta-trabajo "<dir>"] ...
   ```

   `parse` forwards every option to `scripts/md-a-manifiesto.py`, so `python scripts/apa7.py parse --help` lists the full set. With no `--out` the manifest goes to `<md>_apa/datos/MANIFEST.json`, with no `--log` to `<md>_apa/logs/01-analisis.log`.

   The parser fixes glued words, deduplicates numbering, **computes each image's size from the `bbox` of the layout JSON** (`--docling-json`, or the auto-detected `*_content_list.json`; `*_content_list_v2.json` as an alternative; `*_model.json` has no image path and only serves as a proportion fallback), applies the default table/figure note, and leaves the blocking questions in `diagnostico`. **Do not estimate sizes by eye.** If a figure already declared in the `.md` also appears in the JSON, it is matched by file name and measured, never duplicated. Read `analisis.log` / the console report: it prints the inventory, the cover fields that are missing, the index policy and the pending questions.

3. **Complete the cover page** with `references/institutional-cover.md`: a positional 3-zone layout (title at the top, authors vertically centered, institutional block anchored at the bottom), optional logo **only if the user provides it** (always ask). For every field the `.md` does not carry, **ask the user** whether to add it or omit it. This **always** includes asking for the instructor's title/profession, unless the user says to omit it that time. **General rule: for any value not present in the attached files, ask — do not infer or invent.** The answers go into a `portada.json` passed to the **parser** with `--portada`: it merges them with the deduced ones and writes them into `MANIFEST.json`. `build-docx.js` does not receive that file. Three cover fields are **always asked but never block** — logo, instructor's title and vice-rector's office (`vicerrectoria`): if they stay missing, ask anyway, then deliver; the cover is generated without that line and `verificar-pdf.py` warns. The course/NRC line (`materia_nrc`) is now **blocking**: it must be answered before delivery.

4. **Table and figure attribution**: if the source `.md` does not carry a `Nota. …` line under them, the parser sets `Nota. Elaboración propia` automatically (marked with `nota_origen: "default"`). That is already resolved, it is not asked. A note the `.md` already has is taken verbatim (`nota_origen: "md"`). To change or refine the source of a specific one, pass `--notas-tabla-json` / `--notas-figura-json` with `{"indice": "text"}`.

5. **Build the `.docx`** with `python scripts/apa7.py build --manifiesto "MANIFEST.json"` (it runs `scripts/build-docx.js` with the resolved Node, the `docx` library from `.work` and the layout JSON from the manifest). With no `--out` the `.docx` is written next to the source `.md` with the same name, and with no `--log` the build log goes to `<md>_apa/logs/02-build.log`. Formatting rules are in `references/apa7-format.md`; instructions for functional indices are in `references/word-toc-fields.md`. Critical points already resolved:
   - The TOC is generated **always**; the list of tables and the list of figures **only if the document has them**. If it does not, **ask** to confirm and do **not** generate the empty list.
   - Tables with `width` + `columnWidths` + `layout: FIXED`: LibreOffice does not render tables without defined column widths.
   - Captions with a **real `SEQ` field** (`SEQ Tabla` / `SEQ Figura`), which LibreOffice **does** resolve, because `TOC \c "Tabla"` only collects captions that carry that SEQ. The whole legend (`Tabla 1.`) is **bold** — number included, written by the field's cached run — and the title that follows is italics. Headings carry an applied `outlineLevel` for `TOC \u`.
   - **Two-pass build.** `build` runs node once, then (if the manifest asks for TOC fields and LibreOffice is available) exports to a throwaway PDF in `<md>_apa/logs/`, measures it with `paginas-de-pdf.py` into `paginas.json`, and runs node a second time with `--paginas-json`. Both passes are inside the one `build` command, and the throwaway PDF is deleted once the second pass succeeds. If the venv or LibreOffice is missing, the first `.docx` is kept and Word fills the numbers in on open (<code>updateFields</code>) — a degradation, not a failure. Needs LibreOffice (already mandatory in STEP 0).
   - **Aborts with code 4 if a title or caption is missing.** There is no `(sin título)`: it is a defect, not a text.
   - **Figure note below the image** and with `keepNext`, so the note is not orphaned on the previous page. **Table note AFTER the table**.
   - **Tables with HORIZONTAL borders only**: top border, a line under the header row and bottom border. No verticals and no lines between data rows.
   - **Automatic per-table orientation**: the body is split into stretches and each stretch is a `.docx` section. A table moves on its own to a **landscape page** (with its title and note) when it does not fit in portrait: 6 or more columns, or cells so long that rotating the page genuinely relieves them. The criterion is legibility/width, **not the number of rows**. Afterwards the text returns to portrait. See `references/apa7-format.md`.
   - Level-1 headings break with `pageBreakBefore`, **not** with a loose paragraph carrying a `PageBreak`: that empty paragraph produces a blank page when the previous one is already full.

6. **Export to PDF with LibreOffice**:

   ```bash
   python scripts/apa7.py export --docx "<path>/document.docx" [--carpeta-trabajo "<dir>"]
   ```

   It prints **only** the resulting PDF path on stdout and puts the log in the working folder (`<stem>_apa/logs/03-export.log`), so the path can be captured directly. Progress and warnings go to stderr. The PDF goes **next to the `.docx`**, with the same name. With no `--docx` at all, `export` looks in the **current folder** (the one that holds the `.md` and the `.docx`, *not* the working folder) and takes its single `.docx`; with zero or two or more it says so instead of guessing. With `--outdir` and no `--carpeta-trabajo` the old behaviour is kept verbatim (PDF in `--outdir`, log in `--outdir/_logs/`). It works on a temporary copy with an **isolated LibreOffice profile** per run, which is deleted at the end; because of that isolation a LibreOffice the user has open is **not** a problem and is left alone. Only the skill's **own** leftovers (the isolated `lo_profile` runs) are closed. `--cerrar-libreoffice` closes **every** LibreOffice process first and warns that unsaved documents are lost. The `Could not find platform independent libraries <prefix>` warning on stderr is benign. Do **not** launch `soffice` through a shell background operator or `Start-Process`: the child inherits the pipe handle and the reader hangs (see `references/word-toc-fields.md`).

7. **Verify before delivering** with `python scripts/apa7.py verify --pdf "document.pdf" --manifiesto "MANIFEST.json"`: cover in 3 zones and members in a single paragraph (**this is checked, as a warning**), indices with the correct page number, captions with number and title, tables with content, **each table with its note below it in the PDF**, **each figure note below its image**, references with hanging indent. With no `--json` the report goes to `<md>_apa/datos/verificacion.json`. `verify` picks the `.venv` interpreter on its own, because that is where `pymupdf` lives. The cover warnings (logo, instructor title, course/NRC or vice-rector's office not asked about) **are not failures**: they are reported to the user and the document is delivered. If a critical check fails, fix and export again. **A failed verification never deletes anything**: the PDF and the working folder stay on disk to be inspected.

8. **Deliver both files** (`.docx` and `.pdf`), and mention the working folder so the user knows the intermediate files are there and can be deleted.

## Rules already defined (do not ask again)

- It activates by **intent**: asking for an academic document assembled in APA, no need to say the words "APA PDF" (see "When it activates").
- Export engine: headless LibreOffice (the fallback). The default engine is **auto**, which probes for Word on Windows/macOS; Word is used only when available and proven usable.
- **STEP 0 mandatory** before transforming.
- Missing cover field → ask (add or omit), including the instructor's title.
- **Cover in 3 zones**: title at the top (optional logo **above** the title), members vertically centered, institutional block anchored at the bottom. Built with a 3-row borderless table, not with filler paragraphs. **Logo**: only if the user provides it, but **always ask**: if it was not asked about, the verifier flags it.
- **Any missing value → ask (do not infer or invent).** Only values that appear in the attached files count as "existing data". Already-resolved exception: source of tables/figures → `Nota. Elaboración propia` by default, and a `Nota. …` line already present in the `.md` wins over that default and is kept verbatim.
- **Optional cover fields that never block:** the logo, the instructor's title/profession and the vice-rector's office (`vicerrectoria`). They are always asked about, but if they end up missing the cover is generated without that line and delivery advances while `verificar-pdf.py` warns. The course/NRC line (`materia_nrc`) is now **blocking**: it must be answered before delivery. The rest of the cover fields (title, members, faculty, instructor, date) **block delivery** while unanswered.
- **Table/figure title or caption → ask** (should it be written from context or provided by the user?) and **block delivery** until answered. A document with blank titles is not emitted.
- **Figure note below the image** (generator default); **table note below the table** (APA 7 distinguishes them by position, not only by text).
- **Inventory of tables/figures**: the TOC always; lists of tables and figures only if they exist. If the analysis does not find them, **ask** before assuming.
- Format: letter size and APA 7th edition margins, with no extra institutional rules.
- References: always normalized to APA format.
- Glued words: always fixed (missing spaces, "2:Word", "ACTIVIDAD1:"). Pure UPPERCASE acronyms (`OWASP`, `CVSS`) are never split, because the rule needs lowercase followed by uppercase. **Protected automatically, never split and never reported:** URLs, DOIs, e-mails and file names with an extension.
- Cover author list: comma-separated, with "y" before the last one, regardless of count. `autores` accepts `"Nombre Apellido y Nombre Apellido"` or `["Nombre Apellido", "Nombre Apellido"]`; the text is kept as-is, not split.
- **Reviewing the deglue is part of the job.** The parser ships no whitelist of untouchable terms: it carries no data about a document it has not read. The report splits the changes into two queues, both answered in the source `.md`:
  - **Clear** corrections: an obvious missing space (`ACTIVIDAD4` → `ACTIVIDAD 4`).
  - **Review** (ambiguous): mixed alphanumeric or camelCase (`JavaScript` → `Java Script`, `SHA256` → `SHA 256`), i.e. it could well be a real product name or variable.
  For every ambiguous one, do NOT accept blindly: either add the term to a protected-terms file (one per line, exact, case-sensitive) and re-run with `--terminos-protegidos <ARCHIVO>` so it stays intact, or ask the user. Also review the **residue report** (stretches the deglue could not separate with confidence). Fix a real term in the `.md` and re-run the parser. Never patch the generated `.docx` by hand.

## Skill files

| File | Purpose |
|---|---|
| `SKILL.md` | This document. |
| `references/apa7-format.md` | APA 7 formatting rules (page, typography, heading levels, tables/figures, glued words, references). |
| `references/institutional-cover.md` | 3-zone cover layout, fields, and what to do when information is missing. |
| `references/word-toc-fields.md` | How to make the TOC and the indices functional (real `TOC` fields, outline levels, `SEQ` captions, the two-pass page measurement, `updateFields`), `docx` library pitfalls, and how to invoke LibreOffice without hanging it. |
| `references/system-requirements.md` | Requirements, preflight, installation and path resolution. |
| `scripts/apa7.py` | **The entry point.** `check`, `install`, `export`, and `parse` / `build` / `verify` as thin dispatchers that forward their options to the scripts below. |
| `scripts/md-a-manifiesto.py` | `.md` + layout JSON → `MANIFEST.json`. Enriches, deduplicates, deglues, computes sizes (from the JSON `bbox` when present), applies the default note and leaves the blocking questions in `diagnostico`. Reached through `apa7.py parse`; flags: `--md`, `--out`, `--log`, `--base-dir`, `--portada`, `--sin-deglue`, `--terminos-protegidos`, `--titulos-tabla-json`, `--titulos-figura-json`, `--notas-tabla-json`, `--notas-figura-json`, `--detectar-niveles-en-linea`, `--sin-indice-tablas`, `--sin-indice-figuras`, `--docling-json`, `--pdf`, `--ancho-max-tabla` (run `apa7.py parse --help` for the full list). |
| `scripts/build-docx.js` | `MANIFEST.json` → `.docx` (cover, indices as real `TOC` fields, tables, figures, references). Splits the body into sections and moves tables that do not fit in portrait to a landscape page. Accepts `--paginas-json` to cache the index page numbers (second pass). **Aborts with code 4 if a title or caption is missing.** Reached through `apa7.py build`. |
| `scripts/paginas-de-pdf.py` | Reads the throwaway PDF of the first pass and writes the real 1-based page of each section/table/figure, for the second `--paginas-json` build. Reached through `apa7.py build`; run with `--pdf`, `--manifiesto`, `--out`. |
| `scripts/verificar-pdf.py` | Verifies the `.pdf` (pymupdf): cover, indices, captions, indents, table notes below and figure notes below the image, plus that the index page numbers match the real pages. Reached through `apa7.py verify`. |
| `scripts/lib/rutas.py` | Portable resolution of paths, Python, Node and LibreOffice, the version pins, and the isolated-profile LibreOffice runner with its per-profile process cleanup. Also the single source of truth for **where a document's files go**: the working folder name (`<md>_apa/`), `datos/`, `logs/`, the `fuente.json` anchor and the default `--out`/`--log`/`--json`/`--outdir` for each phase. Standard library only. |
| `scripts/lib/instalador.py` | Installation plans per package manager and the non-interactive elevation. |
| `scripts/lib/fuentes.py` | Which embedded font names the verifier accepts: Times New Roman and its metric-compatible substitutes. |
| `scripts/tests/` | Unit tests, run with `python -m unittest discover -s scripts/tests -t scripts/tests`. Setting `APA7_TEST_WORD_REAL=1` adds the opt-in integration test that launches the real Word driver once (`TestWordReal`); without it, nothing in the suite ever starts Word. |
| `.work/` | Generated state: `docx`'s `node_modules`. Not source code; it is regenerated by `apa7.py install`. |
| `.venv/` | Generated state: the Python virtual environment holding the pinned `pymupdf`. Not source code; it is regenerated by `apa7.py install`. |
