---
name: generar-pdf-apa
description: Builds or structures an academic Word (.docx) and PDF document following APA 7th edition, from a source Markdown (.md) file, its images and the layout JSON files (content_list, content_list_v2, middle, model) that carry figure size and position. ALWAYS USE when the user asks to assemble or produce an academic APA document or PDF, formalize a paper, set it up with an institutional cover page, table of contents, list of tables, list of figures and APA references, or apply APA formatting to a text. Activation is by INTENT, not keywords: "generate the APA PDF", "armá mi trabajo en formato APA", "estructura este informe con norma APA", "pasame esto a APA con portada y referencias" all apply. Also use when required data is missing and the user must be asked before continuing. Exports the PDF with headless LibreOffice and does NOT require Microsoft Word. DO NOT USE if the user only wants to read, summarize, translate or spell-check a text.
---

# Generate an APA 7 academic document (Word + PDF)

Turns a source `.md` (plus its images and layout JSON) into a `.docx` with an institutional cover page, a table of contents / list of tables / list of figures generated with the `docx` library (entries with page numbers via `PAGEREF` fields) and APA 7 references — and exports that same document to PDF with **headless LibreOffice**, the only conversion engine. Both files are delivered.

**Mandatory engine: LibreOffice. Microsoft Word is NOT a requirement.**

All paths in this document are **relative to the skill root**
(the folder that contains this `SKILL.md`). In PowerShell, move into it:

```powershell
cd "<skill path>"
```

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

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\comprobar-entorno.ps1"
```

The preflight prints one line per tool: `OK|<tool>|<detail>` or
`MISSING|<tool>|<detail>`, and ends with `RESULT: OK` or `RESULT: MISSING`
(plus an `ENVIRONMENT OK:` / `MISSING: ...` summary).

- `RESULT: OK` → go to STEP 1.
- `RESULT: MISSING` → run `powershell -ExecutionPolicy Bypass -File "scripts\instalar-entorno.ps1"`, which installs what is missing (winget/brew, npm, pip) and re-checks.

If something is still missing after installing, **stop and tell the user**: never transform a document with an incomplete environment. Details in `references/system-requirements.md`.

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
  **Already-resolved exception:** `portada_logo` and `portada_docente_titulo`
  are asked but do **not** block (`bloqueante: false`). The document is delivered
  without them and `verificar-pdf.py` flags it as a warning.

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
propia`, unless the `.md` or a JSON indicate otherwise. That is not asked.

## Workflow

1. **Read the `.md`** and extract: title, authors, course data, sections with their real heading level, tables, figures and references. Make an **explicit inventory of tables and figures** (how many, with which titles and sources). While reading, **fix glued words** (PDF-conversion artifacts: `2:Diferenciaciónentre Bugs`, `ACTIVIDAD1:MapeodeControles`). The full rule and the exceptions file are in `references/apa7-format.md`.

2. **Build `MANIFEST.json`** by running the parser on the `.md` (add the layout JSON with `--docling-json`, or let it auto-detect a neighbouring `*_content_list.json`):

   ```powershell
   <python> "scripts\md-a-manifiesto.py" --md "<document>.md" --out "MANIFEST.json" [--log "<log>"] [--portada "portada.json"] ...
   ```

   The parser fixes glued words, deduplicates numbering, **computes each image's size from the `bbox` of the layout JSON** (`--docling-json`, or the auto-detected `*_content_list.json`; `*_content_list_v2.json` as an alternative; `*_model.json` has no image path and only serves as a proportion fallback), applies the default table/figure note, and leaves the blocking questions in `diagnostico`. **Do not estimate sizes by eye.** If a figure already declared in the `.md` also appears in the JSON, it is matched by file name and measured, never duplicated. Read `analisis.log` / the console report: it prints the inventory, the cover fields that are missing, the index policy and the pending questions.

3. **Complete the cover page** with `references/institutional-cover.md`: a positional 3-zone layout (title at the top, authors vertically centered, institutional block anchored at the bottom), optional logo **only if the user provides it** (always ask). For every field the `.md` does not carry, **ask the user** whether to add it or omit it. This **always** includes asking for the instructor's title/profession, unless the user says to omit it that time. **General rule: for any value not present in the attached files, ask — do not infer or invent.** The answers go into a `portada.json` passed to the **parser** with `--portada`: it merges them with the deduced ones and writes them into `MANIFEST.json`. `build-docx.js` does not receive that file.

4. **Table and figure attribution**: if the `.md` does not indicate a different source, the parser sets `Nota. Elaboración propia` automatically (marked with `nota_origen: "default"`). That is already resolved, it is not asked. To change or refine the source of a specific one, pass `--notas-tabla-json` / `--notas-figura-json` with `{"indice": "text"}`.

5. **Build the `.docx`** with `scripts/build-docx.js` (the `docx` npm library), reading `MANIFEST.json`. Formatting rules are in `references/apa7-format.md`; instructions for functional indices are in `references/word-toc-fields.md`. Critical points already resolved:
   - The TOC is generated **always**; the list of tables and the list of figures **only if the document has them**. If it does not, **ask** to confirm and do **not** generate the empty list.
   - Tables with `width` + `columnWidths` + `layout: FIXED`: LibreOffice does not render tables without defined column widths.
   - Captions with a **literal number** ("Tabla 1"), not a `SEQ` field (LibreOffice does not resolve `SEQ`).
   - **Aborts with code 4 if a title or caption is missing.** There is no `(sin título)`: it is a defect, not a text.
   - **Figure note BEFORE the image** and with `keepNext`, so the note is not orphaned on the previous page. **Table note AFTER the table**.
   - **Tables with HORIZONTAL borders only**: top border, a line under the header row and bottom border. No verticals and no lines between data rows.
   - **Automatic per-table orientation**: the body is split into stretches and each stretch is a `.docx` section. A table moves on its own to a **landscape page** (with its title and note) when it does not fit in portrait: 6 or more columns, or cells so long that rotating the page genuinely relieves them. The criterion is legibility/width, **not the number of rows**. Afterwards the text returns to portrait. See `references/apa7-format.md`.
   - Level-1 headings break with `pageBreakBefore`, **not** with a loose paragraph carrying a `PageBreak`: that empty paragraph produces a blank page when the previous one is already full.

6. **Export to PDF with LibreOffice**:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "scripts\export-pdf.ps1" -Docx "<path>\document.docx" [-OutDir "<folder>\deliverable"]
   ```

   It works on a temporary copy with an **isolated LibreOffice profile** per run, which is deleted at the end. The `Could not find platform independent libraries <prefix>` warning on stderr is benign. Do **not** switch to `soffice.exe` or `Start-Process`: they hang the process (see `references/word-toc-fields.md`).

7. **Verify before delivering** with `scripts/verificar-pdf.py` (uses the Python returned by `Get-PythonPath`, with `pymupdf`): cover in 3 zones and members in a single paragraph (**this is checked, as a warning**), indices with the correct page number, captions with number and title, tables with content, **each table with its note below it in the PDF**, **each figure note above its image**, references with hanging indent. The cover warnings (logo or instructor title not asked about) **are not failures**: they are reported to the user and the document is delivered. If a critical check fails, fix and export again.

8. **Deliver both files** (`.docx` and `.pdf`).

## Rules already defined (do not ask again)

- It activates by **intent**: asking for an academic document assembled in APA, no need to say the words "APA PDF" (see "When it activates").
- Export engine: **headless LibreOffice, only**. Word is not a requirement and is not used.
- **STEP 0 mandatory** before transforming.
- Missing cover field → ask (add or omit), including the instructor's title.
- **Cover in 3 zones**: title at the top (optional logo **above** the title), members vertically centered, institutional block anchored at the bottom. Built with a 3-row borderless table, not with filler paragraphs. **Logo**: only if the user provides it, but **always ask**: if it was not asked about, the verifier flags it.
- **Any missing value → ask (do not infer or invent).** Only values that appear in the attached files count as "existing data". Already-resolved exception: source of tables/figures → `Nota. Elaboración propia` by default.
- **Table/figure title or caption → ask** (should it be written from context or provided by the user?) and **block delivery** until answered. A document with blank titles is not emitted.
- **Figure note above the image; table note below the table** (APA 7 distinguishes them by position, not only by text).
- **Inventory of tables/figures**: the TOC always; lists of tables and figures only if they exist. If the analysis does not find them, **ask** before assuming.
- Format: letter size and APA 7th edition margins, with no extra institutional rules.
- References: always normalized to APA format.
- Glued words: always fixed (missing spaces, "2:Word", "ACTIVIDAD1:"), respecting proper nouns and technical terms protected by `references/terms-whitelist.txt`.
- Cover author list: comma-separated, with "y" before the last one, regardless of count. `autores` accepts `"Nombre Apellido y Nombre Apellido"` or `["Nombre Apellido", "Nombre Apellido"]`; the text is kept as-is, not split.
- **Extending the whitelist is part of the job.** The list carries generic terms, not the document's proper nouns. The stretches the deglue cannot split with confidence (brand or project names glued to the rest of the sentence) come out in the review queue: review them by hand in the `.md` and, those that are a real term, **add them to `references/terms-whitelist.txt` and re-run the parser**. Without that step the deglue splits brands over terms nobody protects.

## Skill files

| File | Purpose |
|---|---|
| `SKILL.md` | This document. |
| `references/apa7-format.md` | APA 7 formatting rules (page, typography, heading levels, tables/figures, glued words, references). |
| `references/institutional-cover.md` | 3-zone cover layout, fields, and what to do when information is missing. |
| `references/word-toc-fields.md` | How to make the TOC and the indices functional (`PAGEREF`, `updateFields`), `docx` library pitfalls, and how to invoke LibreOffice without hanging it. |
| `references/system-requirements.md` | Requirements, preflight, installation and path resolution. |
| `references/terms-whitelist.txt` | Untouchable terms for the deglue. **Single source**: there is no list embedded in the code. |
| `scripts/md-a-manifiesto.py` | `.md` + layout JSON → `MANIFEST.json`. Enriches, deduplicates, deglues, computes sizes (from the JSON `bbox` when present), applies the default note and leaves the blocking questions in `diagnostico`. Flags: `--md`, `--out`, `--log`, `--base-dir`, `--portada`, `--whitelist`, `--sin-deglue`, `--titulos-tabla-json`, `--titulos-figura-json`, `--notas-tabla-json`, `--notas-figura-json`, `--detectar-niveles-en-linea`, `--sin-indice-tablas`, `--sin-indice-figuras`, `--docling-json`, `--pdf`, `--ancho-max-tabla` (run with `--help` for the full list). |
| `scripts/build-docx.js` | `MANIFEST.json` → `.docx` (cover, indices, tables, figures, references). Splits the body into sections and moves tables that do not fit in portrait to a landscape page. **Aborts with code 4 if a title or caption is missing.** |
| `scripts/export-pdf.ps1` | `.docx` → `.pdf` with headless LibreOffice. |
| `scripts/verificar-pdf.py` | Verifies the `.pdf` (pymupdf): cover, indices, captions, indents, table notes below and figure notes above the image. |
| `scripts/lib/rutas.ps1` | Portable resolution of paths, Python, Node and LibreOffice. |
| `scripts/comprobar-entorno.ps1` | STEP 0 preflight. |
| `scripts/instalar-entorno.ps1` | Installs what is missing and re-checks. |
| `.work/` | Generated state: `docx`'s `node_modules`. Not source code; it is regenerated by `instalar-entorno.ps1`. |
