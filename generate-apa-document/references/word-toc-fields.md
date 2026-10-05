# Tabla de contenido, indice de tablas e indice de figuras

How to build the indexes so that they come out **functional**, with real page
numbers, and how to export the result to PDF. This file exists so we don't
trip over the same already-fixed problems again.

> The filename is historic: the conversion is no longer done by Word. The
> engine is LibreOffice and everything here applies to the pipeline with
> `docx` (npm) and LibreOffice.

## The golden rule

LibreOffice evaluates the fields it finds, but it **does not build an empty
`TOC` field from scratch** (it ignores the update-fields option for that). A
`TOC` field with no computed result exports **empty**. A `TOC` field whose
result is already cached in the document, on the other hand, exports **exactly
as cached**. So the indexes are real `TOC` fields, and their result is filled in
by a two-pass build:

1. `build` creates the `.docx` once (indexes without page numbers).
2. `apa7.py` exports that `.docx` to a throwaway PDF with LibreOffice, and
   `scripts/paginas-de-pdf.py` reads the real page of every entry from it.
3. `build-docx.js` runs a second time with `--paginas-json`, and that map
   becomes the **cached result** of the `TOC` fields.

The double pass lives inside `python scripts/apa7.py build`, so the documented
pipeline does not change. If the venv or LibreOffice is missing, the first
`.docx` is kept and Word fills the numbers in when the document is opened
(`features.updateFields`); it is a degradation, not a failure.

Verified against LibreOffice 26.8 on Windows with `docx` 9.7.1:

| Test | Result |
|------|--------|
| `TOC` field with no result | exports **empty** (LO does not build it) |
| `TOC` field whose `contentChildren` are PAGEREF lines | entries render, the nested PAGEREF is **not** resolved |
| `TOC` field with `cachedEntries` (literal pages) | renders correctly, with dot leaders |
| `SEQ` field in a caption / body | **resolved** (`Tabla ` + `SEQ Tabla` -> `Tabla 1`) |
| `TOC \u` / `TOC \o` / `TOC \c "Tabla"` instrText | emitted as real fields |
| `doc.get_toc()` after export | PDF outline from the applied outline levels, with correct pages |

Consequences that must not be undone:

- Each **heading** carries an applied **outline level** (`outlineLevel`): that
  is what `TOC \u` collects and what LibreOffice turns into PDF bookmarks, which
  is how the double pass learns each section's page.
- Each **caption** carries a real **`SEQ` field** (`SEQ Tabla`, `SEQ Figura`),
  not a literal number: `TOC \c "Tabla"` only collects captions with that SEQ.
- The `PAGEREF` recipe that used to build the TOC by hand is **gone**: a
  hand-made list is not a field and Word cannot update it.

## Recipes that work

### 1. Heading with outline level and bookmark

Each heading carries an applied `outlineLevel` (0-based) and a bookmark. The
outline level is what `TOC \u` collects and what LibreOffice exports as a PDF
bookmark, which is how the double pass measures the section's page:

```js
new Paragraph({
  outlineLevel: nivel - 1,          // nivel 1..5 -> 0..4
  children: [new Bookmark({
    id: 'apa_sec_' + hid,
    children: runsDe(segmentos, estilo),
  })],
});
```

### 2. Table of contents as a real `TOC` field

The `TableOfContents` class emits a real `TOC` field (`w:instrText`), with the
result cached. `useAppliedParagraphOutlineLevel: true` adds `\u`:

```js
const entradas = M.secciones.map((s) => ({
  title: s.texto, level: s.nivel,
  page: paginaDe('secciones', s.hid),   // from --paginas-json, may be undefined
  href: 'apa_sec_' + s.hid,
}));

new TableOfContents('Contenido', {
  hyperlink: true,
  useAppliedParagraphOutlineLevel: true, // -> TOC \h \u
  beginDirty: true,
  cachedEntries: entradas,
});
```

`cachedEntries` are rendered by the library as `TOC1..TOC5` paragraphs with a
right dot-leader tab stop and the literal page. Those styles are declared in the
`Document` so the indexes stay Times 12, double spaced:

```js
styles: {
  default: { document: { run: { font: 'Times New Roman', size: 24 },
                         paragraph: { spacing: { line: 480 } } } },
  paragraphStyles: Array.from({ length: 5 }, (_, i) => ({
    id: 'TOC' + (i + 1), name: 'TOC ' + (i + 1), quickFormat: true,
    run: { font: 'Times New Roman', size: 24, bold: i === 0 },
    paragraph: { spacing: { line: 480 }, indent: { left: i * 720 } },
  })),
  characterStyles: [{ id: 'IndexLink', name: 'Index Link',
                      run: { font: 'Times New Roman', size: 24, color: '000000' } }],
},
```

`features: { updateFields: true }` writes `w:updateFields` into the `.docx`, so
Word offers to recalculate the fields on open. LibreOffice, on export, uses the
cached result instead of rebuilding.

### 3. Indice de tablas y de figuras

Same field, different switch: `captionLabelIncludingNumbers` adds `\c "Tabla"`
(`TOC \c "Tabla" \h`). It collects captions that carry a `SEQ Tabla` field, so
the caption number must be a real SEQ, not a literal:

```js
// caption label in the body
new Bookmark({ id: 'apa_tbl_' + n, children: [
  new TextRun({ text: 'Tabla ', bold: true }),
  new SimpleField(' SEQ Tabla \\* ARABIC ', String(n)),   // cached -> "Tabla 1"
]})

// list of tables
new TableOfContents('Tablas', {
  hyperlink: true,
  captionLabelIncludingNumbers: 'Tabla',   // -> TOC \c "Tabla" \h
  beginDirty: true,
  cachedEntries: tablasEntradas,
});
```

Figures use `SEQ Figura` / `captionLabelIncludingNumbers: 'Figura'`.

The caption number and the title stay in **separate paragraphs**, which is what
APA asks for (bold number on its own line, italic title next). Because of that,
a user who updates the index fields **in Word** will get only the numbers in the
lists of tables/figures (Word's `\c` collects the caption paragraph); the
cached result, which is what the PDF ships, keeps the full "Tabla N. Titulo".

**The index is only generated if there are items to index.** If the document
has no tables, no table index is created; if it has no figures, no figure index
is created. An empty index is deleted.

### 4. Tables with explicit width

**LibreOffice does not render the content of a table without column widths**:
the table comes out empty. It is a silent failure, the `.docx` opens fine and
nothing appears inside.

```js
new d.Table({
  rows: rowsOut,
  width: { size: usable, type: d.WidthType.DXA },
  columnWidths,
  layout: d.TableLayoutType.FIXED,
  borders: { /* solo lineas horizontales */ },
})
```

- `usable` is the usable width of the section according to its margins and
  orientation.
- The first column (the criteria one) is usually wider; the rest is split
  evenly.
- Borders: only top, bottom and `insideHorizontal`.

The total width comes from the `bbox` of the table in the MinerU layout JSON,
not from eyeballing it.

### 5. Pitfalls of the `docx` library

- If a helper returns `[para]` and it is inserted as `[ helper(), tabla, nota ]`,
  `Array.prototype.flat()` **flattens only one level**: the nested array is
  serialized as `<0/>` and **the document does not open**. Use spread
  (`[...helper(), tabla, nota]`) or `.flat(2)`.
- The library's `TableOfContents` **is** used, but only with `cachedEntries`:
  the empty `TOC` field exports empty. The cached entries are the two-pass
  result and must be filled in (otherwise the field has no result to export).
- `spacing` in a paragraph style goes in `paragraph.spacing`.
- The switch mapping comes from `index.cjs` (`captionLabel` -> `\a "X"`,
  `captionLabelIncludingNumbers` -> `\c "X"`, `tcFieldIdentifier` -> `\f`,
  `headingStyleRange` -> `\o`, `useAppliedParagraphOutlineLevel` -> `\u`,
  `hyperlink` -> `\h`). The instruction order puts `\c`/`\a`/`\b`/`\d`/`\f`
  before `\h` and `\l`/`\n`/`\o`/`\p` after it, which is why the emitted
  instrText reads `TOC \c "Tabla" \h` and `TOC \h \u`.

## PDF export with LibreOffice

Command: `python scripts/apa7.py export --docx <docx>`. It works on a temporary
copy and creates an isolated user profile per run, so it does not clash with a
LibreOffice session opened by the user.

```python
subprocess.call([soffice,
                 "--headless", "--norestore", "--nolockcheck",
                 "--nofirststartwizard",
                 "-env:UserInstallation=file:///<tmp>/lo_profile_<id>",
                 "--convert-to", "pdf:writer_pdf_Export",
                 "--outdir", tmp_dir, docx],
                stdout=log_file, stderr=log_file)
```

**How it is invoked, and why not any other way.** Contradicting this part makes
the script hang:

- **Redirect stdout and stderr to files.** A child that inherits the caller's
  pipe handle can keep `soffice` alive after the parent has finished, so a
  reader on that pipe never sees EOF. That is what used to leave the pipeline
  hanging with no output. The files are read after the process ends, then
  deleted; the log goes to `<md>_apa/logs/03-export.log`. Passing `--outdir`
  without `--carpeta-trabajo` keeps the historical behaviour instead (PDF and
  log together in `<outdir>/_logs/03-export.log`).
- **Do not use a shell background operator or `Start-Process`.** A nested
  redirection on top of an already-redirected parent leaves LibreOffice never
  finishing.
- **Do not switch to `soffice.exe` on Windows.** The `.exe` detaches; on
  Windows the console wrapper `soffice.com` is what waits for the conversion.
  On macOS and Linux the binary is just `soffice`.
- **The warning `Could not find platform independent libraries <prefix>`** on
  stderr is **benign** and must not be treated as an error.
- **Clean only our own leftovers.** Before and after the run, the skill closes
  the LibreOffice processes whose command line carries the isolated
  `lo_profile`, i.e. its own interrupted runs (a stale `soffice.bin` holds the
  profile lock and makes the next run do nothing). The user's LibreOffice is
  **not** touched: the isolated profile already makes coexistence safe. If a
  foreign instance really has to be closed, `export --cerrar-libreoffice` does
  it and warns that unsaved documents are lost. The run also has a hard timeout
  (`--timeout`, 300 s) reported as exit code 124 rather than as a silent
  failure, and on timeout only the process tree it launched is killed.

## Measuring the pages (the double pass)

`scripts/paginas-de-pdf.py` reads the throwaway PDF and writes the page map that
the second `build` consumes. Invoked by `apa7.py build` it is called **without**
`--out`, so the map lands in the working folder as `<md>_apa/logs/paginas.json`:

```bash
python scripts/paginas-de-pdf.py --pdf <md>_apa/logs/<stem>.pdf \
    --manifiesto <manifiesto.json>
```

The throwaway PDF itself is **deleted once the second pass succeeds**, because it
is a by-product rather than evidence; if the second pass fails it is **kept**,
since it is then the only record of what was actually measured.

- **Sections**: `doc.get_toc()` returns the PDF outline built from the applied
  outline levels, with **1-based** pages. Each manifest section is matched to
  the next outline entry whose normalized title matches; a title that is missing
  from the outline does not shift the following sections.
- **Tables and figures**: the first `Tabla N` / `Figura N` at or after the start
  of the body. A line that starts with `Tabla 1. ...` in the list of tables is
  inside the front matter, so it is ignored.
- **Body start**: first page after the index pages, or the page before the first
  section as a fallback.
- If the manifest or PyMuPDF is missing it still exits 0 with an empty map; only
  a missing PDF is an error (exit 1). An empty map degrades to Word filling the
  numbers on open.

## Verification checklist

On the resulting PDF:

- [ ] Cover page with the fields that exist; no visible page number.
- [ ] Tabla de contenido: page numbers that match the real page of each title.
- [ ] Indices de tablas y de figuras: entries "Tabla N  Titulo...  pagina"
      with the correct number, and only if the document has those elements.
- [ ] Tables with content (an empty landscape page is the symptom of the
      `columnWidths` problem).
- [ ] Captions with number (bold, own paragraph) and italic title (next
      paragraph).
- [ ] Referencias with hanging indent.
- [ ] No blank pages between the table of contents and the body.
