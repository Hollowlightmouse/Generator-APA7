# Tabla de contenido, indice de tablas e indice de figuras

How to build the indexes so that they come out **functional**, with real page
numbers, and how to export the result to PDF. This file exists so we don't
trip over the same already-fixed problems again.

> The filename is historic: the conversion is no longer done by Word. The
> engine is LibreOffice and everything here applies to the pipeline with
> `docx` (npm) and LibreOffice.

## The golden rule

LibreOffice recalculates the page numbers when it opens the document, but it
**does not build an empty `TOC` field from scratch** (it ignores the
update-fields option for that) and **it does not resolve `SEQ` fields**. So:

- The **table of contents** is generated as **literal entries**: the heading
  text, a dot-leader tab and the page number through a `PAGEREF` field.
- Table and figure **captions** carry a **literal number** ("Tabla 1"), not a
  `SEQ` field.
- The **page numbers** in the table of contents and in the indexes are real
  `PAGEREF` fields, which LibreOffice does resolve (bookmark -> page). That is
  why the indexes stay live: if the content moves to a different page, the
  number updates when the document is reopened.

## Recipes that work

### 1. Bookmarks on each heading

After parsing the markdown, each heading is numbered according to its order of
appearance and its run is wrapped in a `_H<id>` bookmark:

```js
case 'h1': return para({ heading: d.HeadingLevel.HEADING_1, /* ... */ },
  [new d.Bookmark({ id: '_H' + b.hid,
     children: [new d.TextRun({ text: b.text, bold: true })] })]);
```

The `_H` prefix is required because Word and LibreOffice do not accept bookmark
names that start with a digit.

### 2. Table of contents entry with PAGEREF

Use the real `TOC 1` / `TOC 2` paragraph styles declared in the `Document` (so
that LibreOffice and Word read the APA format) and a right-aligned dot-leader
tab stop:

```js
function tocEntry(entry) {
  return para({ style: entry.level === 1 ? 'TOC1' : 'TOC2',
                spacing: { before: 0, after: 0, line: 480 },
                tabStops: [{ type: d.TabStopType.RIGHT, position: 9350,
                             leader: d.LeaderType.DOT }],
                alignment: d.AlignmentType.LEFT },
    [new d.TextRun({ text: entry.text }),
     new d.TextRun({ children: [new d.Tab()] }),
     new d.SimpleField('PAGEREF _H' + entry.hid + ' \\h')]);
}
```

Styles in the document, with `spacing` inside **`paragraph`** and not at the
style root level (at root level the library fails):

```js
features: { updateFields: true },
styles: {
  default: { document: { run: { font: 'Times New Roman', size: 24 } } },
  paragraphStyles: [
    { id: 'TOC1', name: 'TOC 1', basedOn: 'Normal',
      run: { font: 'Times New Roman', size: 24, bold: false, color: '000000' },
      paragraph: { spacing: { line: 480 } } },
  ],
},
```

`features: { updateFields: true }` is what makes LibreOffice evaluate the
existing fields on open. It is not a LibreOffice pass of its own: what it does
is write `w:updateFields` into the `.docx`.

### 3. Indice de tablas y de figuras

Same scheme as the table of contents: a bookmark on the caption number, literal
number and title, and `PAGEREF` for the page.

```js
new d.Bookmark({ id: '_Tabla' + n,
  children: [new d.TextRun({ bold: true, children: ['Tabla ' + n] })] })
```

```js
para({ spacing: { before: 0, after: 0, line: 480 },
       alignment: d.AlignmentType.LEFT,
       tabStops: [{ type: d.TabStopType.RIGHT, position: 9350,
                    leader: d.LeaderType.DOT }] },
  [new d.TextRun({ children: ['Tabla ' + n, '  ', title] }),
   new d.TextRun({ children: [new d.Tab()] }),
   new d.SimpleField('PAGEREF _Tabla' + n + ' \\h')])
```

The caption goes in **a single paragraph** with a line break between the number
and the title: that way the number stays bold on its own line and the title is
italic on the next one, which is what APA asks for.

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
- The library's `TableOfContents` is not useful here: LibreOffice ignores the
  empty `TOC` field. The literal entry from recipe 2 is used instead.
- `spacing` in a paragraph style goes in `paragraph.spacing`.

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
  deleted; the log is kept in `<outdir>/_logs/03-export.log`.
- **Do not use a shell background operator or `Start-Process`.** A nested
  redirection on top of an already-redirected parent leaves LibreOffice never
  finishing.
- **Do not switch to `soffice.exe` on Windows.** The `.exe` detaches; on
  Windows the console wrapper `soffice.com` is what waits for the conversion.
  On macOS and Linux the binary is just `soffice`.
- **The warning `Could not find platform independent libraries <prefix>`** on
  stderr is **benign** and must not be treated as an error.
- **Kill leftovers.** Before and after the run, any LibreOffice still running is
  terminated, because a stale instance holds the profile lock and the next run
  exits without converting anything. The run also has a hard timeout
  (`--timeout`, 300 s) reported as exit code 124 rather than as a silent
  failure.

## Verification checklist

On the resulting PDF:

- [ ] Cover page with the fields that exist; no visible page number.
- [ ] Tabla de contenido: page numbers that match the real page of each title.
- [ ] Indices de tablas y de figuras: entries "Tabla N  Titulo...  pagina"
      with the correct number, and only if the document has those elements.
- [ ] Tables with content (an empty landscape page is the symptom of the
      `columnWidths` problem).
- [ ] Captions with number and title, in a single paragraph.
- [ ] Referencias with hanging indent.
- [ ] No blank pages between the table of contents and the body.
