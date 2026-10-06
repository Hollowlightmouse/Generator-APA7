# Format Rules - APA 7th Edition

General rules of the 7th edition of the APA manual, for student papers. This file does not add institutional requirements: if the institution where the paper is submitted requests something specific that is not included here, the file will be updated with what the user confirms, and nothing else will be assumed.

## Page

- Paper size: **Letter** (Letter, 21,59 x 27,94 cm / 8,5 x 11 in).
- Margins: **2,54 cm (1 inch)** on all four sides.
- Page number in the **top right** corner, in the header.
- **The cover page has no visible number.** Numbering starts on page 2, which is
  what APA 7 requires. The cover page counts as page 1 for the rest of the
  calculations, but its number is not printed.

## Typography and Spacing

- Font: Times New Roman 12 pt throughout the document, including the cover page,
  the table of contents, the indexes, the table and figure legends, and the
  references.
- **Double** line spacing throughout the document.
- The body text is aligned **left**, not justified, with a first-line indent of
  1,27 cm (0,5 in) in each paragraph.
- The table of contents entries are Times New Roman 12 pt and **not bold**, even
  if the section title does have it.

## Heading Levels

| Level | Format |
|---|---|
| 1 | Centered, bold, with normal capital and lowercase letters |
| 2 | Left aligned, bold |
| 3 | Left aligned, bold italic |
| 4 | With first-line indent, bold, ending in a period. The text continues on the same line |
| 5 | With first-line indent, bold italic, ending in a period. The text continues on the same line |

The level 1 and 2 headings are **black, bold and Times New Roman 12 pt**. This is applied at run level (`bold: true, color: '000000', font: 'Times New Roman', size: 24`) because the `docx` library heading styles bring their own color and bold that override the document's.

Headings must always be applied as **real styles** (Heading 1/2/3, etc.), never only as bold text: the table of contents depends on that. See `word-toc-fields.md`.

### How it maps from the `.md`

This document uses `##` for the first two levels, like the rest of the skill documentation:

| In the `.md` | APA Level |
|---|---|
| `#` | (cover page, not a body heading) |
| `## Numero. Texto` | 1 (centered, bold, new page) |
| `## Texto` | 2 (left, bold) |
| `###` | 3 (left, bold + italic) |
| `####` | 4 (indent, bold, final period, text on the same line) |
| `#####` | 5 (indent, bold + italic, final period, text on the same line) |

Levels 3, 4 and 5 **are not flattened**: each hash keeps its level. Markdown cannot express "title in line with the body continuing", so a `####` or `#####` comes out as its own paragraph. For the authentic APA case (indented title in line), the parser provides `--detectar-niveles-en-linea`, which recognizes the convention:

- `**Titulo.** El texto sigue en la misma linea.` → level 4
- `***Titulo.*** El texto sigue en la misma linea.` → level 5

It is **opt-in on purpose**: bold at the start of a paragraph is also perfectly normal text, and enabling it without control would fill the table of contents with false headings. The parsing is done on the raw text, before `strip_leading_markers()`, which eats the leading asterisks.

## Very Wide Tables or Figures: Landscape Page

APA 7 allows placing that specific page in landscape within a document that is
otherwise portrait. When doing so:

- The 2,54 cm margins are kept.
- The page number is still displayed, relocated for the new orientation.
- A **section** break is used, not just a page break, before and after, in order to
  change the orientation without affecting the rest of the document.
- The table or figure must fit within the landscape margins.
- The wide table, its title and its note go **on the same landscape page**, and
  the text that continues afterwards returns to portrait in another section.
- `build-docx.js` decides this on its own, for readability and not by row count: a
  200-row, 3-column table can stay portrait without any problem. It switches to
  landscape when the width is not enough:
  - **many columns**: 6 or more (`UMBRAL_COLUMNAS_ANCHAS`), or
  - **cells so long they do not fit**, as long as rotating the page genuinely
    relieves them (the longest cell goes from 8 or more lines in portrait to
    fewer than 8 in landscape, `UMBRAL_LINEAS_APRETADAS`).

  Both cases are measured against the usable width of each page (6,5" in portrait and 9"
  in landscape) and not against a row rule. A 3-column table with extremely long cells
  is **not** rotated: going from 24 to 34 characters per line does not make it
  readable, and it is better to split it or move the content to text.
- Technical note: when declaring the landscape section you have to give `docx` the
  measures **in portrait** (12240 x 15840) and mark `orientation: LANDSCAPE`. The
  library swaps them on its own; giving them already inverted produces a
  `w:pgSz` with `w` smaller than `h`, which Word and LibreOffice read as portrait and
  leaves the table cramped.

## Tables and Figures

- Independent numbering: "Tabla 1", "Tabla 2"... and "Figura 1", "Figura 2"...
- The number and the title go on separate lines, with the number in bold and the
  title in italic.
- **The title and the legend are not optional.** If the `.md` does not bring them,
  the parser leaves a blocking question for each one and `build-docx.js` aborts
  with code 4. There is no `(sin titulo)` that gets printed in their place.
- In italic, only the word **"Nota."**, followed by normal text. By default
  **"Nota. Elaboración propia."**, unless the `.md` indicates another source; the
  parser fills it in on its own (`nota_origen: "default"`) and with
  `--notas-tabla-json` / `--notas-figura-json` a specific one is substituted.
- **The position distinguishes the case:**
  - **FIGURE note: below the image.**
  - **TABLE note: below the table.**
- The figure note has `keepNext` so that it does not end up alone at the bottom of a
  page while its image jumps to the next one. Without that the document looks
  incoherent even though the order in the flow is correct.
- Table borders: **only three horizontal lines** - the top border of the
  table, the line separating the header row from the body and the bottom border.
  **No** vertical lines, no side box and **no** lines between the data rows:
  those separations are made by the line spacing, not by a border. The header row
  may have light shading.

## Lists

When a line starts with two or more markers ("•  •", "– •", "• ••") or with
duplicated numbering ("1. 1.", "4. 1."), it is normalized to a single marker or a
single number. They are conversion artifacts and do not comply with APA.

## Run-Together Words (Extraction Artifacts)

Converting a PDF to text often loses internal spaces. The parser fixes them
automatically, but it is worth reviewing them while reading the `.md`.

**Cases that are separated:**

1. Lowercase followed by uppercase (`[a-z][A-Z]`): `Diferenciacionentre` ->
   `Diferenciacion entre`, `yFlaws` -> `y Flaws`.
2. Colon attached to an uppercase letter (`:[A-Z]`): `2:Diferenciacion` ->
   `2: Diferenciacion`.
3. Number attached to a word: `ACTIVIDAD1` -> `ACTIVIDAD 1`.

**Protected automatically (never separated, never reported):**

- Pure uppercase acronyms (`OWASP`, `CVSS`, `BPMN`): the rule requires lowercase
  followed by uppercase, so they are already safe.
- URLs (`https://...`, `www...`), DOIs (`10.1038/...`), e-mails and file names
  with an extension (`informe_v2.docx`). They are set aside before any rule runs
  and put back byte for byte, so a query string is never broken.

**Ambiguous (separated, but flagged for review):**

camelCase names (`NodeJS`, `OpenID`, `MySQL`) and terms mixing letters and
digits (`IPv6`, `SHA256`) cannot be told apart from a real gluing by pattern
alone. The parser still separates them, but the report lists them under
**Review**, not among the clear corrections, because they may well be the real
spelling of a product or a variable.

**How to keep a term intact:** write the terms, one per line (exact,
case-sensitive), in a text file and pass it with `--terminos-protegidos
<ARCHIVO>` (through `apa7.py parse`). Protected terms are masked just like URLs
and never appear in the change report. This is the alternative to editing the
source `.md`; use whichever the user prefers.

The un-gluing is never silent: **everything that changes is recorded**, split
into clear corrections and items to review, before building the `.docx`. A term
split that should not have been (`NodeJS` → `Node JS`) is either added to the
protected-terms file and the parser re-run, or corrected in the source `.md`
and the parser re-run. Prefixed surnames (`McDonald`, `MacArthur`) are handled
the same way.

## References

- Section heading: "Referencias" (level 1, centered, bold).
- Alphabetical order by the surname of the first author.
- **Hanging indent**: first line without indent, subsequent lines with a 1,27 cm
  (0,5 in) indent.
- Double line spacing, no extra space between references.
- Book and report titles in italic; article titles without italics
  (the journal does go in italics).
- If a reference from the `.md` does not follow the format (missing italics, missing the year
  in parentheses, the order is incorrect), it is corrected before including it.
