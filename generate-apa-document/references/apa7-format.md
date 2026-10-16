# Format Rules - APA 7th Edition

General rules of the 7th edition of the APA manual, for student papers. This file does not add institutional requirements: if the institution where the paper is submitted requests something specific that is not included here, the file will be updated with what the user confirms, and nothing else will be assumed.

## Page

- Paper size: **Letter** (Letter, 21,59 x 27,94 cm / 8,5 x 11 in).
- Margins: **2,54 cm (1 inch)** on all four sides.
- Page number in the **top right** corner, in the header.
- **The page number is visible from the cover page.** The cover is page 1 and shows
  its number, as the UNIMINUTO guide requires ("Numeración de páginas: en la
  esquina superior derecha, a partir de la portada"). This overrides the APA
  student-paper default, where the cover page carries no visible number.

## Typography and Spacing

- **One typeface only**, from the list APA 7 allows for student papers, applied
  to the whole document (cover page, table of contents, indexes, table and figure
  legends and references). Times New Roman 12 pt is the default; the other
  accepted sizes follow the APA manual. The choice goes on the command line
  (`--fuente`, with `--tamano-fuente` to override the size) and is recorded in the
  manifest as `tipografia`, which both `build-docx.js` and `verificar-pdf.py` read.

  | `--fuente` | Default size |
  |---|---|
  | Times New Roman | 12 pt |
  | Arial | 11 pt |
  | Calibri | 11 pt |
  | Lucida Sans Unicode | 10 pt |
  | Georgia | 11 pt |
  | Computer Modern | 10 pt |

- **Double** line spacing throughout the document. The verifier measures the line
  advance against the declared family, so a document in Arial is not judged by the
  Times New Roman advance.
- The body text is aligned **left**, not justified, with a first-line indent of
  1,27 cm (0,5 in) in each paragraph.
- The table of contents entries keep the document's typeface at its size and are
  **not bold**, even if the section title does have it.

## Heading Levels

| Level | Format |
|---|---|
| 1 | Centered, bold, with normal capital and lowercase letters |
| 2 | Left aligned, bold |
| 3 | Left aligned, bold italic |
| 4 | With first-line indent, bold, ending in a period. The text continues on the same line |
| 5 | With first-line indent, bold italic, ending in a period. The text continues on the same line |

Every heading is applied as a **real Word style** (`Heading 1`..`Heading 5`), never only as bold text: that is what makes the section appear in Word's
navigation pane, lets a `TOC \o` field pick it up and lets anyone restyle the
document. See `word-toc-fields.md`.

The `docx` library's built-in Heading styles are **blue** (`2E74B5` / `1F4D78`)
and oversized, so `build-docx.js` **redefines all five** to APA type: the
document's typeface and size (Times New Roman 12 pt by default), **black**
(`000000`), bold, italic for levels 3 and 5, double spacing, no extra space,
and the matching `outlineLevel` (level 1 → 0). The per-level alignment (level 1
centered, the rest left), the first-line indent of the inline levels 4-5 and
the page break stay on each paragraph, so a level 1 only breaks the page when
the page is already in use.

A heading **never carries a number** and is **never written in ALL CAPS** (UNIMINUTO guide). When a `##` heading in the `.md` is written as `## 3. ANÁLISIS DE OWASP`, the parser drops the number (kept apart as `etiqueta`, which is not printed) and rewrites the all-caps text to APA sentence case: `Análisis de OWASP`. Acronyms survive; Spanish words do not. Concretely, inside a fully-uppercase heading the parser:

- capitalizes the first word and lowercases the rest;
- keeps a token uppercase when it is a probable acronym — a short all-caps token (2 to 6 letters) that the document never writes in lowercase (`OWASP`, `SAMM`, `SDL`), or any token with a digit, `&` or `/` (`ISO/IEC`, `27002`);
- always lowercases Spanish closed-class words (`DE`, `DEL`, `LA`, `Y`, ...);
- keeps the tokens listed in `--terminos-protegidos`.

A heading already in sentence case is left untouched, so a proper noun the author wrote by hand (`Microsoft`, `Método`) is never damaged. Limitation to keep in mind: a proper noun written **only** in caps inside a fully-uppercase heading is lowercased unless it is the first word (`## ANÁLISIS DE MICROSOFT SDL` → `Análisis de microsoft SDL`); add it to `--terminos-protegidos` to pin its exact form.

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
- In italic, only the word **"Nota."**, followed by normal text. When the `.md`
  does not indicate a source, the parser fills in **"Nota. Elaboración propia."**
  (`nota_origen: "default"`) **and leaves a blocking question `atribucion_fuente`**:
  the authorship default is not assumed silently. Answer it by giving the source
  or by confirming `Elaboración propia`, via `--notas-tabla-json` /
  `--notas-figura-json` (a specific one is substituted). The APA 7 forms for a
  reproduced or adapted table/figure are:
  - `De "Título del documento" por Autor, año de publicación. (URL)` (reproduced)
  - `Adaptado de "Título del documento" por Autor, año de publicación. (URL)`
    (modified)
  - `Elaboración propia` (own work; the default)
  A `Nota. …` line already in the `.md` wins over the default and is kept
  verbatim (`nota_origen: "md"`), and then the question is not raised.
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
- **Table content spacing and size:** the guide allows the content to go in
  **1.5 or single spacing** (this document uses single: 240 twips) and lets the
  font **drop 2 points** below the body. `build-docx.js` applies that drop to the
  whole content, header row included (`TAM_TABLA = TAM - 4` half-points, floored
  at 8 pt), while the number, the title and the note keep the body size.

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

## Textual Quotes

- **Fewer than 40 words**: they go **inline**, inside the paragraph and between
  quotation marks. Neither the font nor the size is changed.
- **More than 40 words**: they go in their **own paragraph**, with a **left indent
  of 1,27 cm (0,5 in)**, **without quotation marks**, keeping the **double spacing**
  and the font size. They are never boxed nor set in a smaller font.
- In the `.md` a long quote is written as a Markdown block quote: one or more
  lines each starting with `>` (a bare `>` acts as a blank separator). Consecutive
  `>` lines are joined into a **single** paragraph. The parser turns them into a
  `cita` block and `build-docx.js` applies the 1,27 cm left indent, with no
  first-line indent and no quotation marks.
- The `&nbsp;` used in the guide's examples to push the quote to the right is
  normalized to an ordinary space.

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
