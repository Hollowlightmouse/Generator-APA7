# Institutional cover page

Layout of the first page of a document. The cover page is a **template with
placeholders**, not fixed text: it is filled in with whatever the document
brings and with whatever the user answers. It carries no institution, subject,
or person name.

## Positional organization in 3 zones

Letter page (612 x 792 pt) with 72 pt margins (1 inch). The content is NOT
stacked consecutively: it is spread across three vertical zones. Positions are
measured from the top edge of the sheet and `verificar-pdf.py` checks them as a
**warning**, never as a failure.

The usable height (792 - 72 - 72 = 648 pt) is split into three equal bands of
216 pt:

```
   y= 36   Numero de pagina (header, esquina superior derecha)
   y= 72  +----------------------------------------------------+  <- margen
           |  [ZONA ALTA]   72 .. 288 pt                       |
   y=108   |     Logo opcional (160x160 px) encima del titulo   |
           |     Titulo del trabajo (negrita, centrado)        |
   y=288  +----------------------------------------------------+
           |  [ZONA CENTRO]  288 .. 504 pt                     |
   y=383   |     Integrantes: UN parrafo centrado, vertical      |
           |     centrado en la banda (y ~383 con 1 linea)      |
   y=504  +----------------------------------------------------+
           |  [ZONA BAJA]    504 .. 720 pt                     |
   y=555   |     Docente (+ titulo/profesion)                   |
           |     Materia y NRC                                 |
   y=583   |     Facultad / universidad                        |
   y=610   |     Vicerrectoria                                 |
           |     Sede                                         |
   y=693   |     Fecha          <- ultima linea, siempre a ~706 |
   y=720  +----------------------------------------------------+
```

**3-zone rule (not negotiable):**

1. **Top** — optional logo and the document title. The logo ALWAYS goes above
   the title, never beside it.
2. **Vertical center** — the members, the only element in the middle of the
   page.
3. **Bottom** — the rest of the institutional data, with a pitch of about
   27.6 pt between lines and **anchored to the bottom**: the last line always
   ends near y=706 instead of starting at a fixed level. That way the block
   does not shift depending on how many fields the manifest brings.

The order of the bottom block is the usual one and depends on the available
information: instructor, course/NRC, faculty, vice-rector's office, location,
date. A field that does not exist is simply not printed: it leaves no blank
gaps.

### How it is implemented

In `build-docx.js`, `construirPortada()` returns **a borderless 1-column x
3-row table**, with `HeightRule.ATLEAST` and `verticalAlign` set to `TOP` /
`CENTER` / `BOTTOM` respectively. The table is only a positioning mechanism: it
has no borders, it is not printed, and the verifier ignores it.

Two details that are not negotiable:

- The padding of the top zone (720 twips) is **subtracted** from the height of
  its row. The three rows must add up to exactly 12960 twips; if they add up to
  more, the last row overflows onto a new page and the document gains a page.
- The logo paragraph uses **single** line spacing. With double spacing, the
  image box inflates and pushes the rest of the cover page.

Previously this was done by stacking paragraphs with empty filler lines, which
left the members glued to the title and the institutional block in the middle
of the page. Do not go back to that approach: the zones do not hold if the
content changes.


## Fields

| Field | Placeholder | Rule |
|---|---|---|
| Logo | `logo` | Optional. It is only inserted if the user provides the file. Otherwise nothing is inserted: no placeholder, no frame. Maximum width ~2 in, proportional height. **Always ask.** |
| Page number | — | Top right, like the rest of the document. The cover page **does not show** a visible number. |
| Title | `titulo` | Bold, centered. It may span several lines. |
| Authors | `autores` | **A single centered paragraph**, not bold. Full names separated by commas and with "y" before the last one, no matter how many there are. Not one per line. It accepts `"Nombre Apellido y Nombre Apellido"` or `["Nombre Apellido", "Nombre Apellido"]`: the text is respected as is, without splitting it or assuming authors. |
| Faculty and university | `facultad` | One centered line, exactly as it appears in the source document. |
| Location or unit | `vicerrectoria` | Centered, optional. Always asked about when missing, but its absence never blocks delivery: the cover is generated without that line and `verificar-pdf.py` warns. |
| Course and code | `materia_nrc` | Centered, exactly as it appears in the document. The field accepts any course code (for example an NRC), not just that format. Always asked about when missing, but its absence now blocks delivery: the cover is generated without that line and delivery is blocked until the field is provided; `verificar-pdf.py` does not reach this field because it blocks earlier. |
| Instructor | `docente` | Centered. Their degree or title is always asked about (`docente_titulo`), unless the user says to skip it that time. |
| Date | `fecha` | Centered, in the document's format. |

## What to do when a value is missing

General skill rule, not just for the cover page:

- **For any value that does not appear in the input files, ask.**
  Never infer, assume, or invent. This applies to all cover page fields and to
  any other piece of data in the document (a table or figure source, a section
  title, authors, date, etc.).
- If the user asks to omit a field, do **not** leave a blank line or a
  placeholder: simply do not include that line. For the logo, nothing is
  inserted.
- The instructor's degree or title is a separate case: it is **always asked
  about**, even if the rest of the cover page is complete.

Most cover fields **block delivery** until they are answered: title, members,
faculty / university, instructor and date; their absence always opens a blocking
question in `pendientes_bloqueantes`. Three fields are the opposite — **always
asked, never blocking**: the logo, the instructor's title / profession and the
vice-rector's office (`vicerrectoria`). If one of them stays missing, the cover is
generated without that line and the document is delivered; `verificar-pdf.py` flags
each one as a warning so the agent still mentions it. The course/NRC line
(`materia_nrc`) is now **blocking**: it must be answered before delivery; if the
user omits it, delivery is blocked until the field is provided.

## How it is filled in practice

The parser detects by pattern the fields that an academic paper's `.md` usually
brings (faculty, location, course with code, instructor, authors, date) and
puts them in `portada` inside `MANIFEST.json`. Whatever it does not detect
appears in `campos_faltantes` and must be asked about.

The user's answers are passed in a `portada.json` file delivered with
`--portada`, and take priority over anything inferred. The generator never
invents an institution name: if `facultad` does not come, it prints no line at
all.