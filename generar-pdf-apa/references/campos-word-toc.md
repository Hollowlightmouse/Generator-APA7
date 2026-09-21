# Tabla de contenido, índice de tablas e índice de figuras (docx funcional, LibreOffice)

Lecciones aprendidas en el informe "Informe_Comparativo_SDLC" para el **pipeline LibreOffice (F2)**. La exportación final es con **LibreOffice headless**; Word ya no participa. Esta guía existe para no repetir los bugs ya resueltos.

## Regla de oro del pipeline LibreOffice

LibreOffice calcula los números de página al abrir el documento, pero **no construye un campo `TOC` vacío desde cero** (ignora `updateFields` para eso) y **no resuelve el resultado de campos `SEQ`**. Por lo tanto:

- La **tabla de contenido** se genera como **entradas literales** (texto del título + punto suspensivo + número vía campo `PAGEREF`).
- Las **leyendas de tablas/figuras** llevan **número literal** ("Tabla 1", no `SEQ`).
- Los **números de página** de TOC e índices son campos `PAGEREF` reales que LibreOffice **sí resuelve** (bookmark → página). Por eso los índices quedan funcionales: si el contenido cambia de página, el número se actualiza al reabrir.

## Recetas que SÍ funcionan (docx npm 9.7.1 + LibreOffice)

### 1. Índice de encabezados: bookmarks en cada heading
- Tras parsear el markdown, numera los headings (`b.hid = ++hid`, orden de aparición) y guarda `headings[] = { hid, level, text }`.
- Envuelve el run del heading en un bookmark `_H<hid>`:
  ```js
  case 'h1': return para({ heading: d.HeadingLevel.HEADING_1, /* ... */ }, [new d.Bookmark({ id: '_H' + b.hid, children: [new d.TextRun({ text: b.text, bold: true, /* TNR 12 negro */ })] })]);
  ```
- `_H`-prefijo porque Word/LibreOffice no admiten dígitos iniciales en nombres de bookmark.

### 2. Tabla de contenido literal con PAGEREF
- La entrada usa **estilos reales TOC 1 / TOC 2** declarados en `styles.paragraphStyles` del `Document` (así LibreOffice y Word leen formato APA), más `tabStops` con puntos y el campo:
  ```js
  function tocEntry(entry) {
    return para({ style: entry.level === 1 ? 'TOC1' : 'TOC2', spacing: { before: 0, after: 0, line: 480 },
                  tabStops: [{ type: d.TabStopType.RIGHT, position: 9350, leader: d.LeaderType.DOT }],
                  alignment: d.AlignmentType.LEFT },
      [new d.TextRun({ text: entry.text }), new d.TextRun({ children: [new d.Tab()] }),
       new d.SimpleField('PAGEREF _H' + entry.hid + ' \\h')]);
  }
  ```
- Estilos de párrafo declarados en el documento (formato APA: TNR 12, sin negrita, negro, doble interlineado `spacing.line 480` — nota: `spacing` va dentro de **`paragraph`**, NO a nivel raíz del estilo):
  ```js
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: 'Times New Roman', size: 24 } } },
    paragraphStyles: [
      { id: 'TOC1', name: 'TOC 1', basedOn: 'Normal', run: { font: 'Times New Roman', size: 24, bold: false, color: '000000' }, paragraph: { spacing: { line: 480 } } },
    ],
  },
  ```
- Importante: `features: { updateFields: true }` en las opciones del `Document` hace que LibreOffice evalúe **campos existentes** al abrir (PAGEREF y similares). LibreOffice NO dedica un pase propio para ello; el DOCX generado lleva `w:updateFields`.

### 3. Índice de tablas/figuras: bookmark + entrada literal con PAGEREF
- Bookmark `_Tabla<n>` alrededor del número de la leyenda y **número literal** (el pipeline LO no resuelve `SEQ`):
  ```js
  new d.Bookmark({ id: '_Tabla' + n, children: [new d.TextRun({ bold: true, children: ['Tabla ' + n] })] })
  ```
- Entrada en el índice con `PAGEREF _Tabla<n> \h`, número y título literales, tabulación derecha con puntos:
  ```js
  para({ spacing: { before: 0, after: 0, line: 480 }, alignment: d.AlignmentType.LEFT,
         tabStops: [{ type: d.TabStopType.RIGHT, position: 9350, leader: d.LeaderType.DOT }] },
    [new d.TextRun({ children: ['Tabla ' + n, '  ', title] }),
     new d.TextRun({ children: [new d.Tab()] }),
     new d.SimpleField('PAGEREF _Tabla' + n + ' \\h')])
  ```
- La leyenda ("Tabla 1" + título) va en un **mismo párrafo** con un salto de línea (`new d.TextRun({ break: 1 })`) entre el número y el título en cursiva — número y título en líneas separadas pero un solo párrafo (APA).

### 4. Tablas con ancho explícito (bug real de LibreOffice)
LibreOffice **no renderiza el contenido de tablas sin anchos de columna** (matriz en hoja horizontal quedó vacía / cuerpo de tablas siguiente ausente). Solución probada:
```js
new d.Table({
  rows: rowsOut,
  width: { size: usable, type: d.WidthType.DXA },
  columnWidths,
  layout: d.TableLayoutType.FIXED,
  borders: { /* solo líneas horizontales */ },
})
```
- `usable` = ancho útil de la sección: 12960 twips (hoja horizontal, márgenes 1440) o 9360 twips (carta vertical).
- Primera columna más ancha (cabecera de criterios): 2160 twips (horizontal) / 1560 (vertical); el resto se reparte parejo.
- Estándar de bordes: solo líneas horizontales (superior, inferior, `insideHorizontal`).

### 5. Trampas de la librería `docx` (bug real, aplican igual en LO)
- Si una función helper devuelve `[para]` (array con un `Paragraph`) y lo insertas como `[ helper(), tabla, nota ]`, `Array.prototype.flat()` **solo aplana 1 nivel** y deja el array anidado → se serializa como `<0/>` y **el documento no abre**. Usa spread: `[...helper(), tabla, nota]` (o `.flat(2)`).
- `TableOfContents` del la librería no sirve en el pipeline LO: LibreOffice ignora el campo TOC vacío → usar TOC literal (receta 2). El índice de tablas manual ya lleva `line: 480` desde el generador.
- `spacing` de un estilo de párrafo va en `paragraph.spacing` (si lo pones a nivel raíz, `docx` falla).

## Exportación a PDF con LibreOffice (método F2 — motor obligatorio)

Script: `scripts/export-fase2-pdf.ps1`.

- Trabaja sobre **copia temporal** (`lo_copy_<PID>`) con **perfil de usuario aislado** (`lo_profile_<PID>`) por corrida:
  ```powershell
  & $soffice -env:UserInstallation=file:///C:/.../lo_profile_<PID> --headless --convert-to pdf:writer_pdf_Export --outdir <tmp> <docx>
  ```
- En Windows, `soffice.exe` con `&` **se desprende y no espera** → usa `Start-Process -PassThru -NoNewWindow -Wait -RedirectStandardError <log>` (el stderr va a `export-fase2.log`; NO lo redirijas con `2>&1 | Out-File` en PS 5.1 con `$ErrorActionPreference='Stop'`, lanza `NativeCommandError` y aborta).
- El aviso `Could not find platform independent libraries <prefix>` (stderr) es **benigno**.
- Rendimiento medido (Informe_Comparativo_SDLC): conversión ~11 s (F2 total ~10.8 s vs F1 Word ~11.9 s). El PDF de LibreOffice pesa ~+30% bytes frente al de Word por igual calidad (ídem ~354 KB).
- Paginación: LibreOffice recalcula con los anchos FIXED de tabla; en el piloto quedó la misma estructura (portada/TOC/índice/5 secciones/hoja horizontal/refs) y el TOC/índice reflejan los números recién calculados. Puede haber una página de diferencia de cierre vs el artefacto resuelto por Word (contenido idéntico).

## Checklist de verificación (usa pymupdf sobre el PDF resultante)

- [ ] Portada con todos los campos.
- [ ] TOC: títulos TNR 12, no negrita, negro, doble interlineado y **números de página** que coinciden con las páginas reales de cada heading.
- [ ] Índice de tablas: "Tabla N  Título…  página" con números correctos (Tabla 1 en hoja horizontal).
- [ ] Cuerpo: tablas con contenido (especialmente la hoja horizontal, que en LO sin anchos queda vacía), notas "Elaboración propia", referencias con sangría francesa.
- [ ] Sin páginas en blanco entre TOC e índice.

## Anexo histórico — el pipeline anterior con Word COM (ya NO se usa)

Antes del estándar F2 la conversión era Word COM (`ExportAsFixedFormat`, desconectado), con post-proceso COM para que el campo TOC naciera con doble interlineado TNR 12 no negrita tras `$toc.Update()`, y `SEQ Tabla` en el cuerpo. Word COM era frágil (procesos colgados, `~$*`, copia temporal obligatoria) y quedó descartado. Si algún día se mantiene Word COM, conservar las reglas: nunca exportar al mismo nombre base en el mismo directorio; trabajar sobre copia en `export-work`; actualizar campos y exportar el PDF **en sesiones de Word separadas**.