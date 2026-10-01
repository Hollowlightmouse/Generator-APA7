# Tabla de contenido, indice de tablas e indice de figuras

Como construir los indices para que queden **funcionales**, con numeros de
pagina reales, y como exportar el resultado a PDF. Este archivo existe para no
volver a tropezar con los fallos ya resueltos.

> El nombre del archivo es historico: la conversion ya no la hace Word. El
> motor es LibreOffice y todo lo de aqui aplica al pipeline con `docx` (npm) y
> LibreOffice.

## La regla de oro

LibreOffice recalcula los numeros de pagina al abrir el documento, pero **no
construye un campo `TOC` vacio desde cero** (ignora la opcion de actualizar
campos para eso) y **no resuelve los campos `SEQ`**. Por tanto:

- La **tabla de contenido** se genera como **entradas literales**: el texto del
  titulo, un tabulador con puntos y el numero de pagina mediante un campo
  `PAGEREF`.
- Las **leyendas de tablas y figuras** llevan **numero literal** ("Tabla 1"), no
  un campo `SEQ`.
- Los **numeros de pagina** de la tabla de contenido y de los indices son campos
  `PAGEREF` reales, que LibreOffice si resuelve (bookmark -> pagina). Por eso los
  indices quedan vivos: si el contenido se mueve de pagina, el numero se
  actualiza al reabrir.

## Recetas que funcionan

### 1. Bookmarks en cada encabezado

Despues de parsear el markdown, se numera cada heading segun su orden de
aparicion y se envuelve su run en un bookmark `_H<id>`:

```js
case 'h1': return para({ heading: d.HeadingLevel.HEADING_1, /* ... */ },
  [new d.Bookmark({ id: '_H' + b.hid,
     children: [new d.TextRun({ text: b.text, bold: true })] })]);
```

El prefijo `_H` es necesario porque Word y LibreOffice no admiten nombres de
bookmark que empiecen por un digito.

### 2. Entrada de tabla de contenido con PAGEREF

Usa estilos de parrafo reales `TOC 1` / `TOC 2` declarados en el `Document`
(para que LibreOffice y Word lean el formato APA) y un tabulador derecho con
puntos:

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

Estilos en el documento, con `spacing` dentro de **`paragraph`** y no a nivel
raiz del estilo (a nivel raiz la libreria falla):

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

`features: { updateFields: true }` es lo que hace que LibreOffice evalue los
campos existentes al abrir. No es un pase propio de LibreOffice: lo que hace es
escribir `w:updateFields` en el `.docx`.

### 3. Indice de tablas y de figuras

Mismo esquema que la tabla de contenido: bookmark sobre el numero de la leyenda,
numero y titulo literales, y `PAGEREF` para la pagina.

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

La leyenda va en **un solo parrafo** con un salto de linea entre el numero y el
titulo: asi el numero queda en negrita en su linea y el titulo en cursiva en la
siguiente, que es lo que pide APA.

**El indice solo se genera si hay elementos que indexar.** Si el documento no
tiene tablas no se crea indice de tablas; si no tiene figuras, no se crea indice
de figuras. Un indice vacio se elimina.

### 4. Tablas con ancho explicito

**LibreOffice no renderiza el contenido de una tabla sin anchos de columna**: la
tabla sale vacia. Es un fallo silencioso, el `.docx` se abre bien y no aparece
nada dentro.

```js
new d.Table({
  rows: rowsOut,
  width: { size: usable, type: d.WidthType.DXA },
  columnWidths,
  layout: d.TableLayoutType.FIXED,
  borders: { /* solo lineas horizontales */ },
})
```

- `usable` es el ancho util de la seccion segun sus margenes y orientacion.
- La primera columna (la de criterios) suele ser mas ancha; el resto se reparte
  parejo.
- Bordes: solo superior, inferior e `insideHorizontal`.

El ancho total sale del `bbox` de la tabla en el JSON de layout de MinerU, no de
estimarlo a ojo.

### 5. Trampas de la libreria `docx`

- Si un helper devuelve `[para]` y se inserta como `[ helper(), tabla, nota ]`,
  `Array.prototype.flat()` **solo aplana un nivel**: el array anidado se
  serializa como `<0/>` y **el documento no abre**. Usar spread (`[...helper(),
  tabla, nota]`) o `.flat(2)`.
- El `TableOfContents` de la libreria no sirve aqui: LibreOffice ignora el campo
  `TOC` vacio. Se usa la entrada literal de la receta 2.
- `spacing` de un estilo de parrafo va en `paragraph.spacing`.

## Exportacion a PDF con LibreOffice

Script: `scripts/export-pdf.ps1`. Trabaja sobre una copia temporal y crea un
perfil de usuario aislado por corrida, para no chocar con una sesion de
LibreOffice abierta por el usuario.

```powershell
& $soffice.com --headless --norestore --nolockcheck --nofirststartwizard `
    -env:UserInstallation=file:///.../lo_profile_<id> `
    --convert-to pdf:writer_pdf_Export --outdir <tmp> <docx>
```

**Como se invoca, y por que no de otra forma.** Contradecir esta parte hace
colgar el script:

- **Usar `soffice.com`, no `soffice.exe`.** El `.exe` se desprende y el proceso
  hijo hereda el handle del pipe, asi que PowerShell se queda leyendo para
  siempre.
- **No usar `Start-Process`.** Si el proceso padre ya redirige la salida, la
  redireccion anidada deja a LibreOffice sin terminar nunca.
- **La captura es `& $bin @args 2> $fichero`**, con `$ErrorActionPreference`
  puesto a `Continue` mientras dura la llamada y restaurado despues.
- **Nunca `2>&1 | Out-File`** junto a `$ErrorActionPreference = 'Stop'`: PowerShell
  lanza `NativeCommandError` y aborta.
- Con redireccion a fichero de .NET el proceso tampoco termina (los streams
  quedan abiertos por el hijo). Por eso se escribe a fichero con `2>`.
- El aviso `Could not find platform independent libraries <prefix>` en stderr es
  **benigno** y no debe tratarse como error.

## Checklist de verificacion

Sobre el PDF resultante:

- [ ] Portada con los campos que existen; sin numero de pagina visible.
- [ ] Tabla de contenido: numeros de pagina que coinciden con la pagina real de
      cada titulo.
- [ ] Indices de tablas y de figuras: entradas "Tabla N  Titulo...  pagina" con
      el numero correcto, y solo si el documento tiene esos elementos.
- [ ] Tablas con contenido (una pagina horizontal vacia es el sintoma del
      problema de `columnWidths`).
- [ ] Leyendas con numero y titulo, en parrafo unico.
- [ ] Referencias con sangria francesa.
- [ ] Sin paginas en blanco entre la tabla de contenido y el cuerpo.
