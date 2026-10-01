---
name: generar-pdf-apa
description: Genera o estructura un documento académico en Word (.docx) y PDF con Normas APA 7ª edición, a partir de un archivo Markdown (.md) de origen, sus imágenes y los JSON de layout (content_list, content_list_v2, middle, model) que traen tamaño y posición de las figuras. USAR SIEMPRE que el usuario pida montar un documento o PDF académico en APA, formalizar un trabajo, armarlo con portada institucional, tabla de contenido, índice de tablas, índice de figuras y referencias APA, o aplicar formato APA a un texto. La activación es por INTENCIÓN, no por palabras clave: vale igual "genera el PDF APA", "arma mi trabajo en formato APA", "estructura este informe con norma APA" o "pasame esto a APA con portada y referencias". También usar cuando falten datos y haya que preguntarlos antes de continuar. Exporta el PDF con LibreOffice headless y NO requiere Microsoft Word. NO usar si solo se pide leer, resumir, traducir o corregir ortografía de un texto.
---

# Generar documento académico en Normas APA 7 (Word + PDF)

Convierte un `.md` fuente (más sus imágenes y JSON de layout) en un `.docx` con portada institucional, tabla de contenido / índice de tablas / índice de figuras creados con la librería `docx` (entradas con número de página mediante campos `PAGEREF`) y referencias en formato APA 7 — y exporta ese mismo documento a PDF con **LibreOffice headless**, único motor de conversión. Se entregan ambos archivos.

**Motor obligatorio: LibreOffice. Microsoft Word NO es requisito.**

Todas las rutas de este documento son **relativas a la raíz de la skill**
(la carpeta que contiene este `SKILL.md`). En PowerShell, sitúate en ella:

```powershell
cd "<ruta de la skill>"
```

## Cuándo se activa

Por **intención**, no por frase exacta. Se activa cuando el usuario quiere un
documento académico *montado* en APA, sea cual sea la fórmula:

| Dice algo como | Se activa |
|---|---|
| "genera el PDF APA", "documento APA" | sí |
| "arma mi trabajo en APA con portada y referencias" | sí |
| "estructura este informe con norma APA" | sí |
| "pasame este .md a APA, que quede presentable" | sí |
| "formatea esto según APA 7" | sí |
| "solo léeme el .md" / "resúmeme" / "corrígeme la ortografía" | **no** |

No se activa por el simple hecho de que exista un `.md` en la conversación, ni
cuando el texto es solo una porción suelta que hay que leer, resumir, traducir
o revisar. La frontera es el **documento completo montado**, no el APA como
detalle de estilo.

Cuando el usuario pida structuring pero falte información obligatoria, **se
pregunta antes de generar** (ver "Preguntas obligatorias" más abajo).

## Insumos requeridos

1. El **`.md`** fuente (texto, encabezados, tablas en markdown, referencias, marcas de imagen `![]()`).
2. Las **imágenes** referenciadas en el `.md` (carpeta `images/` o archivos sueltos).
3. Los **JSON de layout** (`*_content_list.json`, `*_content_list_v2.json`, `*_middle.json`, `*_model.json` o solo alguno de ellos) — se usan únicamente para saber tamaño y posición real de cada imagen y tabla.

Si falta alguno, pídelo antes de continuar. **El `.md` es obligatorio; los JSON y las imágenes son opcionales.** El pipeline acepta:

- `.md` solo → documento de texto. Sin figuras ni índices de tablas/figuras.
- `.md` + `images/` → figuras a tamaño fijo razonable.
- `.md` + imágenes + JSON → figuras a su **tamaño real** (de `bbox`) y tablas a su ancho real de página.

## PASO 0 — Verificación e instalación del entorno (OBLIGATORIO)

Antes de tocar cualquier documento:

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\comprobar-entorno.ps1"
```

- Si sale `ENTORNO OK` → PASO 1.
- Si sale `FALTAN:` → `powershell -ExecutionPolicy Bypass -File "scripts\instalar-entorno.ps1"`, que instala (winget/brew, npm, pip) lo que falte y vuelve a verificar.

Si tras instalar algo sigue faltando, **detente y avisa al usuario**: nunca transformes un documento con el entorno incompleto. Detalle en `references/requisitos-sistema.md`.

## PASO 0.5 — Preguntas obligatorias (no se salta)

El parser (`scripts/md-a-manifiesto.py`) es el detector de faltantes. Al
ejecutarlo, escribe en `MANIFEST.json → diagnostico`:

- `preguntas`: lista de objetos con `id`, `bloqueante`, `campo`, `tipo`,
  `objetivo`, `indice`, `pregunta`, `opciones` y `como_resolver`.
- `pendientes_bloqueantes`: solo los `id` que impiden entregar.
- `preguntas_texto`: las mismas preguntas como texto plano, para leerlas rápido.

**Regla dura: si `pendientes_bloqueantes` no está vacío, se pregunta al usuario
y NO se construye el `.docx`.** No se parchea el `.docx` a mano ni se rellena el
hueco con un texto inventado.

Qué garantiza cada parte, porque no es lo mismo:

- **Lo comprueba la máquina**: `build-docx.js` aborta con código 4 si falta el
  título de alguna tabla o figura. Es un subconjunto de los bloqueantes: los
  títulos y leyendas, que son los únicos verificables sin preguntar.
- **Lo comprueba el agente**: el resto (`portada_*`, `sin_tablas`, `sin_figuras`).
  Nada en el código impide construir el `.docx` si solo quedan esos, así que
  quien lee `pendientes_bloqueantes` y decide parar es el agente. Si el usuario
  confirma que no hay tablas o figuras, o que se omite un campo de portada,
  basta con dejarlo dicho: no hace falta un flag extra para desbloquear.
  **Excepción ya resuelta:** `portada_logo` y `portada_docente_titulo` se
  preguntan pero **no bloquean** (`bloqueante: false`). El documento se entrega
  sin ellos y `verificar-pdf.py` lo señala como aviso.

### Títulos y leyendas que faltan

El parser pregunta una **elección binaria** por cada tabla o figura sin título:

> La Tabla 3 no tiene título. ¿Redacto el título a partir del contexto del
> informe, o me lo proporcionas tú?

Según la respuesta:

1. **El usuario lo proporciona** → se anota en el `.md` (`Tabla 3. <título>`),
   o se pasa un JSON con `--titulos-tabla-json` / `--titulos-figura-json`
   (formato `{"3": "Título..."}`, índice base 1), y se vuelve a ejecutar el parser.
2. **El usuario pide redactarlo** → se escribe el título en ese mismo JSON,
   leyendo el contexto del documento para que describa de verdad lo que muestra
   la tabla o la figura, y se vuelve a ejecutar el parser.

Las dos rutas terminan igual: el parser deja de reportar el faltante. Lo que no
se hace nunca es decidirlo por el usuario.

### Lo que NO se pregunta

La **fuente** de tablas y figuras tiene default resuelto: `Nota. Elaboración
propia`, salvo que el `.md` o un JSON indiquen otra. No se pregunta por eso.

## Flujo de trabajo

1. **Leer el `.md`** y extraer: título, autores, datos de curso, secciones con su nivel de encabezado real, tablas, figuras y referencias. Hacer un **inventario explícito de tablas y figuras** (cuántas, con qué títulos y fuentes). Al leer, **corregir palabras pegadas** (artefactos de conversión del PDF: `2:Diferenciaciónentre Bugs`, `ACTIVIDAD1:MapeodeControles`). La regla completa y el fichero de excepciones están en `references/normas-apa7.md`.

2. **Calcular el tamaño de cada imagen** con `scripts/calcular_tamano_imagenes.py` a partir del JSON de layout (`*_content_list.json` trae `img_path` y `bbox` juntos: es la fuente principal; `*_content_list_v2.json` como alternativa). El `*_model.json` no trae la ruta del archivo de imagen, así que solo sirve como respaldo de proporción. No estimar tamaños a ojo.

3. **Completar la portada** con `references/portada-institucional.md`: layout posicional en 3 zonas (título arriba, autores al centro vertical, bloque institucional anclado abajo), logo opcional **solo si el usuario lo aporta** (preguntar siempre). Por cada campo que el `.md` no traiga, **preguntar al usuario** si se agrega o se omite. Esto incluye **siempre** preguntar por el título/profesión del docente, salvo que el usuario diga que esa vez se omite. **Regla general: ante cualquier valor que no esté en los archivos adjuntados, preguntar — no inferir ni inventar.** Las respuestas se guardan en un `portada.json` que se pasa al **parser** con `--portada`: las fusiona con lo deducido y las escribe en `MANIFEST.json`. `build-docx.js` no recibe ese fichero.

4. **Autoría de tablas y figuras**: si el `.md` no indica fuente distinta, el parser pone `Nota. Elaboración propia` automáticamente (marcado con `nota_origen: "default"`). Esto ya está resuelto, no se pregunta. Si hay que cambiar o precisar la fuente de una en concreto, se pasa `--notas-tabla-json` / `--notas-figura-json` con `{"indice": "texto"}`.

5. **Construir el `.docx`** con `scripts/build-docx.js` (librería `docx`, npm), leyendo `MANIFEST.json`. Reglas de formato en `references/normas-apa7.md`; instrucciones de índices funcionales en `references/campos-word-toc.md`. Puntos críticos ya resueltos:
   - La TOC se genera **siempre**; el índice de tablas y el de figuras **solo si el documento las tiene**. Si no las tiene, **preguntar** para confirmar y **no** generar el índice vacío.
   - Tablas con `width` + `columnWidths` + `layout: FIXED`: LibreOffice no renderiza tablas sin anchos de columna definidos.
   - Leyendas con **número literal** ("Tabla 1"), no campo `SEQ` (LibreOffice no resuelve `SEQ`).
   - **Abre con código 4 si falta un título o leyenda.** No hay `(sin título)`: es un defecto, no un texto.
   - **Nota de figura ANTES de la imagen** y con `keepNext`, para que la nota no quede huérfana en la página anterior. **Nota de tabla DESPUÉS de la tabla**.
   - **Tablas con bordes SOLO horizontales**: borde superior, línea bajo la fila de encabezado y borde inferior. Sin verticales ni líneas entre filas de datos.
   - **Orientación automática por tabla**: el cuerpo se parte en tramos y cada tramo es una sección del `.docx`. Una tabla se pasa sola a una **página horizontal** (con su título y su nota) cuando no cabe en vertical: 6 o más columnas, o celdas tan largas que girar la página las alivia de verdad. El criterio es legibilidad/anchura, **no el número de filas**. Después, el texto vuelve a vertical. Ver `references/normas-apa7.md`.
   - Los encabezados de nivel 1 se rompen con `pageBreakBefore`, **no** con un párrafo suelto que lleve un `PageBreak`: ese párrafo vacío genera una página en blanco cuando la anterior ya está llena.

6. **Exportar a PDF con LibreOffice**:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "scripts\export-pdf.ps1" -Docx "<ruta>\documento.docx" [-OutDir "<carpeta>\entrega"]
   ```

   Trabaja sobre copia temporal con **perfil aislado** de LibreOffice por corrida, que se borra al terminar. El aviso `Could not find platform independent libraries <prefix>` por stderr es benigno. **No** cambiar a `soffice.exe` ni a `Start-Process`: cuelgan el proceso (ver `references/campos-word-toc.md`).

7. **Verificar antes de entregar** con `scripts/verificar-pdf.py` (usa el Python que devuelve `Get-PythonPath`, con `pymupdf`): portada en 3 zonas e integrantes en un solo párrafo (**esto se comprueba, como aviso**), índices con número de página correcto, leyendas con número y título, tablas con contenido, **cada tabla con su nota debajo en el PDF**, **cada nota de figura encima de su imagen**, referencias con sangría francesa. Los avisos de portada (logo o título del docente sin preguntar) **no son fallas**: se comentan al usuario y se entrega. Si falla una comprobación crítica, corregir y volver a exportar.

8. **Entregar ambos archivos** (`.docx` y `.pdf`).

## Reglas ya definidas (no volver a preguntar)

- Se activa por **intención**: pedir un documento académico montado en APA, no hace falta decir la palabra "PDF APA" (ver "Cuándo se activa").
- Motor de exportación: **LibreOffice headless, únicamente**. Word no es requisito ni se usa.
- **PASO 0 obligatorio** antes de transformar.
- Campo de portada faltante → preguntar (agregar u omitir), incluido el título del docente.
- **Portada en 3 zonas**: título arriba (logo opcional **encima** del título), integrantes al centro vertical, bloque institucional anclado abajo. Se construye con una tabla de 3 filas sin bordes, no con párrafos de relleno. **Logo**: solo si el usuario lo aporta, pero **preguntar siempre**: si no se Preguntó, el verificador lo avisa.
- **Cualquier valor faltante → preguntar (no inferir ni inventar).** Solo son "datos existentes" los que aparecen en los archivos adjuntados. Excepción ya resuelta: fuente de tablas/figuras → "Nota. Elaboración propia" por defecto.
- **Título o leyenda de tabla/figura → preguntar** (¿se redacta del contexto o lo da el usuario?) y **bloquear la entrega** hasta responder. No se emite un documento con títulos en blanco.
- **Nota de figura encima de la imagen; nota de tabla debajo de la tabla** (APA 7 las distingue por posición, no solo por texto).
- **Inventario de tablas/figuras**: la TOC siempre; índices de tablas y de figuras solo si existen. Si el análisis no las encuentra, **preguntar** antes de asumir.
- Formato: tamaño carta y márgenes de APA 7ª edición, sin reglas institucionales adicionales.
- Referencias: siempre se corrigen a formato APA.
- Palabras pegadas: siempre se corrigen (espacios faltantes, "2:Palabra", "ACTIVIDAD1:"), respetando nombres propios y términos técnicos protegidos por `references/terminos-whitelist.txt`.
- Lista de autores en portada: separados por coma, con "y" antes del último, sin importar la cantidad. `autores` acepta `"Nombre Apellido y Nombre Apellido"` o `["Nombre Apellido", "Nombre Apellido"]`; se respeta el texto tal cual, sin partirlo.
- **Ampliar la lista blanca es parte del trabajo.** La lista trae términos genéricos, no los nombres propios del documento. Los tramos que el deglue no puede separar con confianza (nombres de marca o de proyecto pegados al resto de la frase) salen en la cola de revisión: se revisan a mano en el `.md` y, los que sean un término real, se **añaden a `references/terminos-whitelist.txt` y se vuelve a ejecutar el parser**. Sin ese paso el deglue parte marcas por términos que nadie protege.

## Archivos de la skill

| Archivo | Función |
|---|---|
| `SKILL.md` | Este documento. |
| `references/normas-apa7.md` | Reglas de formato APA 7 (página, tipografía, niveles de encabezado, tablas/figuras, palabras pegadas, referencias). |
| `references/portada-institucional.md` | Layout de la portada en 3 zonas, campos, y qué hacer cuando falta información. |
| `references/campos-word-toc.md` | Cómo hacer funcionales la TOC y los índices (`PAGEREF`, `updateFields`), trampas de la librería `docx`, y cómo invocar LibreOffice sin colgarlo. |
| `references/requisitos-sistema.md` | Requisitos, preflight, instalación y resolución de rutas. |
| `references/terminos-whitelist.txt` | Términos intocables para el deglue. **Única fuente**: no hay lista incrustada en el código. |
| `scripts/md-a-manifiesto.py` | `.md` + JSON de layout → `MANIFEST.json`. Enriquece, deduplica, degluea, calcula tamaños, aplica la nota por defecto y deja las preguntas bloqueantes en `diagnostico`. Flags: `--portada`, `--titulos-tabla-json`, `--titulos-figura-json`, `--notas-tabla-json`, `--notas-figura-json`, `--detectar-niveles-en-linea`, `--imagenes`, `--whitelist`, `--sin-deglue`, `--ancho-pagina-landscape`. |
| `scripts/build-docx.js` | `MANIFEST.json` → `.docx` (portada, índices, tablas, figuras, referencias). Parte el cuerpo en secciones y pasa a página horizontal las tablas que no caben en vertical. **Aborta con código 4 si falta un título o leyenda.** |
| `scripts/export-pdf.ps1` | `.docx` → `.pdf` con LibreOffice headless. |
| `scripts/verificar-pdf.py` | Verifica el `.pdf` (pymupdf): portada, índices, leyendas, sangrías, notas de tabla debajo y notas de figura encima de la imagen. |
| `scripts/lib/rutas.ps1` | Resolución portable de rutas, Python, Node y LibreOffice. |
| `scripts/comprobar-entorno.ps1` | Preflight del PASO 0. |
| `scripts/instalar-entorno.ps1` | Instala lo que falte y re-verifica. |
| `scripts/calcular_tamano_imagenes.py` | Calcula el tamaño de cada imagen desde el `bbox` del JSON de layout. |
| `.work/` | Estado generado: `node_modules` de `docx`. No es código fuente; se regenera con `instalar-entorno.ps1`. |
