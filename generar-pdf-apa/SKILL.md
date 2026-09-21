---
name: generar-pdf-apa
description: Genera un documento académico en Word (.docx) y PDF con Normas APA 7ª edición, a partir de un archivo Markdown (.md) de origen, sus imágenes y los JSON de layout (content_list, content_list_v2, middle, model) que traen tamaño y posición de las figuras. Usar esta skill SIEMPRE que el usuario pida explícitamente "PDF APA", "genera el PDF APA", "documento APA" o equivalente para un trabajo académico de Uniminuto, con portada institucional, tabla de contenido, índice de tablas e índice de figuras funcionales y referencias corregidas a formato APA. Exporta el PDF con LibreOffice (método F2) y NO requiere Microsoft Word. No usar esta skill si solo se pide leer o resumir un .md, ni si el usuario no ha usado la frase de activación "PDF APA".
---

# Generar documento académico en Normas APA 7 (Word + PDF)

Convierte un `.md` fuente (más sus imágenes y JSON de layout) en un `.docx` con portada institucional, tabla de contenido / índice de tablas / índice de figuras creados con la librería `docx` (entradas con número de página mediante campos `PAGEREF`) y referencias en formato APA 7 — y exporta ese mismo documento a PDF con **LibreOffice headless** (método F2, único motor de conversión). Se entregan ambos archivos.

**Motor obligatorio: LibreOffice (F2). Microsoft Word NO es requisito.**

## Cuándo se activa

Solo cuando el usuario lo pide explícitamente con una frase tipo "PDF APA" / "genera el PDF APA" para un documento concreto. No se activa por el simple hecho de que exista un `.md` en la conversación.

## Insumos requeridos

El usuario siempre entrega tres cosas juntas — si falta alguna, pídeselas antes de continuar (no asumas ni inventes el contenido):

1. El **`.md`** fuente (texto, encabezados, tablas en markdown, referencias, marcas de imagen `![]()`).
2. Las **imágenes** referenciadas en el `.md` (carpeta `images/` o archivos sueltos).
3. Los **JSON de layout** (`*_content_list.json`, `*_content_list_v2.json`, `*_middle.json`, `*_model.json` o solo alguno de ellos) — se usan únicamente para saber tamaño y posición real de cada imagen; no repitas ese trabajo a mano, usa `scripts/calcular_tamano_imagenes.py` (ver más abajo).

## PASO 0 — Verificación e instalación del entorno (OBLIGATORIO, antes de transformar)

Antes de tocar cualquier documento, ejecuta el preflight. Si algo falta, se instala automáticamente; si algo no se puede instalar, **detente y avisa al usuario antes de continuar** (nunca transformes un documento con herramientas faltantes).

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\User\.agents\skills\generar-pdf-apa\scripts\comprobar-entorno.ps1"
```

- Si sale `ENTORNO OK` → sigue con el PASO 1.
- Si sale la lista `FALTAN:` → ejecuta:
  ```powershell
  powershell -ExecutionPolicy Bypass -File "C:\Users\User\.agents\skills\generar-pdf-apa\scripts\instalar-entorno.ps1"
  ```
  Este script instala (vía winget/npm/pip) lo que falte y re-verifica. Si falla algo, detente y avisa.

Requisitos verificados/instalados: PowerShell 5.1+, Node.js LTS + npm, librería `docx` (npm) en el directorio del generador, venv Python con `pymupdf`, y **LibreOffice**. Detalle en `references/requisitos-sistema.md`.

## Flujo de trabajo

1. **Leer el `.md`** y extraer: título, autores, datos de curso, secciones (con su nivel de encabezado real), tablas, figuras, y la lista de referencias. Al extraer, hacer un **inventario explícito de tablas y figuras** (cuántas hay, sus títulos y fuentes). **Al extraer encabezados y párrafos, revisar y corregir palabras pegadas** (artefactos de la conversión: `EJERCICIO 2:Diferenciaciónentre Bugs yFlaws`, `ACTIVIDAD1:MapeodeControlesISO27002`) — ver la sección "Palabras pegadas" en `references/normas-apa7.md`.
2. **Calcular el tamaño de cada imagen** con `scripts/calcular_tamano_imagenes.py` a partir del JSON `*_content_list.json` (trae `img_path` y `bbox` juntos; úsalo como fuente principal) o `*_content_list_v2.json`. El `*_model.json` no trae la ruta del archivo de imagen, así que solo sirve como respaldo de proporción, no para identificar qué imagen es cada una. No estimes tamaños de imagen a ojo.
3. **Completar la portada** siguiendo exactamente `references/portada-institucional.md` (organización canónica = página 1 del PDF `ARQUITECTURA_X_SEGMED_APA7_corregido.pdf`). La portada usa un **layout posicional en 3 zonas**: **título arriba** (y≈100pt), **autores al centro vertical** (y≈293pt) y **resto del bloque institucional anclado abajo** (facultad, vicerrectoría, materia-NRC, docente, fecha, desde y≈514pt con pitch ≈27.6pt entre líneas). **Logo opcional solo si el usuario lo aporta** (preguntar siempre; si no se aporta, sin logo). Por cada campo que el `.md` no traiga, pregúntale al usuario si lo agrega o si se omite para este documento — nunca lo completes por tu cuenta. Esto incluye siempre preguntar por el título/profesión del docente, salvo que el usuario diga explícitamente que se omita esa vez. **Regla general: ante cualquier valor faltante que no encuentres en los archivos adjuntados (`.md`, imágenes y JSON de layout), pregunta** — no inferir ni inventar.
4. **Autoría de tablas y figuras**: si el `.md` no indica una fuente distinta, usa "Elaboración propia" por defecto (no preguntes esto, ya está resuelto).
5. **Construir el `.docx`** con la librería `docx` (npm), siguiendo:
   - Las reglas de formato de `references/normas-apa7.md` (tamaño carta, márgenes, fuente, interlineado, formato de niveles de encabezado, sangría francesa en referencias).
   - Las instrucciones de `references/campos-word-toc.md` para que la tabla de contenido, el índice de tablas y el índice de figuras queden **funcionales con número de página real** (entradas con estilos TOC 1/2 y campos `PAGEREF`; `features: { updateFields: true }` para que LibreOffice recalcule al abrir). **La TOC se genera siempre; el índice de tablas y el índice de figuras SOLO si el documento tiene tablas/figuras.** Si el análisis no encuentra tablas o figuras, no se genera el índice vacío (y si quedó vacío, se elimina): pregunta primero para confirmar que el documento realmente no las tiene.
   - Tablas con ancho explícito (`width` + `columnWidths` + `layout: FIXED`): LibreOffice no renderiza tablas sin anchos de columna definidos (bug real encontrado); las leyendas usan número literal ("Tabla N", no campo `SEQ`, porque la numeración es fija).
   - Corrige el formato de las referencias a Normas APA (orden alfabético, sangría francesa, cursivas donde corresponde) aunque el `.md` ya traiga una lista de referencias.
6. **Exportar a PDF con LibreOffice** (método F2) usando:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "C:\Users\User\.agents\skills\generar-pdf-apa\scripts\export-fase2-pdf.ps1" -SrcDocx "ruta\al\documento.docx" [-OutDir "carpeta\entrega"]
   ```
   Trabaja sobre copia temporal con **perfil aislado** de LibreOffice por corrida. El aviso `Could not find platform independent libraries <prefix>` por stderr es benigno.
7. **Verificar antes de entregar** con `pymupdf` del venv de Docling: la portada tenga todos los campos (título, autores en un solo párrafo, facultad/vicerrectoría/materia-NRC/docente/fecha), los índices presentes (TOC siempre; índice de tablas/figuras solo si el documento las tiene) muestren entradas con número de página correcto, las leyendas de tablas/figuras muestren número y título, la hoja horizontal de la matriz se vea completa, y las referencias tengan sangría francesa.
8. **Entregar ambos archivos** (`.docx` y `.pdf`) con `present_files`.

## Reglas ya definidas (no volver a preguntar)

- Se activa solo con la frase explícita "PDF APA".
- Insumo siempre triple: `.md` + imágenes + JSON.
- **Motor de exportación: LibreOffice (método F2), únicamente. Word NO es requisito ni se usa.**
- **PASO 0 obligatorio**: verificar/instalar el entorno antes de transformar; si algo no se puede instalar, detener y avisar.
- Campo de portada faltante → preguntar (agregar u omitir), incluido el título del docente.
- **Portada en 3 zonas (layout posicional del PDF canónico)**: título arriba, autores al centro, bloque institucional anclado abajo. **Logo opcional**: solo si el usuario lo aporta (preguntar siempre; sin logo si no se aporta).
- **Cualquier valor faltante durante el análisis/estructuración → preguntar (no inferir ni inventar):** tablas/figuras sin fuente (→ "Elaboración propia" solo si no hay fuente, regla ya resuelta y no se pregunta), NRC, título de sección, datos de autores, fechas, logo de portada, etc. Solo se consideran datos existentes los que aparecen en los archivos adjuntados (`.md`, imágenes y JSON de layout).
- **Inventario de tablas/figuras**: la TOC se genera siempre; el índice de tablas y el índice de figuras **solo si el documento las tiene**. Si el análisis no encuentra tablas o figuras, **preguntar siempre** (confirmar que el documento realmente no las tiene) y NO generar el índice vacío — si quedó generado vacío, eliminarlo. Lo mismo aplica a figuras: sin figuras, no hay índice de figuras.
- Formato: tamaño carta y márgenes de Normas APA 7ª edición exclusivamente (sin reglas institucionales adicionales).
- Referencias: siempre se corrigen a formato APA.
- Palabras pegadas / artefactos de extracción: siempre corregirlas (espacios faltantes entre palabras, "2:Palabra", "ACTIVIDAD1:"), respetando nombres propios y términos técnicos (AppSec, DevSecOps, FinTechX, MediCloud, etc.).
- Lista de autores en portada: separados por coma, con "y" antes del último, sin importar la cantidad.

## Archivos de referencia

- `references/normas-apa7.md` — reglas de formato APA 7 (márgenes, fuente, interlineado, niveles de encabezado, referencias).
- `references/portada-institucional.md` — organización canónica de la portada (página 1 del PDF `ARQUITECTURA_X_SEGMED_APA7_corregido.pdf`) y qué hacer cuando falta información (preguntar siempre).
- `references/campos-word-toc.md` — cómo generar TOC / índice de tablas / índice de figuras funcionales con la librería `docx` y el flujo de exportación a PDF por LibreOffice (método F2), con las trampas ya resueltas (TOC literal con `PAGEREF`, bug de `.flat()`, tablas sin `columnWidths` que LibreOffice no renderiza, `updateFields`).
- `references/requisitos-sistema.md` — requisitos obligatorios del entorno (F2) y forma de verificar/instalar cada herramienta.
- `scripts/calcular_tamano_imagenes.py` — calcula el tamaño (en pulgadas) que debe tener cada imagen al insertarla, preservando su proporción real según el JSON de layout.
- `scripts/ejemplo-build.js` — ejemplo funcional de generación del `.docx` (portada, TOC, índice de tablas, tablas con hoja horizontal, referencias).
- `scripts/comprobar-entorno.ps1` y `scripts/instalar-entorno.ps1` — preflight del PASO 0 (verifica e instala Node+npm+docx, venv+pymupdf y LibreOffice).
- `scripts/export-fase2-pdf.ps1` — exportación del `.docx` a PDF con LibreOffice headless (perfil aislado), sobre copia temporal.