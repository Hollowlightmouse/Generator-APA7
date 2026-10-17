# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

No unreleased changes.

## [1.1.0] - 2026-10-17

### Added

- Selector de motores `--motor auto|word|libreoffice` con sondeo real de Word (`scripts/lib/motores.py`).
- Exportación directa con Word vía COM (`scripts/lib/word.py`) incluyendo `--actualizar-campos`, sin reutilizar sesiones abiertas, con timeout y limpieza de perfil solo cuando aplica.
- Caché de sondeo de Word: `.work/motor-word.json` para acelerar decisiones.
- Prueba de integración del verificador de figuras (`scripts/tests/test_verificador.py`, clase `TestNotaDeFigura`).
- Pruebas E2E de selección y exportación con dobles: `TestSeleccionDeMotor`, `TestExportConWord`, `TestBuildUnSoloMotor`.
- Pruebas de rutas, PIDs y registro COM (`scripts/tests/test_rutas.py`).
- `scripts/tests/test_build_docx.py` con pruebas reales de Node para orden imagen→nota y limpieza de etiqueta `Nota.`.
- `LICENSE` AGPL-3.0 oficial.
- Citas textuales APA: un bloque Markdown (`>`) se convierte en un párrafo de cita con sangría izquierda de 1,27 cm, sin sangría de primera línea y sin comillas (`md-a-manifiesto.py`, `build-docx.js`); el `&nbsp;` de los ejemplos de la guía se normaliza a un espacio.
- Atribución de tablas y figuras: cuando el `.md` no trae `Nota. …`, el parser ya no asume en silencio `Elaboración propia`; deja una pregunta **bloqueante** (`atribucion_fuente`) que ofrece proporcionar la fuente (`De "Título" por Autor, año. (URL)` o `Adaptado de …`) o confirmar `Elaboración propia`. Resolución vía `--notas-tabla-json` / `--notas-figura-json`.
- Tipografía seleccionable: `--fuente "<familia>"` (Times New Roman 12 por defecto; Arial 11, Calibri 11, Lucida Sans Unicode 10, Georgia 11 o Computer Modern 10) con `--tamano-fuente "<pt>"` opcional. La elección viaja en `MANIFEST.json` como `tipografia: {nombre, pt}` y la aplican `build-docx.js` (familia y tamaño en todo el documento) y `verificar-pdf.py` (chequeo de tipografía y de doble espacio **font-aware**: mide contra los métricos de la familia declarada vía `scripts/lib/fuentes.py`).
- Títulos APA: un título nunca conserva su número y nunca queda en MAYÚSCULAS SOSTENIDAS. El parser separa el número (clave `etiqueta`, no se imprime) y, si todo el título está en mayúsculas, lo reescribe a caja de oración conservando siglas (`OWASP`, `ISO/IEC`) y los términos de `--terminos-protegidos`; las palabras funcionales del español (`DE`, `DEL`, `LA`) siempre se convierten a minúscula. Un título ya en caja de oración no se toca.
- Contenido de tabla APA: el contenido de las celdas va en espacio **sencillo** y con la fuente **2 puntos menor** que el cuerpo (fila de encabezado incluida) vía `TAM_TABLA` en `build-docx.js`, con piso de 8 pt; el número, el título y la nota conservan el tamaño del cuerpo. Las líneas siguen siendo solo horizontales.
- Estilos de título APA **reales**: cada título se aplica como `Heading 1`..`Heading 5` (no solo negrita), de modo que la sección aparece en el panel de navegación de Word y un campo `TOC \o` la recolecta. `build-docx.js` redefine los cinco estilos de la librería (que vienen en azul y sobredimensionados) a la tipografía y el tamaño del documento, **negro** (`000000`), negrita, itálica en los niveles 3 y 5, doble espacio y `outlineLevel` correcto; la alineación por nivel, la sangría de primera línea de los niveles 4-5 y el salto de página siguen por párrafo.

### Changed

- `scripts/apa7.py`: validación estricta del nombre de motor, resolución única antes del build, cmd_export evita consultas a LibreOffice cuando usa Word, limpieza selectiva de perfil LibreOffice.
- `scripts/build-docx.js`: la imagen de figura usa `keepNext` cuando existe nota; la nota se escribe sin duplicar etiqueta `Nota.`.
- `scripts/verificar-pdf.py`: ignora fuentes usadas solo en espacios (no visibles) para aceptar Aptos en TOC de Word; criterio geométrico de nota bajo imagen, con aceptación documentada si pasa a página siguiente.
- `scripts/lib/rutas.py`, `scripts/lib/motores.py`, `scripts/lib/word.py`: soporte completo de motores.
- Portada: el número de página ahora es visible **desde la portada** (página 1), como pide la guía UNIMINUTO; `verificar-pdf.py` exige el número en la portada y en todas las páginas.
- Carpeta de trabajo única: cada documento tiene ahora **una sola** carpeta `deliverable/` junto al `.md`, con los entregables (`.docx` y `.pdf`, con el nombre del `.md`) en la raíz y **todos** los intermedios (`fuente.json`, `portada.json`, `MANIFEST.json`, `verificacion.json`, `paginas.json` y los logs) bajo `process/`. Sustituye a la antigua `<nombre>_apa/` con `datos/` y `logs/`. Como el nombre es un literal fijo, deja de necesitar saneamiento: se elimina `sanea_nombre` (y sus constantes `_INVALIDOS`, `_RESERVADOS`, `MAX_NOMBRE`). Varios `.md` en una misma carpeta comparten `deliverable/`; los entregables se distinguen por el nombre del `.md`. `.gitignore` actualizado (`deliverable/` en lugar de `*_apa/`).

### Fixed

- Evita falso crítico de tipografía cuando Word embebe fuentes solo en runs en blanco (TOC).
- Limpieza selectiva del perfil de LibreOffice: solo al usar ese motor.
- Resolución única del motor para evitar doble pasada con distinto motor.
- Numeración de páginas de los títulos **en línea** (niveles 4-5): `paginas-de-pdf.py` ahora ignora el punto final al emparejar el título del esquema del PDF, porque `build-docx.js` cierra esos títulos con un punto (`"...Texto."`) y el manifiesto los guarda sin él; sin esto, esas secciones nunca recibían número de página en el índice (`clave_titulo`).

### Changed

- Nota de figura: ahora va **debajo** de la imagen (con `keepNext` en la imagen); nota de tabla sigue debajo de la tabla. Actualizada documentación (SKILL.md, README.md, apa7-format.md).
- `materia_nrc` (curso/NRC) ahora es **campo bloqueante** en portada: la materia es obligatoria, solo el código NRC es opcional dentro del mismo campo. Ajustado parser, verificador y pruebas.
- Parser: consume la línea `Nota.` del `.md` sin duplicar (`nota_origen: "md"`); leyenda de caption (`SEQ`) completamente en negrita; entradas de índice (TOC) sin negrita.
- Pruebas de plataforma: `skipUnless` por SO; suite verde en Windows, macOS y Linux (simulado).
- Documentación de motores actualizada: `--motor auto|word|libreoffice`; auto sondea Word en Windows/macOS; LibreOffice es fallback; `exit 2` si motor pedido inutilizable; caché `.work/motor-word.json` y `APA7_WORD`.
- README: licencia AGPL-3.0, estructura del repo, ejemplo rápido, sección de contribución, CI con `unittest`.

### Test coverage

- 426 tests passing in `scripts/tests` (one opt-in integration test skipped by default).
  Verificación real de corpus Word/LibreOffice OK.


## [1.0.0] - 2026-10-03

First release of the `generate-apa-document` skill. There is no earlier tag, so
this section records the state that is being published, including the hardening
that was done before the first release.

### Added

- Initial release of the `generate-apa-document` skill.
- Interactive-free pipeline: `check` -> `install` -> `parse` -> `build` ->
  `export` -> `verify`, driven by `scripts/apa7.py`.
- Markdown to `MANIFEST.json` parser with automatic correction of conversion
  artifacts (glued words, duplicated list markers, glued numbering) and a
  report split into corrections and items to review.
- Word `.docx` builder with institutional 3-zone cover, real `TOC` fields for
  the table of contents and the lists of tables and figures, wide tables on an
  occasional landscape page, and APA 7 references with hanging indent.
- Headless LibreOffice export to PDF (Microsoft Word is not required).
- PDF verification with `pymupdf`: cover, indices, captions, landscape pages,
  notes, complete figure images and references.
- `scripts/paginas-de-pdf.py`: reads a throwaway PDF and reports the real page
  of every heading, table and figure, used by the TOC double pass.
- `scripts/tests/fixtures/sintetico.md`: synthetic source document for the
  end-to-end parse test (never a real document).

### Changed

- `install --dry-run` creates nothing at all, not even the working directory or
  the npm anchor, and distinguishes "simulated" from "not confirmed" in its
  messages.
- Every command suggested by `apa7.py` prints the real, copyable command with
  the interpreter that is running (`sys.executable`), instead of assuming a
  `python` on `PATH`.
- The `soffice` candidate list uses the launcher names Snap and Flatpak
  document (`/snap/bin/libreoffice`, `org.libreoffice.LibreOffice`), marked as
  best effort and unverified.
- `scripts/lib/fuentes.py`: `ArialMT` is no longer accepted as a Times New
  Roman substitute, and the comments about the accepted metric clones were
  corrected.

### Fixed

- Arial PDFs are now rejected by the typography check instead of passing it.
- The test suite is independent of the account it runs under: the elevation
  tests state the privilege they assume, so they pass both as root and as an
  ordinary user.
- The TOC fields ship the measured page numbers as their cached result, so the
  exported PDF shows the indices correctly without Microsoft Word.
