# Skill `generar-pdf-apa`

Genera un documento académico en **Word (.docx) y PDF** con las **Normas APA 7ª edición**, listo para entregar: portada institucional, tabla de contenido e índices de tablas/figuras con números de página reales, todo a partir de un Markdown `.md` de origen (extraído con MinerU/Docling, por ejemplo).

> **Motor de exportación: LibreOffice (headless). No se requiere Microsoft Word.**

---

## Qué es y para qué sirve

Es una **skill de agente** (un `SKILL.md` + scripts auxiliares) que convierte la salida —normalmente imperfecta— de una conversión de PDF a Markdown en un trabajo académico **final, verificado y presentable**.

No es un conversor de PDF a Word: parte de un `.md` que ya existe. Lo que hace es **estructurarlo y formalizarlo**:

- **Portada institucional** con layout en 3 zonas (título arriba, autores en el centro vertical, bloque institucional anclado abajo) y logo opcional solo si lo aportas.
- **Tabla de contenido siempre** e **índice de tablas y figuras solo cuando existen**, con páginas reales vía campos `PAGEREF` que se actualizan al abrir el documento.
- **Corrección automática** de artefactos de la conversión: palabras pegadas (`2:Diferenciaciónentre Bugs` → `2: Diferenciación entre Bugs`), marcadores de lista duplicados (`• •`, `1. 1.`), numeración pegada (`ACTIVIDAD1:`), siempre respetando nombres propios y términos técnicos.
- **Referencias corregidas a APA 7**: orden alfabético, sangría francesa y cursivas donde corresponde, aunque el `.md` ya las traiga como lista.
- **Tablas amplias en hoja horizontal** puntual (permitido por APA 7) para matrices comparativas que no caben en vertical.
- **Verificación final con `pymupdf`** antes de entregar: portada, índices, leyendas, hojas horizontales, notas y referencias.
- Entrega **ambos archivos** (`.docx` y `.pdf`).

## Requisitos (herramientas que necesita)

| Herramienta | Para qué se usa | Cómo se consigue |
|---|---|---|
| **PowerShell 5.1+** | ejecutar los scripts de la skill | incluido en Windows |
| **Node.js LTS + npm** | generar el `.docx` con la librería `docx` | `winget install OpenJS.NodeJS.LTS` o `brew install node` |
| **Librería `docx` (npm)** | construir portada, TOC, índices y tablas | `npm install docx@9.7.1 --no-save` |
| **Python 3 + `pymupdf`** | verificar el PDF final | `pip install pymupdf` |
| **LibreOffice** | **único motor** de conversión `.docx → .pdf` | `winget install TheDocumentFoundation.LibreOffice` o `brew install --cask libreoffice` |
| **Conexión a internet** | descargas de winget/brew, npm y pip | — |

La skill valida estas herramientas automáticamente antes de tocar el documento (**PASO 0 / preflight**). Si falta alguna, la instala sola; si algo no puede instalarse, **se detiene y avisa** sin transformar el documento.

Rutas de herramientas sobrescribibles por variable de entorno: `APA7_SOFFICE`, `APA7_PYTHON`, `APA7_NODEDIR`, `APA7_WORKDIR`, `APA7_SKILL_ROOT`. Detalle en [`generar-pdf-apa/references/requisitos-sistema.md`](generar-pdf-apa/references/requisitos-sistema.md).

## Instalación

1. **Copia la carpeta** `generar-pdf-apa/` al directorio de skills de tu agente de IA, de modo que `SKILL.md` quede en la raíz del directorio de skills. Por ejemplo: `<tu-usuario>\.agents\skills\generar-pdf-apa`.
2. **El entorno se arma solo.** La primera vez se ejecuta el preflight:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "...\generar-pdf-apa\scripts\comprobar-entorno.ps1"
   ```

   Si sale `FALTAN:`, el instalador automático lo resuelve y vuelve a verificar:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "...\generar-pdf-apa\scripts\instalar-entorno.ps1"
   ```

   > No instales nada a mano ni transformes documentos con herramientas faltantes.

La dependencia `docx` se instala en `generar-pdf-apa/.work/node_modules`, dentro de la propia skill: así `require('docx')` resuelve siempre, sin importar desde dónde se ejecute. Esa carpeta es estado generado y está en `.gitignore`.

## Cómo se usa

La skill se **activa por intención**: basta con que pidas un documento académico montado en APA (*"genera el PDF APA"*, *"arma mi trabajo en formato APA"*, *"estructura este informe con norma APA"*, *"pasame esto a APA con portada y referencias"*). No se activa por el simple hecho de que exista un `.md`, ni cuando solo pides leer, resumir, traducir o corregir ortografía.

### 1. Entrega los insumos

| # | Insumo | Para qué sirve | Obligatorio |
|---|--------|----------------|-------------|
| 1 | El **`.md`** fuente | texto, encabezados, tablas en markdown, referencias y marcas de imagen `![]()` | sí |
| 2 | Las **imágenes** referenciadas | se insertan en el documento (carpeta `images/` o archivos sueltos) | no |
| 3 | Los **JSON de layout** (`*_content_list.json`, `*_content_list_v2.json`, `*_middle.json`, `*_model.json`) | tamaño y posición real de cada figura y ancho de tabla | no |

El pipeline acepta tres grados de entrada:

- `.md` solo → documento de texto, sin figuras ni índices de tablas/figuras.
- `.md` + `images/` → figuras a tamaño fijo razonable.
- `.md` + imágenes + JSON → figuras a su **tamaño real** (desde el `bbox`) y tablas a su ancho real de página.

> Regla de oro: si falta algo imprescindible, la skill **te lo pide** antes de continuar. Nunca asume ni inventa contenido.

### 2. Responde solo lo que falta

De tus archivos saldrá casi todo. La skill te preguntará únicamente por los **datos que no aparezcan** en tus documentos: datos de portada (docente y su título/profesión, materia y su código, fecha), logo (opcional, solo si lo aportas), títulos/leyendas de tablas o figuras que no los traigan y, ante la duda, si el documento realmente no tiene tablas o figuras.

**Regla dura:** mientras haya preguntas marcadas como bloqueantes, el documento **no se construye**. Nada se rellena con texto inventado ni se parchea el `.docx` a mano.

### 3. Recibe los dos archivos

`.docx` y `.pdf` verificados página por página, listos para entregar.

## Flujo de trabajo

1. **Activación** — el usuario pide un documento académico en APA (por intención).
2. **PASO 0 · Preflight** — `comprobar-entorno.ps1` verifica las herramientas. Si falta algo, `instalar-entorno.ps1` lo instala y re-verifica. Si no queda `ENTORNO OK`, se detiene y avisa.
3. **Insumos** — se recibe el `.md` (obligatorio), las imágenes y los JSON de layout (opcionales).
4. **Parser** — `md-a-manifiesto.py` convierte `.md` + JSON de layout en `MANIFEST.json`: enriquece, deduplica, separa palabras pegadas, calcula tamaños reales de imagen y aplica la nota por defecto. Deja las preguntas pendientes en `diagnostico`.
5. **Preguntas obligatorias (PASO 0.5)** — si `pendientes_bloqueantes` no está vacío, se pregunta al usuario (datos de portada, títulos/leyendas faltantes, confirmación de "sin tablas/figuras") y no se construye hasta resolverlas. Las respuestas de portada se guardan en `portada.json` y se pasan al parser con `--portada`.
6. **Construir el `.docx`** — `build-docx.js` lee `MANIFEST.json` y arma portada, TOC, índices, tablas, figuras y referencias. Aborta con código 4 si falta un título o leyenda.
7. **Exportar a PDF** — `export-pdf.ps1` convierte `.docx → .pdf` con LibreOffice headless (copia temporal + perfil aislado por corrida).
8. **Verificar** — `verificar-pdf.py` (con `pymupdf`) comprueba portada, índices con página correcta, leyendas, notas de tabla debajo, notas de figura encima de su imagen y sangría francesa en referencias. Si falla algo crítico, se corrige y se vuelve a exportar.
9. **Entregar** — se entregan ambos archivos (`.docx` y `.pdf`).

```
Activación
   └─ PASO 0  Preflight del entorno ──► (falta algo) instalar-entorno.ps1 ──► re-verificar
        └─ Insumos: .md (+ imágenes + JSON de layout)
             └─ md-a-manifiesto.py ──► MANIFEST.json (+ diagnostico)
                  └─ PASO 0.5  ¿pendientes bloqueantes? ──► (sí) preguntar al usuario
                       └─ build-docx.js ──► documento.docx
                            └─ export-pdf.ps1 (LibreOffice headless) ──► documento.pdf
                                 └─ verificar-pdf.py ──► (fallas) corregir ──► re-exportar
                                      └─ Entrega: .docx + .pdf
```

## Cómo funciona por dentro

- **Tamaño de imágenes sin estimar a ojo:** el parser lee el `bbox` de cada figura en el JSON de layout (normalizado a 0..1000), conserva su proporción real y acota el ancho al contenido útil.
- **Prosa y orden salen del `.md`**, no del JSON: los JSON enriquecen el ancho de las tablas y el tamaño de las figuras, pero no reescriben el texto.
- **TOC funcional de verdad:** las entradas usan campos `PAGEREF` sobre *bookmarks*, así el número de página se recalcula al abrir el documento. LibreOffice no resuelve campos `SEQ`, por eso la numeración de tablas y figuras es literal.
- **Tablas renderizables en LibreOffice:** ancho explícito (`width` + `columnWidths` + `layout: FIXED`) y bordes solo horizontales; sin esto, LibreOffice no las muestra (bug conocido).
- **Orientación automática por tabla:** una tabla que no cabe en vertical (6 o más columnas, o celdas muy largas) se pasa sola a una **página horizontal** con su título y su nota; después el texto vuelve a vertical. El criterio es legibilidad, no el número de filas.
- **Nota de figura encima de la imagen; nota de tabla debajo de la tabla** (APA 7 las distingue por posición, no solo por texto).
- **Exportación limpia:** `soffice.com` (no `soffice.exe`), perfil aislado de LibreOffice por corrida y copia temporal; el mensaje `Could not find platform independent libraries <prefix>` por stderr es benigno y se ignora.
- **Nada de listas blancas incrustadas:** los términos que el deglue no debe separar viven solo en `references/terminos-whitelist.txt`.

## Estructura del repositorio

```
Generator-APA7/
├── README.md
└── generar-pdf-apa/
    ├── SKILL.md                          # definición de la skill (activación, flujo, reglas)
    ├── .gitignore                        # excluye estado generado (.work, __pycache__)
    ├── references/
    │   ├── normas-apa7.md                # reglas de formato APA 7
    │   ├── portada-institucional.md      # layout de la portada en 3 zonas
    │   ├── campos-word-toc.md            # TOC/índices funcionales y trampas de docx/LibreOffice
    │   ├── requisitos-sistema.md         # herramientas, preflight y resolución de rutas
    │   └── terminos-whitelist.txt        # términos intocables para el deglue
    ├── scripts/
    │   ├── comprobar-entorno.ps1         # preflight del entorno (PASO 0)
    │   ├── instalar-entorno.ps1          # autoinstalación de lo faltante
    │   ├── md-a-manifiesto.py            # .md + JSON → MANIFEST.json
    │   ├── build-docx.js                 # MANIFEST.json → .docx
    │   ├── export-pdf.ps1                # .docx → .pdf con LibreOffice
    │   ├── verificar-pdf.py              # verificación del .pdf (pymupdf)
    │   ├── calcular_tamano_imagenes.py   # tamaño real de imágenes desde los JSON de layout
    │   └── lib/
    │       └── rutas.ps1                 # resolución portable de rutas y herramientas
    └── .work/                            # estado generado (node_modules de docx)
```

## Limitaciones

- **Windows / PowerShell 5.1+** (los scripts están pensados para este entorno).
- No convierte directamente de PDF a Word: parte de un `.md` ya extraído por MinerU o Docling. Los JSON de layout son opcionales, pero sin ellos las figuras no conservan su tamaño real.
- Formato **APA 7 carta con márgenes 1 in**, sin reglas institucionales adicionales.
- La portada no lleva número de página visible; la numeración arranca en la página 2.

---

*Skill `generar-pdf-apa` — documento académico APA 7 (.docx + PDF) sin depender de Microsoft Word.*
