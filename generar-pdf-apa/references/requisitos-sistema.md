# Requisitos del sistema (entorno obligatorio — método F2 / LibreOffice)

Requisitos verificados e instalados por el **PASO 0** de la skill. Microsoft **Word NO es un requisito**.

## Lista literal (F2)

| Herramienta | Uso | Verificación | Instalación |
|---|---|---|---|
| PowerShell 5.1+ | ejecutar scripts de la skill | `$PSVersionTable.PSVersion` | incluido en Windows |
| Node.js LTS + npm | generar el `.docx` con la librería `docx` | `node --version` / `npm --version` | `winget install OpenJS.NodeJS.LTS` |
| librería `docx` (npm) | construir el `.docx` (TOC, índices, tablas) | `npm ls docx` en el directorio del generador | `npm install docx@9.7.1 --no-save` en el directorio del generador |
| venv de Docling + `pymupdf` | verificación del PDF (páginas, fuentes, índices) | `python -c "import fitz"` | `pip install pymupdf` en el venv |
| LibreOffice | **único motor** de conversión `.docx → .pdf` | `soffice.exe --version` | `winget install TheDocumentFoundation.LibreOffice` |

Detalles de instalación automática: `scripts/instalar-entorno.ps1`.

## Comandos de verificación/instalación

- **Preflight completo** (check + reporte tabla OK/FALTA, exit 0/1):
  ```powershell
  powershell -ExecutionPolicy Bypass -File "C:\Users\User\.agents\skills\generar-pdf-apa\scripts\comprobar-entorno.ps1"
  ```
- **Autoinstalación** de lo faltante (winget/npm/pip) y re-verificación:
  ```powershell
  powershell -ExecutionPolicy Bypass -File "C:\Users\User\.agents\skills\generar-pdf-apa\scripts\instalar-entorno.ps1"
  ```
  Si al final el re-check no da `ENTORNO OK`, **detener y avisar** (no transformar sin entorno completo).

## Ubicaciones comprometidas

- Generador del `.docx` (donde debe estar `node_modules\docx`): `C:\Users\User\AppData\Local\Temp\opencode\generar-pdf-apa`.
- venv de Docling: `C:\Users\User\Downloads\Docling\venv\Scripts\python.exe` (usa el `.exe` directo; el `activate` no persiste entre sesiones).
- LibreOffice: `C:\Program Files\LibreOffice\program\soffice.exe` (la ruta winget estándar).

## Notas operativas

- `winget` necesita `Refresh-Path` tras instalar para exponer `soffice`/`node` (o invocar por ruta completa).
- Las instalaciones requieren la ruta del **directorio del generador** por parámetro ↔ no asumas el directorio de trabajo actual para `npm install docx`; si el `docx` no está donde se ejecuta el generador, el `require('docx')` falla.
- El soundex de `verify.py` es delta-dos y ESLint no forma parte del pipeline: lo importante es el `ENTORNO OK` del preflight y al final el PDF verificado.