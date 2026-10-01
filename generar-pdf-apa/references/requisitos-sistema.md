# Requisitos del sistema

Herramientas que el pipeline necesita para convertir un `.md` en un `.docx` y
despues en un `.pdf`. **Microsoft Word no es un requisito**: la conversion la
hace LibreOffice en modo headless.

Todo lo que se necesita se resuelve solo, sin rutas escritas a mano. Este
archivo explica como se comprueba y como se ajusta si la deteccion falla.

## Herramientas

| Herramienta | Para que se usa | Como se comprueba | Instalacion automatica |
|---|---|---|---|
| PowerShell 5.1 o superior | ejecutar los scripts de la skill | `$PSVersionTable.PSVersion` | viene con Windows |
| Node.js LTS | generar el `.docx` con la libreria `docx` | `node --version` | `winget install OpenJS.NodeJS.LTS` o `brew install node` |
| libreria `docx` (npm) | construir portada, indices, tablas y figuras | se busca en el directorio de trabajo | `npm install docx@9.7.1 --no-save` |
| Python con `pymupdf` | verificar el `.pdf` resultante | `python -c "import fitz"` | `pip install pymupdf` |
| LibreOffice | **unico motor** de conversion `.docx` a `.pdf` | `soffice --version` | `winget install TheDocumentFoundation.LibreOffice` o `brew install --cask libreoffice` |

El instalador elige el gestor de paquetes disponible (`winget` en Windows,
`brew` en macOS/Linux) y despues vuelve a ejecutar la comprobacion.

## Preflight

Comprueba todo y devuelve codigo de salida 0 si el entorno esta completo, 1 si
falta algo. Se ejecuta siempre antes de transformar un documento.

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\comprobar-entorno.ps1"
```

- Si responde `ENTORNO OK`, se puede continuar.
- Si responde con la lista `FALTAN:`, se instala lo que falte:

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\instalar-entorno.ps1"
```

El instalador aplica los cambios y repite la comprobacion. Si al terminar sigue
sin dar `ENTORNO OK`, **detenerse y avisar al usuario**: no se transforma un
documento con el entorno incompleto, porque el fallo aparece mas tarde como un
PDF mal formado y es mucho mas dificil de diagnosticar.

## Rutas: como se resuelven

Ningun script tiene una ruta absoluta escrita dentro. Todas pasan por
`scripts/lib/rutas.ps1`, que busca en este orden:

1. La variable de entorno correspondiente, si esta definida.
2. Rutas de instalacion habituales del sistema.
3. El `PATH`.

| Variable | Que sobrescribe | Valor por defecto |
|---|---|---|
| `APA7_SKILL_ROOT` | raiz de la skill | la carpeta que contiene `SKILL.md` |
| `APA7_WORKDIR` | directorio de trabajo generado | `<raiz de la skill>\.work` |
| `APA7_NODEDIR` | donde vive `node_modules` | el directorio de trabajo |
| `APA7_SOFFICE` | ejecutable de LibreOffice | deteccion automatica |
| `APA7_PYTHON` | interprete Python con `pymupdf` | deteccion automatica |

Ejemplo: si la libreria `docx` esta en otro sitio porque el proyecto ya tenia
sus propias dependencias,

```powershell
$env:APA7_NODEDIR = "C:\ruta\al\proyecto\node_modules"
```

### Por que existe el directorio de trabajo

La libreria `docx` se instala en un directorio propio de la skill y no junto al
documento del usuario, para que `require('docx')` resuelva siempre sin importar
desde donde se ejecute el generador. Por defecto ese directorio es `.work/` y
solo contiene estado generado:

- `node_modules/` con la libreria `docx`.
- (antes tambien un perfil de LibreOffice; ya no, cada corrida crea el suyo
  temporal junto al PDF de salida).

Se regenera con `instalar-entorno.ps1` y esta en `.gitignore`. **No es codigo
fuente y no debe editarse ni entregarse.**

## Notas operativas

- Tras instalar con `winget` hace falta refrescar el `PATH` antes de invocar
  `node` o `soffice`; el instalador ya lo hace, pero si se instala a mano hay que
  reabrir la terminal.
- El `PATH` de la sesion y el del proceso que lanza el agente son distintos: por
  eso los scripts buscan las rutas por su cuenta y no asumen que `node` este
  visible.
- Si `node` esta instalado pero `docx` no, es un fallo distinto al anterior y el
  preflight lo reporta por separado.
- La deteccion de Python prueba varios interpretes, porque puede haber mas de uno
  instalado y no todos tienen `pymupdf`. Se puede fijar con `APA7_PYTHON`.
