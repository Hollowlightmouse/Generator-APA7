param(
    [string]$WorkDir = 'C:\Users\User\AppData\Local\Temp\opencode\generar-pdf-apa',
    [string]$PythonExe = 'C:\Users\User\Downloads\Docling\venv\Scripts\python.exe'
)
$ErrorActionPreference = 'Stop'
$skillDir = Split-Path -Parent $PSScriptRoot
$envscript = Join-Path $skillDir 'scripts\comprobar-entorno.ps1'

function Step($m) { Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
function Refresh-Path {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machine;$user;$env:Path"
}

Step '--- Verificacion inicial del entorno ---'
& $envscript
$code = $LASTEXITCODE
if ($code -eq 0) { Step 'Nada que instalar.'; exit 0 }
Step 'Instalando herramientas faltantes...'

# winget debe existir para las instalaciones por winget
$winget = & winget --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Output 'ERROR: winget no esta disponible. Instale App Installer desde Microsoft Store o manualmente los requisitos.'
    exit 1
}

# --- Node.js ---
$node = & node --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Step 'Instalando Node.js LTS via winget...'
    & winget install --id OpenJS.NodeJS.LTS -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity
    if ($LASTEXITCODE -ne 0) { Write-Output 'FALLO al instalar Node.js'; exit 1 }
    Refresh-Path
    Step ('Node.js instalado: ' + (& node --version 2>&1))
} else { Step ('Node.js presente: ' + $node) }

# --- npm (depende de node) ---
$npm = & npm --version 2>&1
if ($LASTEXITCODE -ne 0) { Write-Output 'FALLO: npm no disponible'; exit 1 }
Step ('npm presente: v' + $npm)

# --- libreria docx ---
$docxPkg = Join-Path $WorkDir 'node_modules\docx\package.json'
if (-not (Test-Path -LiteralPath $docxPkg)) {
    Step "Instalando docx@9.7.1 en $WorkDir (npm install --no-save)..."
    Push-Location $WorkDir
    try { & npm install docx@9.7.1 --no-save }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $docxPkg)) { Write-Output 'FALLO al instalar docx (npm)'; exit 1 }
    Step 'docx (npm) instalado.'
} else { Step 'docx (npm) presente.' }

# --- Python venv + pymupdf ---
if (-not (Test-Path -LiteralPath $PythonExe)) {
    Write-Output "FALLO: no existe el venv Python en $PythonExe"
    Write-Output 'Nota: la skill usa el venv del proyecto Docling (venv\Scripts\python.exe). Active el venv correspondiente y vuelva a ejecutar.'
    exit 1
}
$pyImport = & $PythonExe -c "import pymupdf; print(pymupdf.__version__)" 2>&1
if ($LASTEXITCODE -ne 0) {
    Step "Instalando pymupdf en el venv ($PythonExe)..."
    & $PythonExe -m pip install pymupdf
    if ($LASTEXITCODE -ne 0) { Write-Output 'FALLO al instalar pymupdf'; exit 1 }
    Step 'pymupdf instalado.'
} else { Step ("pymupdf presente: " + $pyImport) }

# --- LibreOffice (F2, motor obligatorio) ---
$soffice = 'C:\Program Files\LibreOffice\program\soffice.exe'
if (-not (Test-Path -LiteralPath $soffice)) {
    Step 'Instalando LibreOffice via winget (incluye VCRedist como dependencia; puede tardar varios minutos)...'
    & winget install --id TheDocumentFoundation.LibreOffice -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity
    if ($LASTEXITCODE -ne 0) { Write-Output 'FALLO al instalar LibreOffice. Reintente manualmente o desde Microsoft Store.'; exit 1 }
    Step ("LibreOffice instalado: " + (& "$soffice" --version 2>&1 | Select-Object -First 1))
} else { Step ("LibreOffice presente: " + (& $soffice --version 2>&1 | Select-Object -First 1)) }

Step '--- Verificacion final del entorno ---'
& $envscript
exit $LASTEXITCODE