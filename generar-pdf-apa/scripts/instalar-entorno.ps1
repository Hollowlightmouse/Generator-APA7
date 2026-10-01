<#
    instalar-entorno.ps1 - Autoinstalacion de lo que falte (PASO 0)

    Instala Node.js, la libreria docx, pymupdf y LibreOffice usando el gestor
    de paquetes del sistema: winget en Windows, brew en macOS/Linux. No contiene
    rutas absolutas de ninguna maquina: todo se resuelve con lib\rutas.ps1.

    Es idempotente: si algo ya esta, no lo reinstala. Al final vuelve a pasar
    el preflight y devuelve el codigo de salida de comprobar-entorno.ps1
    (0 = entorno completo, 1 = sigue faltando algo).

    Variables de entorno opcionales:
        APA7_SOFFICE, APA7_PYTHON, APA7_WORKDIR, APA7_NODEDIR
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot 'lib\rutas.ps1')

$checkScript = Join-Path $PSScriptRoot 'comprobar-entorno.ps1'

function Step { param([string]$M) Write-Apa7Log $M }

function Refresh-Path {
    # Tras instalar, ni la sesion ni el agente ven el PATH nuevo.
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($machine) { $env:Path = "$machine;$user;$env:Path" }
}

# --- 0. Situacion inicial -------------------------------------------------
Step '--- Verificacion inicial del entorno ---'
& $checkScript
$initial = $LASTEXITCODE
if ($initial -eq 0) { Step 'Nada que instalar: entorno completo.'; exit 0 }
Step 'Faltan herramientas. Intentando instalarlas...'

$pm = Get-PackageManagerCmd
if (-not $pm) {
    Write-Output ''
    Write-Output 'ERROR: no hay gestor de paquetes (winget en Windows, brew en macOS/Linux).'
    Write-Output 'La instalacion automatica no puede ejecutarse. Instale a mano lo que falte'
    Write-Output 'segun references\requisitos-sistema.md y vuelva a ejecutar el preflight.'
    exit 1
}
Step ("Gestor de paquetes disponible: " + $pm.Name)

# --- 1. Node.js + npm -----------------------------------------------------
$node = Get-NodeCmd
if (-not $node) {
    Step 'Instalando Node.js LTS...'
    if ($pm.Name -eq 'winget') {
        & $pm.Path install --id OpenJS.NodeJS.LTS -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity | Out-Null
    } else {
        & $pm.Path install node | Out-Null
    }
    Refresh-Path
    $node = Get-NodeCmd
    if (-not $node) { Write-Output 'FALLO: no se pudo instalar Node.js'; exit 1 }
}
Step ('Node.js presente: ' + (& $node --version))

$npm = Get-NpmCmd
if (-not $npm) { Write-Output 'FALLO: npm no disponible tras instalar Node.js'; exit 1 }
Step ('npm presente: v' + ((& $npm --version) -join ''))

# --- 2. Libreria docx -----------------------------------------------------
# BUG CORREGIDO: antes se hacia Push-Location a un directorio que podia no
# existir, y npm install fallaba con un error poco claro.
$nodeDir = Get-NodeDir
if (-not (Test-Path -LiteralPath $nodeDir)) {
    New-Item -ItemType Directory -Path $nodeDir -Force | Out-Null
    Step "Creado directorio de trabajo: $nodeDir"
}
$modDir = Join-Path $nodeDir 'node_modules\docx'
if (-not (Test-Path -LiteralPath $modDir)) { $modDir = Join-Path $nodeDir 'node_modules/docx' }

if (-not (Test-Path -LiteralPath $modDir)) {
    Step "Instalando docx@9.7.1 en $nodeDir (npm install)..."
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    Push-Location $nodeDir
    try { $installOut = & $npm install docx@9.7.1 --no-save --no-audit --no-fund 2>&1 }
    finally { Pop-Location }
    $ErrorActionPreference = $prevEap
    if (-not (Test-Path -LiteralPath $modDir)) {
        Write-Output 'FALLO al instalar docx (npm). Salida de npm:'
        $installOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
        exit 1
    }
    Step 'docx (npm) instalado.'
} else { Step 'docx (npm) presente.' }

# --- 3. pymupdf -----------------------------------------------------------
# El interprete concreto NO lo fija la skill: usa APA7_PYTHON si esta definido,
# si no el primero encontrado. Si ninguno trae pymupdf, se informa con el
# comando exacto para cada sistema, en lugar de fallar con un mensaje generico.
$py = $env:APA7_PYTHON
if (-not $py) { $py = Get-PythonPath }
if (-not $py) {
    Write-Output ''
    Write-Output 'FALLO: no hay ningun interprete Python disponible.'
    Write-Output 'Instale Python y despues ejecute uno de estos comandos:'
    if (Test-Apa7IsMac) {
        Write-Output '    python3 -m venv <carpeta>\venv'
        Write-Output '    <carpeta>\venv\bin\python -m pip install pymupdf'
    } else {
        Write-Output '    python -m venv <carpeta>\.venv'
        Write-Output '    <carpeta>\.venv\Scripts\python -m pip install pymupdf'
    }
    Write-Output 'Despues defina APA7_PYTHON con esa ruta y vuelva a ejecutar el preflight.'
    exit 1
}

$r = Invoke-Native -FilePath $py -Arguments @('-c', 'import pymupdf; print(pymupdf.__version__)')
if ($r.ExitCode -ne 0) {
    Step "Instalando pymupdf con $py ..."
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $pipOut = & $py -m pip install pymupdf 2>&1
    $pipCode = $LASTEXITCODE
    $ErrorActionPreference = $prevEap
    if ($pipCode -ne 0) {
        Write-Output 'FALLO al instalar pymupdf. Salida de pip:'
        $pipOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
        exit 1
    }
    Step 'pymupdf instalado.'
} else {
    Step ('pymupdf presente: ' + $r.First)
}

# --- 4. LibreOffice (motor unico de PDF) ---------------------------------
$soffice = Get-SofficePath
if (-not $soffice) {
    Step 'Instalando LibreOffice (puede tardar varios minutos)...'
    if ($pm.Name -eq 'winget') {
        & $pm.Path install --id TheDocumentFoundation.LibreOffice -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity | Out-Null
    } else {
        & $pm.Path install --cask libreoffice | Out-Null
    }
    Refresh-Path
    $soffice = Get-SofficePath
    if (-not $soffice) {
        Write-Output ''
        Write-Output 'FALLO: no se encontro LibreOffice tras la instalacion.'
        Write-Output 'Rutas examinate: ' + ((Get-SofficeCandidates | Where-Object { $_ }) -join ' | ')
        Write-Output 'Si esta instalado en otro sitio, defina APA7_SOFFICE con la ruta completa.'
        exit 1
    }
    Step ('LibreOffice instalado: ' + $soffice)
} else { Step ('LibreOffice presente: ' + $soffice) }

# --- 5. Verificacion final ------------------------------------------------
Step '--- Verificacion final del entorno ---'
& $checkScript
exit $LASTEXITCODE
