<#
    comprobar-entorno.ps1 - Preflight (PASO 0) de la skill generar-pdf-apa

    Comprueba que existan TODAS las herramientas necesarias antes de tocar un
    documento. Si algo falta, avisa y detiene el pipeline (nunca se transforma
    un documento con herramientas incompletas).

    Motor de PDF: LibreOffice. No se requiere Microsoft Word. Tipst NO es
    necesario (ver references/decisiones-motor.md).

    Salida: una linea por herramienta, en formato parseable:
        OK|<herramienta>|<detalle>
        FALTA|<herramienta>|<detalle>
    followed by:
        RESULTADO: OK | FALTA
    Codigo de salida: 0 = entorno completo, 1 = falta algo.

    Este script NO contiene rutas absolutas de ninguna maquina. Todo se
    resuelve con lib\rutas.ps1 (entorno -> PATH -> rutas tipicas del SO).

    Variables de entorno opcionales:
        APA7_SOFFICE, APA7_PYTHON, APA7_WORKDIR, APA7_NODEDIR
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot 'lib\rutas.ps1')

$results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param([string]$Tool, [bool]$Ok, [string]$Detail)
    $state = if ($Ok) { 'OK' } else { 'FALTA' }
    $results.Add([pscustomobject]@{ Estado = $state; Herramienta = $Tool; Detalle = $Detail })
    Write-Output ("{0}|{1}|{2}" -f $state, $Tool, $Detail)
}

# --- PowerShell ------------------------------------------------------------
Add-Result 'PowerShell' ($PSVersionTable.PSVersion -ge [version]'5.1') ('v' + $PSVersionTable.PSVersion)

# --- Node.js ---------------------------------------------------------------
# Se resuelve con Get-Command. Consultar $LASTEXITCODE cuando el comando NO
# existe arrastra el valor del comando anterior y puede dar un OK falso: ese
# era un fallo real de la version anterior de este script.
$node = Get-NodeCmd
if ($node) {
    $r = Invoke-Native -FilePath $node -Arguments @('--version')
    Add-Result 'Node.js' ($r.ExitCode -eq 0) ("$($r.First)  ($node)")
} else {
    Add-Result 'Node.js' $false 'no encontrado en PATH'
}

# --- npm -------------------------------------------------------------------
$npm = Get-NpmCmd
if ($npm) {
    $r = Invoke-Native -FilePath $npm -Arguments @('--version')
    Add-Result 'npm' ($r.ExitCode -eq 0) ("v$($r.First)  ($npm)")
} else {
    Add-Result 'npm' $false 'no encontrado en PATH'
}

# --- libreria docx (npm) ---------------------------------------------------
# Comprobacion REAL: se pide a Node que resuelva el modulo desde el directorio
# de trabajo. Verificar solo que existe node_modules\docx\package.json daba
# falsos positivos cuando el modulo estaba instalado en otro sitio, y fue
# justamente lo que rompio el pipeline (el generador lo requeria con una ruta
# relativa distinta de la que usaba el instalador).
$nodeDir = Get-NodeDir
$modDir = Join-Path $nodeDir 'node_modules\docx'
if (-not (Test-Path -LiteralPath $modDir)) { $modDir = Join-Path $nodeDir 'node_modules/docx' }
$pkgJson = Join-Path $modDir 'package.json'

if (-not $node) {
    Add-Result 'docx (npm)' $false 'no verificable: falta Node.js'
} elseif (-not (Test-Path -LiteralPath $pkgJson)) {
    Add-Result 'docx (npm)' $false "no instalado en $nodeDir (ejecute instalar-entorno.ps1)"
} else {
    $ver = ''
    try { $ver = (Get-Content -LiteralPath $pkgJson -Raw -Encoding UTF8 | ConvertFrom-Json).version } catch { }
    $posixMod = $modDir.Replace('\', '/')
    $probe = "require('$posixMod'); console.log('ok')"
    $r = Invoke-Native -FilePath $node -Arguments @('-e', $probe)
    Add-Result 'docx (npm)' ($r.ExitCode -eq 0 -and $r.First -eq 'ok') ("v$ver  ($nodeDir)")
}

# --- Python + pymupdf ------------------------------------------------------
# El venv concreto ya NO se fija dentro de la skill: se resuelve por variable
# de entorno o por PATH. Ver references/requisitos-sistema.md.
$py = $env:APA7_PYTHON
if (-not $py) { $py = Get-PythonPath }

if (-not $py) {
    Add-Result 'Python' $false 'no encontrado (defina APA7_PYTHON con un interprete que tenga pymupdf)'
    Add-Result 'pymupdf' $false 'no verificable: falta Python'
} else {
    $r = Invoke-Native -FilePath $py -Arguments @('-c', 'import sys; print(sys.version.split()[0])')
    Add-Result 'Python' ($r.ExitCode -eq 0) ("$($r.First)  ($py)")

    $r2 = Invoke-Native -FilePath $py -Arguments @('-c', 'import pymupdf; print(pymupdf.__version__)')
    Add-Result 'pymupdf' ($r2.ExitCode -eq 0) ("$($r2.First)  (interprete: $py)")
}

# --- LibreOffice (motor unico de PDF) -------------------------------------
# IMPORTANTE: se consulta la version con Invoke-Soffice, nunca con
# `& soffice.exe --version`. soffice.exe se desprende, el hijo hereda el pipe
# de salida y PowerShell espera indefinidamente: eso cuelga el script entero.
$soffice = Get-SofficePath
if ($soffice) {
    $r = Invoke-Soffice -Arguments @('--version') -TimeoutSeconds 60
    $version = ($r.StdOut -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
    $ok = ($r.ExitCode -eq 0) -and ([string]::IsNullOrWhiteSpace($version) -eq $false)
    if (-not $ok) { $version = 'no se pudo leer la version (ExitCode ' + $r.ExitCode + ')' }
    Add-Result 'LibreOffice' $ok ("$version  ($soffice)")
} else {
    $cands = (Get-SofficeCandidates | Where-Object { $_ }) -join ' | '
    if (-not $cands) { $cands = '(sin candidatos: defina APA7_SOFFICE)' }
    Add-Result 'LibreOffice' $false "no encontrado. Buscado: $cands"
}

# --- Gestor de paquetes (necesario solo para la autoinstalacion) -----------
$pm = Get-PackageManagerCmd
if ($pm) {
    Add-Result 'Gestor de paquetes' $true ("$($pm.Name)  ($($pm.Path))")
} else {
    Add-Result 'Gestor de paquetes' $false 'sin winget (Windows) ni brew (macOS/Linux): la instalacion automatica no podra ejecutarse; instale a mano segun references/requisitos-sistema.md'
}

# --- Resumen ---------------------------------------------------------------
$missing = @($results | Where-Object { $_.Estado -eq 'FALTA' })
Write-Output ''
if ($missing.Count -eq 0) {
    Write-Output 'RESULTADO: OK'
    Write-Output 'ENTORNO OK: todas las herramientas requeridas estan presentes.'
    exit 0
} else {
    Write-Output ('FALTAN: ' + (($missing | ForEach-Object { $_.Herramienta }) -join ', '))
    Write-Output 'RESULTADO: FALTA'
    Write-Output 'Ejecute scripts\instalar-entorno.ps1 para instalarlas automaticamente. Si no puede, detenga el pipeline e informe al usuario.'
    exit 1
}
