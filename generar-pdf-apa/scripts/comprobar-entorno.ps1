param(
    [string]$WorkDir = 'C:\Users\User\AppData\Local\Temp\opencode\generar-pdf-apa',
    [string]$PythonExe = 'C:\Users\User\Downloads\Docling\venv\Scripts\python.exe'
)
$ErrorActionPreference = 'SilentlyContinue'
$rows = @()

function Add-Row([string]$name, [bool]$ok, [string]$detail) {
    $script:rows += [pscustomobject]@{ Herramienta = $name; Estado = $(if ($ok) { 'OK' } else { 'FALTA' }); Detalle = $detail }
}

$psVer = $PSVersionTable.PSVersion.ToString()
Add-Row 'PowerShell' ($PSVersionTable.PSVersion -ge [version]'5.1') "v$psVer"

$node = & node --version 2>&1
Add-Row 'Node.js' ($LASTEXITCODE -eq 0) "$node"

$npm = & npm --version 2>&1
Add-Row 'npm' ($LASTEXITCODE -eq 0) "v$npm"

$docxPkg = Join-Path $WorkDir 'node_modules\docx\package.json'
$docxVer = ''
if (Test-Path -LiteralPath $docxPkg) {
    try { $docxVer = (Get-Content -LiteralPath $docxPkg -Raw | ConvertFrom-Json).version } catch {}
}
Add-Row 'docx (npm)' (Test-Path -LiteralPath $docxPkg) "v$docxVer en $WorkDir"

$pyVer = ''
if (Test-Path -LiteralPath $PythonExe) {
    $pyVer = & $PythonExe -c "import sys; print(sys.version.split()[0])" 2>&1
}
$pymupdf = ''
$pymupdfOk = $false
if (Test-Path -LiteralPath $PythonExe) {
    $pymupdf = & $PythonExe -c "import pymupdf; print(pymupdf.__doc__.split(chr(10))[0])" 2>&1
    $pymupdfOk = $LASTEXITCODE -eq 0
}
Add-Row 'Python venv' (Test-Path -LiteralPath $PythonExe) "$pyVer en $PythonExe"
Add-Row 'pymupdf' $pymupdfOk "$pymupdf"

$soffice = 'C:\Program Files\LibreOffice\program\soffice.exe'
$loDetail = ''
$loOk = Test-Path -LiteralPath $soffice
if ($loOk) {
    $loDetail = (& $soffice --version 2>&1 | Select-Object -First 1)
} else {
    $winget = & winget --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        $loList = & winget list --id TheDocumentFoundation.LibreOffice --accept-source-agreements 2>&1
        if ($LASTEXITCODE -eq 0 -and ($loList -match 'LibreOffice')) { $loDetail = 'instalado vía winget pero soffice.exe no encontrado' }
    }
}
Add-Row 'LibreOffice (F2)' $loOk "$loDetail"

$winget = & winget --version 2>&1
Add-Row 'winget' ($LASTEXITCODE -eq 0) "$winget"

$rows | Format-Table -AutoSize
$missing = @($rows | Where-Object { $_.Estado -eq 'FALTA' })
if ($missing.Count -eq 0) {
    Write-Output ''
    Write-Output 'ENTORNO OK: todas las herramientas requeridas estan presentes.'
    exit 0
} else {
    Write-Output ''
    Write-Output ("FALTAN: " + (($missing | ForEach-Object { $_.Herramienta }) -join ', '))
    Write-Output 'Ejecute scripts/instalar-entorno.ps1 para instalarlas automaticamente antes de transformar el documento.'
    exit 1
}