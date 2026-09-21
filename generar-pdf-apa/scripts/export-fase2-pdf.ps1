param(
    [Parameter(Mandatory = $true)][string]$SrcDocx,
    [string]$OutDir,
    [string]$OutPdf,
    [string]$Soffice = 'C:\Program Files\LibreOffice\program\soffice.exe',
    [string]$WorkDir = 'C:\Users\User\AppData\Local\Temp\opencode\export-work'
)
$ErrorActionPreference = 'Stop'
$base = [IO.Path]::GetFileNameWithoutExtension($SrcDocx)
if (-not $OutDir) { $OutDir = Split-Path -Parent $SrcDocx }
if (-not $OutPdf) { $OutPdf = Join-Path $OutDir ($base + '.pdf') }

function Step($m) { Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

Step "START-2 | SrcDocx=$SrcDocx"
if (-not (Test-Path -LiteralPath $SrcDocx)) { Write-Output 'ERROR: no existe el docx de entrada'; exit 1 }
if (-not (Test-Path -LiteralPath $Soffice)) { Write-Output "ERROR: no se encontro sofice en $Soffice. Ejecute scripts/instalar-entorno.ps1"; exit 1 }

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
$copyDir = Join-Path $WorkDir ("lo_copy_" + $PID)
New-Item -ItemType Directory -Path $copyDir -Force | Out-Null
$tmpDocx = Join-Path $copyDir ([IO.Path]::GetFileName($SrcDocx))
Copy-Item -LiteralPath $SrcDocx -Destination $tmpDocx -Force

$profileDir = Join-Path $WorkDir ("lo_profile_" + $PID)   # perfil aislado por corrida
$profileUri = ('file:///' + ($profileDir -replace '\\', '/'))

$stderrLog = Join-Path $WorkDir 'export-fase2.log'
$proc = Start-Process -FilePath $Soffice -ArgumentList @(
    '--headless', '--norestore', '--nolockcheck',
    "-env:UserInstallation=$profileUri",
    '--convert-to', 'pdf', '--outdir', $copyDir, $tmpDocx
) -PassThru -NoNewWindow -Wait -RedirectStandardError $stderrLog
# Nota: "Could not find platform independent libraries <prefix>" es un aviso benigno de LibreOffice
# que sale por stderr (va a export-fase2.log, no a la consola). No es un error de conversion.
if ($proc.ExitCode -ne 0) { Write-Output ('ERROR: LibreOffice fallo en la conversion (ExitCode ' + $proc.ExitCode + ', ver export-fase2.log)'); exit 1 }

$tmpPdf = Join-Path $copyDir ([IO.Path]::GetFileName($SrcDocx) -replace '\.docx$', '.pdf')
$ready = $false
for ($t = 0; $t -lt 120; $t++) {
    if (Test-Path -LiteralPath $tmpPdf) { $ready = $true; break }
    Start-Sleep -Milliseconds 500
}
if (-not $ready) { Write-Output 'ERROR: no se genero el PDF (timeout)'; exit 1 }

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
Copy-Item -LiteralPath $tmpPdf -Destination $OutPdf -Force
Remove-Item -LiteralPath $copyDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue
Step "DONE | PDF -> $OutPdf ($((Get-Item -LiteralPath $OutPdf).Length) bytes)"