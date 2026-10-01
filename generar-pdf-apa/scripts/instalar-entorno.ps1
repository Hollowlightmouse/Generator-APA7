<#
    instalar-entorno.ps1 - Self-install of whatever is missing (STEP 0)

    Installs Node.js, the docx library, pymupdf and LibreOffice using the system
    package manager: winget on Windows, brew on macOS/Linux. It contains no
    absolute paths of any machine: everything is resolved with lib\rutas.ps1.

    It is idempotent: if something is already present, it is not reinstalled.
    At the end it runs the preflight again and returns the exit code of
    comprobar-entorno.ps1 (0 = full environment, 1 = something is still missing).

    Optional environment variables:
        APA7_SOFFICE, APA7_PYTHON, APA7_WORKDIR, APA7_NODEDIR
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot 'lib\rutas.ps1')

$checkScript = Join-Path $PSScriptRoot 'comprobar-entorno.ps1'

function Step { param([string]$M) Write-Apa7Log $M }

function Refresh-Path {
    # After installing, neither the session nor the agent sees the new PATH.
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($machine) { $env:Path = "$machine;$user;$env:Path" }
}

# --- 0. Initial situation --------------------------------------------------
Step '--- Initial environment check ---'
& $checkScript
$initial = $LASTEXITCODE
if ($initial -eq 0) { Step 'Nothing to install: environment complete.'; exit 0 }
Step 'Tools missing. Trying to install them...'

$pm = Get-PackageManagerCmd
if (-not $pm) {
    Write-Output ''
    Write-Output 'ERROR: no package manager (winget on Windows, brew on macOS/Linux).'
    Write-Output 'Automatic installation cannot run. Install by hand whatever is missing'
    Write-Output 'according to references\system-requirements.md and run the preflight again.'
    exit 1
}
Step ("Package manager available: " + $pm.Name)

# --- 1. Node.js + npm ------------------------------------------------------
$node = Get-NodeCmd
if (-not $node) {
    Step 'Installing Node.js LTS...'
    if ($pm.Name -eq 'winget') {
        & $pm.Path install --id OpenJS.NodeJS.LTS -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity | Out-Null
    } else {
        & $pm.Path install node | Out-Null
    }
    Refresh-Path
    $node = Get-NodeCmd
    if (-not $node) { Write-Output 'FAILED: could not install Node.js'; exit 1 }
}
Step ('Node.js present: ' + (& $node --version))

$npm = Get-NpmCmd
if (-not $npm) { Write-Output 'FAILED: npm not available after installing Node.js'; exit 1 }
Step ('npm present: v' + ((& $npm --version) -join ''))

# --- 2. docx library -------------------------------------------------------
# BUG FIXED: it used to Push-Location to a directory that might not exist, and
# npm install failed with an unclear error.
#
# A minimal package.json is written inside the workdir on purpose: without it,
# `npm install` walks UP the directory tree looking for a package.json and can
# drop node_modules in an unexpected parent folder. With the anchor file, docx
# always lands in <workdir>\node_modules, which is exactly where the preflight
# and build-docx.js look for it.
$nodeDir = Get-NodeDir
if (-not (Test-Path -LiteralPath $nodeDir)) {
    New-Item -ItemType Directory -Path $nodeDir -Force | Out-Null
    Step "Created workdir: $nodeDir"
}
$anchor = Join-Path $nodeDir 'package.json'
if (-not (Test-Path -LiteralPath $anchor)) {
    $anchorJson = '{ "name": "apa7-workdir", "private": true, "version": "1.0.0", ' +
                  '"description": "Anchors npm so that node_modules/docx installs here." }'
    Set-Content -LiteralPath $anchor -Value $anchorJson -Encoding UTF8
    Step "Created npm anchor: $anchor"
}
$modDir = Join-Path $nodeDir 'node_modules\docx'
if (-not (Test-Path -LiteralPath $modDir)) { $modDir = Join-Path $nodeDir 'node_modules/docx' }

if (-not (Test-Path -LiteralPath $modDir)) {
    Step "Installing docx@9.7.1 in $nodeDir (npm install)..."
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    Push-Location $nodeDir
    try { $installOut = & $npm install docx@9.7.1 --no-save --no-package-lock --no-audit --no-fund 2>&1 }
    finally { Pop-Location }
    $ErrorActionPreference = $prevEap
    if (-not (Test-Path -LiteralPath $modDir)) {
        Write-Output 'FAILED to install docx (npm). npm output:'
        $installOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
        exit 1
    }
    Step 'docx (npm) installed.'
} else { Step 'docx (npm) present.' }

# --- 3. pymupdf ------------------------------------------------------------
# The concrete interpreter is NOT fixed by the skill: it uses APA7_PYTHON if
# defined, otherwise the first WORKING one found. If none has pymupdf, it
# reports the exact command for each system instead of a generic failure.
$py = $env:APA7_PYTHON
if (-not $py) { $py = Get-PythonPath }
if (-not $py) {
    Write-Output ''
    Write-Output 'FAILED: no working Python interpreter available.'
    Write-Output 'Install Python and then run one of these commands:'
    if (Test-Apa7IsMac) {
        Write-Output '    python3 -m venv <folder>/venv'
        Write-Output '    <folder>/venv/bin/python -m pip install pymupdf'
    } else {
        Write-Output '    python -m venv <folder>\.venv'
        Write-Output '    <folder>\.venv\Scripts\python -m pip install pymupdf'
    }
    Write-Output 'Then set APA7_PYTHON to that path and run the preflight again.'
    exit 1
}

$r = Invoke-Native -FilePath $py -Arguments @('-c', 'import pymupdf; print(pymupdf.__version__)')
if ($r.ExitCode -ne 0) {
    Step "Installing pymupdf with $py ..."
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $pipOut = & $py -m pip install pymupdf 2>&1
    $pipCode = $LASTEXITCODE
    $ErrorActionPreference = $prevEap
    if ($pipCode -ne 0) {
        Write-Output 'FAILED to install pymupdf. pip output:'
        $pipOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
        exit 1
    }
    Step 'pymupdf installed.'
} else {
    Step ('pymupdf present: ' + $r.First)
}

# --- 4. LibreOffice (single PDF engine) ------------------------------------
$soffice = Get-SofficePath
if (-not $soffice) {
    Step 'Installing LibreOffice (this can take several minutes)...'
    if ($pm.Name -eq 'winget') {
        & $pm.Path install --id TheDocumentFoundation.LibreOffice -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity | Out-Null
    } else {
        & $pm.Path install --cask libreoffice | Out-Null
    }
    Refresh-Path
    $soffice = Get-SofficePath
    if (-not $soffice) {
        Write-Output ''
        Write-Output 'FAILED: LibreOffice not found after installation.'
        Write-Output 'Paths examined: ' + ((Get-SofficeCandidates | Where-Object { $_ }) -join ' | ')
        Write-Output 'If it is installed elsewhere, set APA7_SOFFICE to the full path.'
        exit 1
    }
    Step ('LibreOffice installed: ' + $soffice)
} else { Step ('LibreOffice present: ' + $soffice) }

# --- 5. Final check --------------------------------------------------------
Step '--- Final environment check ---'
& $checkScript
exit $LASTEXITCODE
