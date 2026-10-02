<#
    instalar-entorno.ps1 - Self-install of whatever is missing (STEP 0)

    Installs Node.js, the docx library, pymupdf and LibreOffice using the system
    package manager: winget on Windows, brew on macOS/Linux. It contains no
    absolute paths of any machine: everything is resolved with lib\rutas.ps1.

    It is idempotent: if something is already present, it is not reinstalled.
    At the end it runs the preflight again and returns the exit code of
    comprobar-entorno.ps1 (0 = full environment, 1 = something is still missing).

    Python is installed into a VIRTUAL ENVIRONMENT inside the skill folder
    (<skill>\.venv) with pymupdf pinned to lib\rutas.ps1 $script:APA7_DEPS. That
    keeps the library off the system interpreter (so a PEP 668 "externally
    managed" Python can no longer block the install) and makes the version
    reproducible. APA7_PYTHON still overrides which interpreter is used.

    Parameters:
        -Only  <tools>  install only these of: node, docx, python, pymupdf,
                       libreoffice. Default: all of them.
    Supports -WhatIf: reports what would be installed without installing.

    Optional environment variables:
        APA7_SOFFICE, APA7_PYTHON, APA7_WORKDIR, APA7_NODEDIR
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$Only = @()
)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot 'lib\rutas.ps1')

$checkScript = Join-Path $PSScriptRoot 'comprobar-entorno.ps1'

function Step { param([string]$M) Write-Apa7Log $M }

function ShouldInstall {
    param([string]$Tool)
    if (-not $Only -or $Only.Count -eq 0) { return $true }
    foreach ($t in $Only) { if ($t.ToLower() -eq $Tool.ToLower()) { return $true } }
    return $false
}

function Confirm-Install {
    # Returns $true when the action may run. Under -WhatIf it logs what it WOULD
    # do and returns $false. Kept as a plain boolean (no scriptblock argument)
    # because a scriptblock inside an if() condition does not parse reliably.
    #
    # The notice goes to the HOST on purpose (Write-Host, not Write-Apa7Log).
    # Write-Apa7Log uses Write-Output, so calling it here would append the log
    # line to this function's output: the result would be the array
    # @('[WhatIf] would: …', $false), and `if (<array>)` is TRUE for any
    # non-empty array in PowerShell. The installation would then run during the
    # very dry run that was supposed to prevent it. This function must emit
    # exactly one boolean and nothing else.
    param([string]$Description)
    $approved = [bool]$PSCmdlet.ShouldProcess($Description, 'Execute')
    if (-not $approved) { Write-Host ("[WhatIf] would: $Description") }
    return $approved
}

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
if (-not $node -and (ShouldInstall 'node')) {
    if (Confirm-Install 'Install Node.js LTS') {
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
} elseif (-not $node) {
    Step 'Node.js is missing and -Only excludes it: skipping.'
    exit 1
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
    $docxPin = $script:APA7_DEPS['docx']
    # Both guards are needed: -Only selects the tool, -WhatIf forbids the action.
    # Without the second one, `npm install` really runs during a dry run.
    if ((ShouldInstall 'docx') -and (Confirm-Install "npm install docx@$docxPin in $nodeDir")) {
        Step "Installing docx@$docxPin in $nodeDir (npm install)..."
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        Push-Location $nodeDir
        try { $installOut = & $npm install "docx@$docxPin" --no-save --no-package-lock --no-audit --no-fund 2>&1 }
        finally { Pop-Location }
        $ErrorActionPreference = $prevEap
        if (-not (Test-Path -LiteralPath $modDir)) {
            Write-Output 'FAILED to install docx (npm). npm output:'
            $installOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
            exit 1
        }
        Step 'docx (npm) installed.'
    } else {
        # Two different reasons land here, so they are reported separately: a
        # silent "skipped" makes a dry run look like a filtered run.
        if (ShouldInstall 'docx') {
            Step 'docx is missing and -WhatIf forbids installing it: skipped.'
        } else {
            Step 'docx is missing and -Only excludes it: skipping.'
        }
    }
} else { Step 'docx (npm) present.' }

# --- 3. Python interpreter -------------------------------------------------
# Previously the installer DID NOT install Python: when no interpreter was found
# it just printed manual instructions and exited 1. That made the skill
# uninstallable on a clean machine (which is exactly the case on the machine
# where this was written: only the Microsoft Store stub was on PATH). Node,
# docx and LibreOffice were installed automatically, Python was not.
#
# The venv in the next step needs a real base interpreter, so this step comes
# first. It is after Node on purpose: the Windows Python installer adds shims
# that need to be on PATH, and doing it after Refresh-Path keeps them visible.
$pyBase = $null
$venvExe = Get-PythonVenvPath
if ($venvExe -and (Test-Apa7PythonWorks -File $venvExe)) {
    # Already good: no base interpreter needed at all.
    $pyBase = $venvExe
    Step ("Virtual environment usable: $venvExe")
}

if (-not $pyBase) {
    $pyBase = $env:APA7_PYTHON
    if ($pyBase -and -not (Test-Apa7PythonWorks -File $pyBase)) { $pyBase = $null }
}
if (-not $pyBase) { $pyBase = Get-PythonPath }

if (-not $pyBase) {
    if (-not (ShouldInstall 'python')) {
        Write-Output ''
        Write-Output 'FAILED: no working Python interpreter, and -Only excludes installing it.'
        Write-Output 'Install Python and run this script again.'
        exit 1
    }
    if (Confirm-Install 'Install Python 3.12') {
        Step 'Installing Python 3.12 (no interpreter found on this machine)...'
        if ($pm.Name -eq 'winget') {
            & $pm.Path install --id Python.Python.3.12 -e --accept-package-agreements --accept-source-agreements --silent --disable-interactivity | Out-Null
        } else {
            & $pm.Path install python@3.12 | Out-Null
        }
        Refresh-Path
        $pyBase = Get-PythonPath
        if (-not $pyBase) {
            Write-Output ''
            Write-Output 'FAILED: Python was not usable after the installation.'
            Write-Output 'Install it by hand (references\system-requirements.md) and retry.'
            exit 1
        }
    }
}
Step ('Python interpreter available: ' + $pyBase)

# --- 4. venv + pymupdf (pinned) ---------------------------------------------
# pymupdf used to be installed with a bare `pip install pymupdf` against whatever
# interpreter was found: unpinned, and against the SYSTEM interpreter, which on a
# PEP 668 "externally managed" Python (Debian, Homebrew, some 3.13+ distros)
# fails with error: externally-managed-environment. The instructions to build a
# venv by hand only printed in the "no Python at all" branch, never on that
# failure, so it was a dead end.
#
# Installing into <skill>\.venv fixes both: the library is pinned, and pip is
# never run against the system interpreter.
$venvDir = Join-Path (Get-SkillRoot) '.venv'
$venvExe = if (Test-Apa7IsMac) { Join-Path $venvDir 'bin/python3' } else { Join-Path $venvDir 'Scripts\python.exe' }

if ((Test-Path -LiteralPath $venvExe) -and -not (Test-Apa7PythonWorks -File $venvExe)) {
    # A venv records the absolute path of its base interpreter (pyvenv.cfg), so
    # if that interpreter was removed or upgraded, the venv exists but cannot
    # run. It cannot be repaired: it is deleted and rebuilt.
    Step 'The virtual environment exists but does not run (its base interpreter moved or was upgraded). Recreating it...'
    Remove-Item -LiteralPath $venvDir -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $venvExe)) {
    if (-not (ShouldInstall 'pymupdf')) {
        Write-Output ''
        Write-Output "FAILED: no virtual environment at $venvDir and -Only excludes creating it."
        exit 1
    }
    if (Confirm-Install "Create the virtual environment at $venvDir") {
        Step "Creating the virtual environment: $venvDir"
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $venvOut = & $pyBase -m venv $venvDir 2>&1
        $ErrorActionPreference = $prevEap
        if (-not (Test-Path -LiteralPath $venvExe)) {
            Write-Output 'FAILED to create the virtual environment. Output:'
            $venvOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
            exit 1
        }
        Step 'Virtual environment created.'
    }
} else { Step ("Virtual environment present: $venvExe") }

# From here on the venv interpreter is the one that matters, whatever the
# override or the discovery said.
$py = $venvExe
$pymupdfPin = $script:APA7_DEPS['pymupdf']
$r = Invoke-Native -FilePath $py -Arguments @('-c', 'import pymupdf; print(pymupdf.__version__)')
if ($r.ExitCode -ne 0 -or $r.First -ne $pymupdfPin) {
    if ((ShouldInstall 'pymupdf') -and (Confirm-Install "pip install pymupdf==$pymupdfPin into $venvDir")) {
        Step "Installing pymupdf==$pymupdfPin into the virtual environment..."
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $pipOut = & $py -m pip install --disable-pip-version-check "pymupdf==$pymupdfPin" 2>&1
        $pipCode = $LASTEXITCODE
        $ErrorActionPreference = $prevEap
        if ($pipCode -ne 0) {
            Write-Output "FAILED to install pymupdf==$pymupdfPin. pip output:"
            $pipOut | ForEach-Object { Write-Output ('    ' + [string]$_) }
            exit 1
        }
        Step 'pymupdf installed.'
    } else {
        Step 'pymupdf missing or wrong version and -Only excludes it: skipping.'
    }
} else {
    Step ('pymupdf present: ' + $r.First)
}

# --- 5. LibreOffice (single PDF engine) ------------------------------------
$soffice = Get-SofficePath
if (-not $soffice) {
    if (ShouldInstall 'libreoffice') {
        if (Confirm-Install 'Install LibreOffice') {
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
        }
    } else {
        Step 'LibreOffice is missing and -Only excludes it: skipping.'
    }
} else { Step ('LibreOffice present: ' + $soffice) }

# --- 6. Final check --------------------------------------------------------
Step '--- Final environment check ---'
& $checkScript
exit $LASTEXITCODE
