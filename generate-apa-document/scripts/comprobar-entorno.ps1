<#
    comprobar-entorno.ps1 - Preflight (STEP 0) of the generate-apa-document skill

    Checks that ALL required tools exist before touching a document. If anything
    is missing, it warns and stops the pipeline (a document is never transformed
    with an incomplete environment).

    PDF engine: LibreOffice. Microsoft Word is not required.

    Output: one line per tool, in a parseable format:
        OK|<tool>|<detail>
        MISSING|<tool>|<detail>
    followed by:
        RESULT: OK | MISSING
    Exit code: 0 = full environment, 1 = something is missing.

    This script contains no absolute paths of any machine. Everything is
    resolved with lib\rutas.ps1 (environment -> PATH -> typical OS paths).

    Optional environment variables:
        APA7_SOFFICE, APA7_PYTHON, APA7_WORKDIR, APA7_NODEDIR
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot 'lib\rutas.ps1')

$results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param([string]$Tool, [bool]$Ok, [string]$Detail)
    $state = if ($Ok) { 'OK' } else { 'MISSING' }
    $results.Add([pscustomobject]@{ State = $state; Tool = $Tool; Detail = $Detail })
    Write-Output ("{0}|{1}|{2}" -f $state, $Tool, $Detail)
}

# --- PowerShell ------------------------------------------------------------
Add-Result 'PowerShell' ($PSVersionTable.PSVersion -ge [version]'5.1') ('v' + $PSVersionTable.PSVersion)

# --- Node.js ---------------------------------------------------------------
# Resolved with Get-Command. Reading $LASTEXITCODE when the command does NOT
# exist drags along the previous command's value and can give a false OK: that
# was a real bug in the previous version of this script.
$node = Get-NodeCmd
if ($node) {
    $r = Invoke-Native -FilePath $node -Arguments @('--version')
    Add-Result 'Node.js' ($r.ExitCode -eq 0) ("$($r.First)  ($node)")
} else {
    Add-Result 'Node.js' $false 'not found on PATH'
}

# --- npm -------------------------------------------------------------------
$npm = Get-NpmCmd
if ($npm) {
    $r = Invoke-Native -FilePath $npm -Arguments @('--version')
    Add-Result 'npm' ($r.ExitCode -eq 0) ("v$($r.First)  ($npm)")
} else {
    Add-Result 'npm' $false 'not found on PATH'
}

# --- docx library (npm) ----------------------------------------------------
# REAL check: Node is asked to resolve the module from the workdir. Verifying
# only that node_modules\docx\package.json exists gave false positives when the
# module was installed elsewhere, and that was exactly what broke the pipeline
# (the generator required it from a different relative path than the installer).
$nodeDir = Get-NodeDir
$modDir = Join-Path $nodeDir 'node_modules\docx'
if (-not (Test-Path -LiteralPath $modDir)) { $modDir = Join-Path $nodeDir 'node_modules/docx' }
$pkgJson = Join-Path $modDir 'package.json'

if (-not $node) {
    Add-Result 'docx (npm)' $false 'not verifiable: Node.js missing'
} elseif (-not (Test-Path -LiteralPath $pkgJson)) {
    Add-Result 'docx (npm)' $false "not installed in $nodeDir (run instalar-entorno.ps1)"
} else {
    $ver = ''
    try { $ver = (Get-Content -LiteralPath $pkgJson -Raw -Encoding UTF8 | ConvertFrom-Json).version } catch { }
    $posixMod = $modDir.Replace('\', '/')
    $probe = "require('$posixMod'); console.log('ok')"
    $r = Invoke-Native -FilePath $node -Arguments @('-e', $probe)
    $okDocx = ($r.ExitCode -eq 0 -and $r.First -eq 'ok')
    # A version that is not the pinned one is reported INSIDE the detail, and the
    # state stays OK: the library resolves and works, it is simply not the tested
    # one, and the library is not MISSING. The output contract documented in the
    # header (OK|/MISSING| + RESULT:) is intentionally left untouched.
    $esperadoDocx = $script:APA7_DEPS['docx']
    $notaDocx = ''
    if ($okDocx -and $ver -and $ver -ne $esperadoDocx) {
        $notaDocx = "  [MISMATCH: pinned $esperadoDocx]"
    }
    Add-Result 'docx (npm)' $okDocx ("v$ver$notaDocx  ($nodeDir)")
}

# --- Python + pymupdf ------------------------------------------------------
# The concrete venv is NOT fixed inside the skill: it is resolved by environment
# variable or PATH. See references\system-requirements.md.
$py = $env:APA7_PYTHON
if (-not $py) { $py = Get-PythonPath }

if (-not $py) {
    # "No usable Python" has two very different causes and they need different
    # fixes, so they are not reported the same way. A venv whose base interpreter
    # was removed or upgraded still EXISTS but cannot run: the executable is
    # there, so a bare "not found" sends the user to install Python, which does
    # nothing for them.
    $venvPath = Get-PythonVenvPath
    if ($venvPath) {
        Add-Result 'Python' $false "the skill's virtual environment exists but does not run: $venvPath (its base interpreter was moved, removed or upgraded). Recreate it with scripts\instalar-entorno.ps1"
    } else {
        Add-Result 'Python' $false 'not found (set APA7_PYTHON to an interpreter that has pymupdf)'
    }
    Add-Result 'pymupdf' $false 'not verifiable: Python missing'
} else {
    $r = Invoke-Native -FilePath $py -Arguments @('-c', 'import sys; print(sys.version.split()[0])')
    Add-Result 'Python' ($r.ExitCode -eq 0) ("$($r.First)  ($py)")

    $r2 = Invoke-Native -FilePath $py -Arguments @('-c', 'import pymupdf; print(pymupdf.__version__)')
    $okPyMu = ($r2.ExitCode -eq 0)
    $esperadoMu = $script:APA7_DEPS['pymupdf']
    $notaMu = ''
    if ($okPyMu -and $r2.First -and $r2.First -ne $esperadoMu) {
        $notaMu = "  [MISMATCH: pinned $esperadoMu]"
    }
    Add-Result 'pymupdf' $okPyMu ("$($r2.First)$notaMu  (interpreter: $py)")
}

# --- LibreOffice (single PDF engine) ---------------------------------------
# IMPORTANT: the version is queried with Invoke-Soffice, never with
# `& soffice.exe --version`. soffice.exe detaches, the child inherits the output
# pipe and PowerShell waits indefinitely: that hangs the whole script.
$soffice = Get-SofficePath
if ($soffice) {
    $r = Invoke-Soffice -Arguments @('--version') -TimeoutSeconds 60
    $version = ($r.StdOut -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
    $ok = ($r.ExitCode -eq 0) -and ([string]::IsNullOrWhiteSpace($version) -eq $false)
    if (-not $ok) { $version = 'could not read the version (ExitCode ' + $r.ExitCode + ')' }
    Add-Result 'LibreOffice' $ok ("$version  ($soffice)")
} else {
    $cands = (Get-SofficeCandidates | Where-Object { $_ }) -join ' | '
    if (-not $cands) { $cands = '(no candidates: set APA7_SOFFICE)' }
    Add-Result 'LibreOffice' $false "not found. Searched: $cands"
}

# --- Package manager (only needed for the self-install) --------------------
$pm = Get-PackageManagerCmd
if ($pm) {
    Add-Result 'Package manager' $true ("$($pm.Name)  ($($pm.Path))")
} else {
    Add-Result 'Package manager' $false 'no winget (Windows) or brew (macOS/Linux): automatic installation cannot run; install by hand according to references\system-requirements.md'
}

# --- Summary ---------------------------------------------------------------
$missing = @($results | Where-Object { $_.State -eq 'MISSING' })
Write-Output ''
if ($missing.Count -eq 0) {
    Write-Output 'RESULT: OK'
    Write-Output 'ENVIRONMENT OK: all required tools are present.'
    exit 0
} else {
    Write-Output ('MISSING: ' + (($missing | ForEach-Object { $_.Tool }) -join ', '))
    Write-Output 'RESULT: MISSING'
    Write-Output 'Run scripts\instalar-entorno.ps1 to install them automatically. If that is not possible, stop the pipeline and tell the user.'
    exit 1
}
