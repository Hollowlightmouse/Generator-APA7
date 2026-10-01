# lib/rutas.ps1 - Path discovery for the generar-pdf-apa skill
#
# GOLDEN RULE: this file contains NO absolute path of any specific machine.
# Everything is resolved at runtime from:
#   1) Environment variables (let you override without editing the skill)
#   2) The skill's own location ($PSScriptRoot)
#   3) The system PATH
#   4) Default locations of each operating system (Windows / macOS / Linux)
#
# Supported environment variables (all optional):
#   APA7_SOFFICE  path to the LibreOffice executable
#   APA7_PYTHON   path to the Python interpreter that has pymupdf installed
#   APA7_WORKDIR  working directory (where node_modules/docx lives)
#   APA7_NODEDIR  directory that contains node_modules/docx
#
# Usage:  . "<path>\scripts\lib\rutas.ps1"
#
# NOTE: this file does NOT enable Set-StrictMode on purpose. Dot-sourcing it
# would pollute the scope of the loading script and abort with errors such as
# "cannot retrieve $IsMacOS" on PowerShell 5.1, where that variable does not
# exist. Platform detection uses Test-Path and Get-Variable.

# ---------------------------------------------------------------------------
# Platform detection (compatible with PowerShell 5.1 and 7+)
# ---------------------------------------------------------------------------
function Test-Apa7IsMac {
    $v = Get-Variable -Name 'IsMacOS' -ErrorAction SilentlyContinue
    if ($v -and $v.Value) { return $true }
    if (Test-Path -LiteralPath '/Applications') { return $true }
    if (Test-Path -LiteralPath '/System/Library/CoreServices') { return $true }
    return $false
}

function Test-Apa7IsWindows {
    $v = Get-Variable -Name 'IsWindows' -ErrorAction SilentlyContinue
    if ($v -and $null -ne $v.Value) { return [bool]$v.Value }
    return ($null -ne (Get-Variable -Name 'env:SystemRoot' -ErrorAction SilentlyContinue))
}

# ---------------------------------------------------------------------------
# Skill root: two levels up from scripts/lib
# ---------------------------------------------------------------------------
$script:APA7_SKILL_ROOT = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

function Get-SkillRoot {
    <#  Root folder of the skill.  #>
    return $script:APA7_SKILL_ROOT
}

function Get-SkillScript {
    <#  Absolute path of a script inside scripts\.  #>
    param([Parameter(Mandatory = $true)][string]$Name)
    $p = Join-Path $script:APA7_SKILL_ROOT "scripts\$Name"
    if (-not (Test-Path -LiteralPath $p)) { throw "Script not found: $p" }
    return $p
}

# ---------------------------------------------------------------------------
# Working directory: inside the skill, so that require('docx') resolves without
# depending on the current directory (bug fixed from the previous pipeline).
# ---------------------------------------------------------------------------
function Get-WorkDir {
    if ($env:APA7_WORKDIR) { return $env:APA7_WORKDIR }
    $wd = Join-Path $script:APA7_SKILL_ROOT '.work'
    if (-not (Test-Path -LiteralPath $wd)) {
        New-Item -ItemType Directory -Path $wd -Force | Out-Null
    }
    return $wd
}

function Get-NodeDir {
    <#  Directory that MUST contain node_modules\docx.  #>
    if ($env:APA7_NODEDIR) { return $env:APA7_NODEDIR }
    return (Get-WorkDir)
}

function Get-LogDir {
    <#  Log folder of a run. Created if missing.  #>
    param([string]$ForRun)
    if (-not $ForRun) { $ForRun = [IO.Path]::GetTempPath() }
    if (-not (Test-Path -LiteralPath $ForRun)) {
        New-Item -ItemType Directory -Path $ForRun -Force | Out-Null
    }
    return $ForRun
}

# ---------------------------------------------------------------------------
# LibreOffice: environment -> PATH -> typical paths per OS
# ---------------------------------------------------------------------------
function Get-SofficeCandidates {
    $c = New-Object System.Collections.Generic.List[string]

    if ($env:APA7_SOFFICE) { $c.Add($env:APA7_SOFFICE) }

    $cmd = Get-Command 'soffice' -ErrorAction SilentlyContinue
    if ($cmd) { $c.Add($cmd.Source) }
    $cmd2 = Get-Command 'soffice.com' -ErrorAction SilentlyContinue
    if ($cmd2) { $c.Add($cmd2.Source) }

    if (Test-Apa7IsMac) {
        # macOS
        $c.Add('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        $c.Add('/opt/homebrew/bin/soffice')
        $c.Add('/usr/local/bin/soffice')
    } else {
        $pf = ${env:ProgramFiles}
        $pf86 = ${env:ProgramFiles(x86)}
        $la = $env:LOCALAPPDATA
        if ($pf) { $c.Add((Join-Path $pf 'LibreOffice\program\soffice.exe')) }
        if ($pf86) { $c.Add((Join-Path $pf86 'LibreOffice\program\soffice.exe')) }
        if ($la) { $c.Add((Join-Path $la 'Programs\LibreOffice\program\soffice.exe')) }
    }
    return $c
}

function Get-SofficePath {
    <#
        Returns the path of the LibreOffice executable found, or $null.
        It does not throw: the preflight decides whether it is a fatal error.
    #>
    foreach ($p in (Get-SofficeCandidates)) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return $null
}

function Get-SofficeConsolePath {
    <#
        On Windows returns the CONSOLE launcher of LibreOffice (soffice.com)
        if it exists. It is MANDATORY to use it instead of soffice.exe:

        soffice.exe detaches, the child process inherits the output pipe handle
        and PowerShell waits forever. That hangs the script (and the whole
        terminal). With soffice.com + Start-Process redirecting to a file, the
        process finishes on its own and returns the code.

        On macOS/Linux it returns the same path as Get-SofficePath (the unix
        executable, assuming it behaves well with &).
    #>
    $exe = Get-SofficePath
    if (-not $exe) { return $null }
    if ($exe.ToLower().EndsWith('.exe')) {
        $com = [IO.Path]::ChangeExtension($exe, '.com')
        if (Test-Path -LiteralPath $com) { return $com }
    }
    return $exe
}

function Invoke-Soffice {
    <#
        Runs LibreOffice RELIABLY and bounded in time.

        Returns @{ ExitCode; StdOut; StdErr; TimedOut }

        Why NOT `& soffice.exe`:
          On Windows soffice.exe detaches, the child (soffice.bin) inherits the
          pipe handle and PowerShell keeps reading forever: it hangs the whole
          script. The console launcher soffice.com is used instead.

        Why the output goes to FILES and not to pipes:
          With .NET RedirectStandardOutput the LibreOffice process does not
          finish (ExitCode 124 / timeout) because the streams stay open in the
          child process. By redirecting to a file with Start-Process and polling
          HasExited, the file is flushed as soon as the process ends and both
          streams are read normally.

        Never redirect with 2>&1 | Out-File next to $ErrorActionPreference='Stop':
        PowerShell throws NativeCommandError and aborts.
    #>
    param(
        [string[]]$Arguments = @(),
        [int]$TimeoutSeconds = 300,
        [string]$LogDir
    )
    $bin = Get-SofficeConsolePath
    if (-not $bin) { throw 'LibreOffice not found. Set APA7_SOFFICE or run instalar-entorno.ps1.' }

    # If the caller does not pass -LogDir we create our own temporary directory.
    # It must be deleted before returning: otherwise every preflight check leaves
    # an orphan apa7-lo-<pid> folder in TEMP (they accumulate indefinitely).
    $runDirTemporal = -not $LogDir
    $runDir = if ($LogDir) { $LogDir } else { Join-Path ([IO.Path]::GetTempPath()) ('apa7-lo-' + $PID) }
    if (-not (Test-Path -LiteralPath $runDir)) { New-Item -ItemType Directory -Path $runDir -Force | Out-Null }

    # FIXED names: a run overwrites the previous one instead of leaving a new
    # pair of files behind (the old random stamp piled up one pair per call).
    # Any leftovers from a crashed run are removed first.
    $errFile = Join-Path $runDir 'lo.err.log'
    $outFile = Join-Path $runDir 'lo.out.log'
    $profile = Join-Path $runDir 'lo_profile'
    Remove-Item -LiteralPath $errFile, $outFile -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $profile) {
        Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction SilentlyContinue
    }

    $args = @('--headless', '--norestore', '--nolockcheck', '--nofirststartwizard',
        ('-env:UserInstallation=' + (ConvertTo-Apa7FileUri $profile))) + $Arguments

    # Capture with `&`: stdout by assignment, stderr to a file. Do NOT use
    # 2>&1 | Out-File (throws NativeCommandError with ErrorActionPreference Stop)
    # nor Start-Process: if the parent process already redirects its output, the
    # nested redirection keeps LibreOffice from ever finishing.
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $out = & $bin @args 2> $errFile
    $code = $LASTEXITCODE
    $ErrorActionPreference = $prevEap

    $outText = (($out | ForEach-Object { [string]$_ }) -join "`n")
    Set-Content -LiteralPath $outFile -Value $outText -Encoding UTF8 -ErrorAction SilentlyContinue

    $result = @{
        ExitCode   = $code
        StdOut     = $outText
        StdErr     = (Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue)
        TimedOut   = $false
        LogDir     = $runDir
        StdOutPath = $outFile
        StdErrPath = $errFile
        ProfilePath = $profile
    }

    # We already read everything we needed. If the directory is ours (temporary),
    # it is deleted here so we do not leave garbage in TEMP.
    if ($runDirTemporal) {
        Remove-Item -LiteralPath $runDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    return $result
}

function ConvertTo-Apa7FileUri {
    <#  Turns a local path into a file:/// URI for -env:UserInstallation.  #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path).Replace('\', '/')
    if (-not $full.StartsWith('/')) { $full = '/' + $full }
    return 'file://' + $full
}

function ConvertTo-Apa7Arg {
    <#  Quotes an argument for the Windows command line.  #>
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }
    $escaped = $Value -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function Stop-ProcessTree {
    <#
        Kills one or more processes and their children (soffice -> soffice.bin).
        Accepts an array ($huerfanos.Id returns one), which is how export-pdf.ps1
        calls it. The old single [int]$Id signature rejected the array and the
        call silently failed.
    #>
    param([Parameter(Mandatory = $true)][int[]]$Ids)
    foreach ($Id in $Ids) {
        if (-not $Id) { continue }
        if (Test-Apa7IsWindows) {
            $tk = Join-Path $env:SystemRoot 'System32\taskkill.exe'
            if (Test-Path -LiteralPath $tk) {
                try { & $tk /PID $Id /T /F | Out-Null; continue } catch { }
            }
        }
        try { Stop-Process -Id $Id -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Test-SofficeStderrIsBenign {
    <#
        Filters the known LibreOffice headless noise and returns ONLY the lines
        that deserve attention (empty array = all stderr was noise).

        Discarded:
          * "Could not find platform independent libraries <prefix>": appears
            when the process inherits PYTHONHOME/PYTHONPATH. It does not affect
            the conversion.
          * "Warning: failed to launch javaldx", libpng warnings.
          * The wrapper PowerShell adds when redirecting stderr
            (CategoryInfo / FullyQualifiedErrorId / RemoteException), which is
            not process output but PowerShell's.
    #>
    param([string]$Stderr)
    if (-not $Stderr) { return @() }
    $real = @()
    foreach ($line in ($Stderr -split "`r?`n")) {
        $l = $line.Trim()
        if (-not $l) { continue }
        if ($l -match 'Could not find platform independent libraries') { continue }
        if ($l -match 'Warning: failed to launch javaldx') { continue }
        if ($l -match 'libpng warning') { continue }
        # PowerShell wrapper, not the process
        if ($l -match 'CategoryInfo|FullyQualifiedErrorId|RemoteException') { continue }
        if ($l -match '^\+|^soffice(\.exe|\.com)?\s*:') { continue }
        if ($l -match '\.ps1:\s*\d+\s+Car') { continue }   # "At <script>.ps1:194 char: 12"
        $real += $l
    }
    return $real
}

# ---------------------------------------------------------------------------
# Python with pymupdf: environment -> explicit candidates -> PATH
# ---------------------------------------------------------------------------
function Get-PythonCandidates {
    $c = New-Object System.Collections.Generic.List[string]

    if ($env:APA7_PYTHON) { $c.Add($env:APA7_PYTHON) }

    # venv with conventional name inside or near the skill
    $root = $script:APA7_SKILL_ROOT
    if (Test-Apa7IsMac) {
        $pyNames = @('python3', 'python')
    } else {
        $pyNames = @('python.exe', 'python3.exe')
    }
    foreach ($sub in @('venv\Scripts', 'venv\bin', '.venv\Scripts', '.venv\bin')) {
        foreach ($nm in $pyNames) { $c.Add((Join-Path $root (Join-Path $sub $nm))) }
    }

    $cmd = Get-Command 'python3' -ErrorAction SilentlyContinue
    if ($cmd) { $c.Add($cmd.Source) }
    $cmd2 = Get-Command 'python' -ErrorAction SilentlyContinue
    if ($cmd2) { $c.Add($cmd2.Source) }

    if (-not (Test-Apa7IsMac)) {
        # Typical install locations when Python is not on PATH.
        $patterns = @()
        if ($env:LOCALAPPDATA) { $patterns += (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python*\python.exe') }
        if (${env:ProgramFiles}) { $patterns += (Join-Path ${env:ProgramFiles} 'Python*\python.exe') }
        foreach ($pat in $patterns) {
            foreach ($p in (Get-ChildItem -Path $pat -ErrorAction SilentlyContinue)) {
                $c.Add($p.FullName)
            }
        }
        # The py launcher is tried LAST: it still works when a Microsoft Store
        # stub shadows python.exe on PATH.
        $py = Get-Command 'py' -ErrorAction SilentlyContinue
        if ($py) { $c.Add($py.Source) }
    }

    return $c
}

function Test-Apa7PythonWorks {
    <#
        True only if the path is a WORKING Python interpreter (it prints its
        version). The Microsoft Store stub lives under \WindowsApps\ and exists
        but is not a real interpreter, so it is rejected by path and, as a
        belt-and-suspenders, by actually running it.
    #>
    param([Parameter(Mandatory = $true)][string]$File)
    if (-not $File) { return $false }
    if ($File -match '\\WindowsApps\\') { return $false }
    if (-not (Test-Path -LiteralPath $File)) { return $false }
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $File '-c' 'import sys; print(sys.version.split()[0])' 2>$null
        $code = $LASTEXITCODE
    } catch {
        $code = 1; $out = @()
    } finally {
        $ErrorActionPreference = $prevEap
    }
    if ($code -ne 0) { return $false }
    return [bool]($out | Where-Object { $_ })
}

function Get-PythonPath {
    <#
        Returns the first candidate that actually RUNS. Interpreters that exist
        but are broken (Microsoft Store stubs, dangling venvs) are skipped.
        Whether pymupdf is installed is checked later by the preflight with the
        returned interpreter. Returns $null if none works.
    #>
    foreach ($p in (Get-PythonCandidates)) {
        if (Test-Apa7PythonWorks -File $p) { return $p }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Shared utilities
# ---------------------------------------------------------------------------
function Get-NodeCmd {
    <#  Path to the Node executable, or $null. Never uses $LASTEXITCODE.  #>
    foreach ($n in @('node', 'node.exe')) {
        $cmd = Get-Command $n -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Get-NpmCmd {
    # On Windows npm.cmd is preferred: it is a native executable and sets
    # $LASTEXITCODE, which npm.ps1 (a PowerShell wrapper) does not.
    $order = if (Test-Apa7IsWindows) { @('npm.cmd', 'npm', 'npm.exe', 'npm.ps1') } else { @('npm', 'npm.cmd') }
    foreach ($n in $order) {
        $cmd = Get-Command $n -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Invoke-Native {
    <#
        Runs a native command and returns @{ Output; ExitCode }.
        It does NOT use Select-Object -First 1 on the output: that cuts the
        pipeline early and leaves the native process hung or with a non-zero
        exit code, producing FALSE missing-tool results.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @()
    )
    $global:LASTEXITCODE = 0
    $out = & $FilePath @Arguments 2>&1
    $code = $LASTEXITCODE
    $lines = @($out | ForEach-Object { if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.ToString() } else { [string]$_ } })
    $first = if ($lines.Count -gt 0) { $lines[0] } else { '' }
    return @{ Output = $lines; First = $first; ExitCode = $code }
}

function Get-WingetCmd {
    $cmd = Get-Command 'winget' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Get-BrewCmd {
    $cmd = Get-Command 'brew' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Get-PackageManagerCmd {
    <#  winget on Windows, brew on macOS/Linux. $null if none.  #>
    $w = Get-WingetCmd
    if ($w) { return @{ Name = 'winget'; Path = $w } }
    $b = Get-BrewCmd
    if ($b) { return @{ Name = 'brew'; Path = $b } }
    return $null
}

function Write-Apa7Log {
    <#  Writes a timestamped line.  #>
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message)
}
