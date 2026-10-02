# lib/rutas.ps1 - Path discovery for the generate-apa-document skill
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
# Pinned versions of the libraries this skill's own code talks to
# ---------------------------------------------------------------------------
# ONE place for both. They used to be literal strings inside instalar-entorno.ps1
# (the log line and the npm command, four lines apart) with nothing ever checking
# them, so the pin was documentation rather than a restriction. The preflight
# now compares what is installed against these values and says so in its detail.
#
# pymupdf is pinned because verificar-pdf.py calls its API directly. 1.28.2 is the
# current release; above 1.24.3 the top-level module is `pymupdf` and the `fitz`
# alias no longer exists.
$script:APA7_DEPS = @{
    docx    = '9.7.1'
    pymupdf = '1.28.2'
}

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
    # Plain .NET call, not New-Item: an empty workdir is scratch state and must
    # exist even when the caller runs with -WhatIf. An empty directory installs
    # nothing, so honouring -WhatIf here only breaks the probes that follow.
    [void][IO.Directory]::CreateDirectory($wd)
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
    # Same reason as Get-WorkDir: creating a log folder installs nothing, and a
    # suppressed -WhatIf would make the run fail while trying to write its log.
    [void][IO.Directory]::CreateDirectory($ForRun)
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
        Runs LibreOffice with a REAL time bound.

        Returns @{ ExitCode; StdOut; StdErr; TimedOut; LaunchError; LogDir;
                  StdOutPath; StdErrPath; ProfilePath }

        Why the CONSOLE launcher (soffice.com), never soffice.exe:
          soffice.exe detaches, the child (soffice.bin) inherits the output pipe
          handle and PowerShell keeps reading forever. See
          Get-SofficeConsolePath.

        Why Start-Process redirecting to FILES, and not pipes:
          The child never writes to a pipe, so it cannot block on a full pipe
          buffer, and the files are complete the moment the process ends. With
          .NET RedirectStandardOutput the streams stay open in the child and it
          never reports a result. This is also the approach that
          Get-SofficeConsolePath has always documented.

        Why NOT `& $bin @args`: the call operator blocks with no way to bound the
        wait. That was the actual defect being fixed here: -TimeoutSeconds was
        accepted and then ignored, so a wedged LibreOffice hung the preflight
        forever and TimedOut was hardcoded to $false.

        On timeout the whole TREE is killed, not just the launcher: the real work
        happens in the soffice.bin child, which would survive and keep holding
        the output files. ExitCode 124 is the conventional "timed out" code.

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
    # Internal scratch, NOT an installation: it is created and deleted with plain
    # .NET calls so that it is never suppressed by the caller's -WhatIf. If the
    # directory is missing, LibreOffice is handed a -env:UserInstallation pointing
    # at nothing, prints nothing on --version and the preflight wrongly reports
    # LibreOffice as MISSING (exit code 0 and an empty version).
    [void][IO.Directory]::CreateDirectory($runDir)

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

    # -ArgumentList does NOT quote its elements, it just joins them with spaces.
    # Any argument containing a space (the document path, the user profile URI)
    # would be split by the child, so each one is quoted here.
    $argQuoted = @($args | ForEach-Object { ConvertTo-Apa7Arg $_ })

    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $proc = $null
    $exited = $true
    $launchError = $null
    try {
        # -NoNewWindow plus explicit redirection to FILES means the child writes
        # to the files, not to this process's console: it does not inherit the
        # caller's redirection, which is what used to keep LibreOffice running.
        $proc = Start-Process -FilePath $bin -ArgumentList $argQuoted -NoNewWindow -PassThru `
            -RedirectStandardOutput $outFile -RedirectStandardError $errFile
        # Reading .Handle is what makes .ExitCode usable later. Without it,
        # Start-Process -PassThru hands back a Process whose exit code is $null
        # even after WaitForExit() reports success (verified against
        # soffice.com: without this, ExitCode comes back empty, not 0).
        $null = $proc.Handle
        $exited = $proc.WaitForExit([Math]::Max(1000, $TimeoutSeconds * 1000))
    } catch {
        $launchError = $_.Exception.Message
        $exited = $false
    } finally {
        $ErrorActionPreference = $prevEap
    }

    if (-not $exited -and $proc) {
        try { Stop-ProcessTree -Ids @($proc.Id) } catch { }
        try { $proc.WaitForExit(10000) | Out-Null } catch { }
    }

    if ($exited) {
        try { $code = $proc.ExitCode } catch { $code = $null }
        if ($null -eq $code) { $code = 0 }   # exited cleanly, code not reported
    } elseif ($launchError) {
        $code = 127
    } else {
        $code = 124
    }

    $outText = (Get-Content -LiteralPath $outFile -Raw -ErrorAction SilentlyContinue)
    $errText = (Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue)

    $result = @{
        ExitCode   = $code
        StdOut     = $outText
        StdErr     = $errText
        TimedOut   = (-not $exited)
        LaunchError = $launchError
        LogDir     = $runDir
        StdOutPath = $outFile
        StdErrPath = $errFile
        ProfilePath = $profile
    }

    # We already read everything we needed. If the directory is ours (temporary),
    # it is deleted here so we do not leave garbage in TEMP.
    if ($runDirTemporal) {
        # Symmetric with the creation above: our own scratch is always cleaned up,
        # including under -WhatIf, so a dry run does not leak apa7-lo-* folders.
        try { [IO.Directory]::Delete($runDir, $true) } catch { }
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
function Get-PythonVenvPath {
    <#
        Path of the interpreter inside the skill's own virtual environment, or
        $null when there is no venv at all. It deliberately does NOT check that
        the interpreter RUNS: the preflight has to tell "no venv" apart from
        "venv present but broken", because the fix is different in each case
        (install Python vs recreate the venv).
    #>
    $root = $script:APA7_SKILL_ROOT
    $subs = if (Test-Apa7IsMac) { @('venv/bin', '.venv/bin') } else { @('venv\Scripts', '.venv\Scripts') }
    $names = if (Test-Apa7IsMac) { @('python3', 'python') } else { @('python.exe', 'python3.exe') }
    foreach ($sub in $subs) {
        foreach ($nm in $names) {
            $p = Join-Path $root (Join-Path $sub $nm)
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    return $null
}

function Get-PythonCandidates {
    $c = New-Object System.Collections.Generic.List[string]

    # 1) explicit override, always wins
    if ($env:APA7_PYTHON) { $c.Add($env:APA7_PYTHON) }

    # 2) the skill's own venv, BEFORE any interpreter on PATH. The venv carries
    #    the pinned pymupdf, so it has to win over a system-wide one that may be
    #    a different version. Order in this list decides which one is used.
    $venv = Get-PythonVenvPath
    if ($venv) { $c.Add($venv) }
    # also probe the alternative conventional names so a venv created by hand
    # under a different name is still found
    $root = $script:APA7_SKILL_ROOT
    if (Test-Apa7IsMac) {
        $subs = @('venv/bin', 'venv/Scripts', '.venv/bin', '.venv/Scripts')
        $pyNames = @('python3', 'python', 'python.exe')
    } else {
        $subs = @('venv\Scripts', 'venv/bin', '.venv\Scripts', '.venv/bin')
        $pyNames = @('python.exe', 'python3.exe', 'python3', 'python')
    }
    foreach ($sub in $subs) {
        foreach ($nm in $pyNames) { $c.Add((Join-Path $root (Join-Path $sub $nm))) }
    }

    # 3) interpreters on PATH
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
