<#
    export-pdf.ps1 - Converts the .docx into PDF with LibreOffice (single engine).

    Contains no absolute paths: everything is discovered with lib\rutas.ps1 from
    $PSScriptRoot.

    Points that matter and were already verified on this machine:
      * An ISOLATED LibreOffice user profile is used (-env:UserInstallation). If
        the real profile is used, an already-open LibreOffice (e.g. from the GUI)
        makes the conversion hang silently.
      * soffice.com is invoked, not soffice.exe. The .exe is a launcher that
        returns control before finishing and its stderr breaks PowerShell's
        redirection; the .com is the console and truly waits.
      * If there is an orphan LibreOffice process from a previous conversion, it
        is cleaned up before exporting. Without this, the second --convert-to
        fails.

    Parameters:
      -Docx    <path>   input .docx        (mandatory)
      -OutDir  <path>   output folder      (default: the .docx folder)
      -Log     <path>   log file           (default <OutDir>\_logs\03-export.log)
      -Timeout <sec>    maximum wait time  (default 300)

    Returns 0 if the PDF was produced, 1 if not.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$OutDir,
    [string]$Log,
    [int]$Timeout = 300
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\rutas.ps1')

$logLines = New-Object System.Collections.Generic.List[string]
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO ')
    $line = '[{0}] {1} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Write-Host $line
    $logLines.Add($line) | Out-Null
}

try {
    $raiz = Get-SkillRoot
    $work = Get-WorkDir

    $docxAbs = (Resolve-Path -LiteralPath $Docx -ErrorAction SilentlyContinue)
    if (-not $docxAbs) {
        Write-Log "The .docx does not exist: $Docx" 'FAIL'
        exit 1
    }
    $docxAbs = $docxAbs.Path

    if (-not $OutDir) { $OutDir = Split-Path -Parent $docxAbs }
    if (-not (Test-Path -LiteralPath $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path

    $logDir = Join-Path $OutDir '_logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    if (-not $Log) { $Log = Join-Path $logDir '03-export.log' }

    $expected = Join-Path $OutDir ((Get-Item -LiteralPath $docxAbs).BaseName + '.pdf')

    Write-Log '=== PHASE 3: export to PDF with LibreOffice ==='
    Write-Log "Skill root      : $raiz"
    Write-Log "Workdir         : $work"
    Write-Log "Input (.docx)   : $docxAbs"
    Write-Log "Expected output : $expected"

    $consola = Get-SofficeConsolePath
    if (-not $consola) {
        Write-Log 'LibreOffice is not installed. Run scripts\instalar-entorno.ps1' 'FAIL'
        exit 1
    }
    Write-Log "LibreOffice     : $consola (console; .exe would hang the script)"

    # Clean up orphan processes from previous conversions: if one stays alive,
    # the next --convert-to hangs without saying why.
    $huerfanos = Get-Process -Name 'soffice', 'soffice.bin' -ErrorAction SilentlyContinue
    if ($huerfanos) {
        Write-Log ("Cleaning up {0} leftover LibreOffice process(es)" -f $huerfanos.Count)
        Stop-ProcessTree -Ids $huerfanos.Id
        Start-Sleep -Seconds 2
    }

    if (Test-Path -LiteralPath $expected) { Remove-Item -LiteralPath $expected -Force }

    # Invoke-Soffice already adds --headless --norestore --nolockcheck and an
    # ISOLATED user profile. Here only the conversion arguments are passed.
    $loArgs = @('--convert-to', 'pdf:writer_pdf_Export', '--outdir', $OutDir, $docxAbs)

    Write-Log ("Running: soffice.com {0}" -f ($loArgs -join ' '))
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Invoke-Soffice -Arguments $loArgs -TimeoutSeconds $Timeout -LogDir $logDir
    $sw.Stop()

    Write-Log ("soffice finished with code {0} in {1:N1} s" -f $r.ExitCode, $sw.Elapsed.TotalSeconds)
    if ($r.StdOut) { Write-Log ("stdout: " + ($r.StdOut -replace '\s+', ' ').Trim()) }
    if ($r.StdErr) {
        $reales = @(Test-SofficeStderrIsBenign -Stderr $r.StdErr)
        if ($reales.Count -eq 0) {
            Write-Log 'stderr: only benign LibreOffice noise (ignored)'
        } else {
            Write-Log ("stderr with {0} real line(s):" -f $reales.Count) 'WARN'
            foreach ($l in $reales) { Write-Log ("  " + $l) 'WARN' }
        }
    }
    if ($r.LogDir) { Write-Log ("LibreOffice trace in: {0}" -f $r.LogDir) }

    # A timeout is reported BEFORE the "was the PDF produced?" check, otherwise a
    # wedged LibreOffice is diagnosed as a missing output file, which is a
    # different problem with a different fix.
    if ($r.TimedOut) {
        Write-Log ("LibreOffice did not finish within {0} s and was killed (ExitCode 124)." -f $Timeout) 'FAIL'
        Write-Log 'The conversion was aborted, not failed silently. Raise -Timeout, or check' 'FAIL'
        Write-Log 'whether a previous soffice process is stuck, and retry.' 'FAIL'
        $sofficeVivos = Get-Process -Name 'soffice', 'soffice.bin' -ErrorAction SilentlyContinue
        if ($sofficeVivos) {
            Stop-ProcessTree -Ids $sofficeVivos.Id
            Write-Log ("Killed {0} leftover LibreOffice process(es) from the aborted run." -f $sofficeVivos.Count) 'WARN'
        }
        exit 1
    }

    if (-not (Test-Path -LiteralPath $expected)) {
        Write-Log "The expected PDF was not produced: $expected" 'FAIL'
        exit 1
    }

    $pdf = Get-Item -LiteralPath $expected
    Write-Log ("PDF generated: {0} ({1:N0} bytes)" -f $pdf.FullName, $pdf.Length)
    Write-Log 'RESULT: OK'
    $code = 0
}
catch {
    Write-Log ("Unexpected error: " + $_.Exception.Message) 'FAIL'
    $code = 1
}
finally {
    # The tree, not just the launcher: a soffice.bin child survives a flat kill
    # and then makes the NEXT --convert-to fail. Same reason as line 84 above.
    $restantes = Get-Process -Name 'soffice', 'soffice.bin' -ErrorAction SilentlyContinue
    if ($restantes) { Stop-ProcessTree -Ids $restantes.Id }

    # The isolated LibreOffice profile is only useful during the conversion: it
    # is thousands of files and, if left behind, thousands of directories pile up
    # per run. It is always deleted (also on error) and the .log files, which are
    # the useful diagnostics, are kept.
    if ($r -and $r.ProfilePath -and (Test-Path -LiteralPath $r.ProfilePath)) {
        try {
            Remove-Item -LiteralPath $r.ProfilePath -Recurse -Force -ErrorAction Stop
            Write-Log ("Temporary LibreOffice profile deleted: {0}" -f $r.ProfilePath)
        } catch {
            Write-Log ("Could not delete the temporary profile {0}: {1}" -f $r.ProfilePath, $_.Exception.Message) 'WARN'
        }
    }

    if ($Log) {
        Set-Content -LiteralPath $Log -Value ($logLines -join "`n") -Encoding UTF8
        Write-Host ("Log written: {0}" -f $Log)
    }
}

exit $code
