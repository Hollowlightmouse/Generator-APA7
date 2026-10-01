<#
    export-pdf.ps1 - Convierte el .docx en PDF con LibreOffice (motor unico).

    No contiene rutas absolutas: todo se descubre con lib\rutas.ps1 a partir de
    $PSScriptRoot.

    Puntos que importan y que ya se comprobaron en este equipo:
      * Se usa un perfil de usuario AISLADO de LibreOffice (-env:UserInstallation).
        Si se usa el perfil real, un LibreOffice ya abierto (por ejemplo desde
        la interfaz) hace que la conversion se quede colgada en silencio.
      * Se invoca soffice.com, no soffice.exe. El .exe es un lanzador que
        devuelve el control antes de terminar y su stderr rompe la redireccion
        en PowerShell; el .com es la consola y espera de verdad.
      * Si hay un proceso LibreOffice huérfano de una conversion anterior, se
        limpia antes de exportar. Sin esto, el segundo --convert-to falla.

    Parametros:
      -Docx    <ruta>   .docx de entrada   (obligatorio)
      -OutDir  <ruta>   carpeta de salida  (por defecto, la del .docx)
      -Log     <ruta>   fichero de log     (por defecto <OutDir>\_logs\03-export.log)
      -Timeout <seg>    tiempo maximo de espera (por defecto 300)

    Devuelve 0 si el PDF se genero, 1 si no.
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
        Write-Log "No existe el .docx: $Docx" 'FALLA'
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

    Write-Log '=== FASE 3: exportacion a PDF con LibreOffice ==='
    Write-Log "Raiz de la skill : $raiz"
    Write-Log "Workdir          : $work"
    Write-Log "Entrada (.docx)  : $docxAbs"
    Write-Log "Salida esperada  : $expected"

    $consola = Get-SofficeConsolePath
    if (-not $consola) {
        Write-Log 'LibreOffice no esta instalado. Ejecute scripts\instalar-entorno.ps1' 'FALLA'
        exit 1
    }
    Write-Log "LibreOffice      : $consola (consola; .exe colgaria el script)"

    # Limpiar procesos huerfanos de conversiones anteriores: si queda uno vivo,
    # el --convert-to siguiente se queda colgado sin decir por que.
    $huerfanos = Get-Process -Name 'soffice', 'soffice.bin' -ErrorAction SilentlyContinue
    if ($huerfanos) {
        Write-Log ("Limpando {0} proceso(s) LibreOffice previo(s)" -f $huerfanos.Count)
        Stop-ProcessTree -Ids $huerfanos.Id -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }

    if (Test-Path -LiteralPath $expected) { Remove-Item -LiteralPath $expected -Force }

    # Invoke-Soffice ya anade --headless --norestore --nolockcheck y un perfil
    # de usuario AISLADO. Aqui solo se pasan los argumentos de la conversion.
    $loArgs = @('--convert-to', 'pdf:writer_pdf_Export', '--outdir', $OutDir, $docxAbs)

    Write-Log ("Ejecutando: soffice.com {0}" -f ($loArgs -join ' '))
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Invoke-Soffice -Arguments $loArgs -TimeoutSeconds $Timeout -LogDir $logDir
    $sw.Stop()

    Write-Log ("soffice terminó con codigo {0} en {1:N1} s" -f $r.ExitCode, $sw.Elapsed.TotalSeconds)
    if ($r.StdOut) { Write-Log ("stdout: " + ($r.StdOut -replace '\s+', ' ').Trim()) }
    if ($r.StdErr) {
        $reales = @(Test-SofficeStderrIsBenign -Stderr $r.StdErr)
        if ($reales.Count -eq 0) {
            Write-Log 'stderr: solo ruido benigno de LibreOffice (ignorado)'
        } else {
            Write-Log ("stderr con {0} linea(s) real(es):" -f $reales.Count) 'AVISO'
            foreach ($l in $reales) { Write-Log ("  " + $l) 'AVISO' }
        }
    }
    if ($r.LogDir) { Write-Log ("Traza de LibreOffice en: {0}" -f $r.LogDir) }
    if (-not (Test-Path -LiteralPath $expected)) {
        Write-Log "No se produjo el PDF esperado: $expected" 'FALLA'
        exit 1
    }

    $pdf = Get-Item -LiteralPath $expected
    Write-Log ("PDF generado: {0} ({1:N0} bytes)" -f $pdf.FullName, $pdf.Length)
    Write-Log 'RESULTADO: OK'
    $code = 0
}
catch {
    Write-Log ("Error inesperado: " + $_.Exception.Message) 'FALLA'
    $code = 1
}
finally {
    Get-Process -Name 'soffice', 'soffice.bin' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue

    # El perfil aislado de LibreOffice solo sirve durante la conversion: son
    # miles de archivos y, si se deja, se acumulan miles de directorios por
    # corrida. Se borra siempre (tambien en error) y se conservan los .log,
    # que si son la diagnostica util.
    if ($r -and $r.ProfilePath -and (Test-Path -LiteralPath $r.ProfilePath)) {
        try {
            Remove-Item -LiteralPath $r.ProfilePath -Recurse -Force -ErrorAction Stop
            Write-Log ("Perfil temporal de LibreOffice eliminado: {0}" -f $r.ProfilePath)
        } catch {
            Write-Log ("No se pudo borrar el perfil temporal {0}: {1}" -f $r.ProfilePath, $_.Exception.Message) 'AVISO'
        }
    }

    if ($Log) {
        Set-Content -LiteralPath $Log -Value ($logLines -join "`n") -Encoding UTF8
        Write-Host ("Log escrito: {0}" -f $Log)
    }
}

exit $code
