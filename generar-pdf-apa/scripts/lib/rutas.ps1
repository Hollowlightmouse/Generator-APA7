# lib/rutas.ps1 - Descubrimiento de rutas de la skill generar-pdf-apa
#
# REGLA DE ORO: este archivo NO contiene NINGUNA ruta absoluta de una maquina
# concreta. Todo se resuelve en tiempo de ejecucion a partir de:
#   1) Variables de entorno (permiten sobrescribir sin editar la skill)
#   2) La propia ubicacion de la skill ($PSScriptRoot)
#   3) El PATH del sistema
#   4) Rutas por defecto de cada sistema operativo (Windows / macOS / Linux)
#
# Variables de entorno soportadas (todas opcionales):
#   APA7_SOFFICE  ruta al ejecutable de LibreOffice
#   APA7_PYTHON   ruta al interprete Python que tiene pymupdf instalado
#   APA7_WORKDIR  directorio de trabajo (donde vive node_modules/docx)
#   APA7_NODEDIR  directorio que contiene node_modules/docx
#
# Uso:  . "<ruta>\scripts\lib\rutas.ps1"
#
# NOTA: este archivo NO activa Set-StrictMode a proposito. Al hacer dot-source
# contaminaria el ambito del script que lo carga y abortaria con errores como
# "no se puede recuperar $IsMacOS" en PowerShell 5.1, donde esa variable no
# existe. La deteccion de plataforma usa Test-Path y Get-Variable.

# ---------------------------------------------------------------------------
# Deteccion de plataforma (compatible con PowerShell 5.1 y 7+)
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
# Raiz de la skill: dos niveles arriba desde scripts/lib
# ---------------------------------------------------------------------------
$script:APA7_SKILL_ROOT = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

function Get-SkillRoot {
    <#  Raiz de la carpeta de la skill.  #>
    return $script:APA7_SKILL_ROOT
}

function Get-SkillScript {
    <#  Ruta absoluta de un script de scripts\.  #>
    param([Parameter(Mandatory = $true)][string]$Name)
    $p = Join-Path $script:APA7_SKILL_ROOT "scripts\$Name"
    if (-not (Test-Path -LiteralPath $p)) { throw "Script no encontrado: $p" }
    return $p
}

# ---------------------------------------------------------------------------
# Directorio de trabajo: dentro de la skill, para que require('docx') resuelva
# sin depender del directorio actual (bug corregido del pipeline anterior).
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
    <#  Directorio que DEBE contener node_modules\docx.  #>
    if ($env:APA7_NODEDIR) { return $env:APA7_NODEDIR }
    return (Get-WorkDir)
}

function Get-LogDir {
    <#  Carpeta de logs de una corrida. Se crea si no existe.  #>
    param([string]$ForRun)
    if (-not $ForRun) { $ForRun = [IO.Path]::GetTempPath() }
    if (-not (Test-Path -LiteralPath $ForRun)) {
        New-Item -ItemType Directory -Path $ForRun -Force | Out-Null
    }
    return $ForRun
}

# ---------------------------------------------------------------------------
# LibreOffice: entorno -> PATH -> rutas tipicas de cada SO
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
        Devuelve la ruta de LibreOffice encontrada, o $null.
        No lanza excepcion: el preflight decide si es un error fatal.
    #>
    foreach ($p in (Get-SofficeCandidates)) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return $null
}

function Get-SofficeConsolePath {
    <#
        En Windows devuelve el lanzador CONSOLA de LibreOffice (soffice.com)
        si existe. Es OBLIGATORIO usarlo en lugar de soffice.exe:

        soffice.exe se desprende, el proceso hijo hereda el handle del pipe de
        salida y PowerShell se queda esperando para siempre. Eso cuelga el
        script (y el terminal entero). Con soffice.com + Start-Process con
        redireccion a fichero, el proceso termina solo y devuelve el codigo.

        En macOS/Linux se devuelve la misma ruta que Get-SofficePath (el
        ejecutable unix si esta bien comportado con &).
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
        Ejecuta LibreOffice de forma CONFIABLE y acotada en tiempo.

        Devuelve @{ ExitCode; StdOut; StdErr; TimedOut }

        Por que NO se usa `& soffice.exe`:
          En Windows soffice.exe se desprende, el hijo (soffice.bin) hereda el
          handle del pipe y PowerShell se queda leyendo para siempre: cuelga el
          script entero. Se usa el lanzador de consola soffice.com.

        Por que la salida va a FICHEROS y no a pipes:
          Con RedirectStandardOutput de .NET el proceso de LibreOffice no
          termina (ExitCode 124 / timeout) porque los streams quedan abiertos
          por el proceso hijo. Redirigiendo a fichero con Start-Process y
          sondeando HasExited, el fichero esta volcado en cuanto el proceso
          termina y se leen los dos streams con normalidad.

        Nunca redirigir con 2>&1 | Out-File junto a $ErrorActionPreference='Stop':
        PowerShell lanza NativeCommandError y aborta.
    #>
    param(
        [string[]]$Arguments = @(),
        [int]$TimeoutSeconds = 300,
        [string]$LogDir
    )
    $bin = Get-SofficeConsolePath
    if (-not $bin) { throw 'LibreOffice no encontrado. Defina APA7_SOFFICE o ejecute instalar-entorno.ps1.' }

    $stamp = [Guid]::NewGuid().ToString('N').Substring(0, 8)
    # Si el llamador no da -LogDir creamos un directorio temporal propio. Hay que
    # borrarlo antes de salir: si no, cada comprobacion del preflight deja un
    # apa7-lo-<pid> huerfano en TEMP (se acumulan indefinidamente).
    $runDirTemporal = -not $LogDir
    $runDir = if ($LogDir) { $LogDir } else { Join-Path ([IO.Path]::GetTempPath()) ('apa7-lo-' + $PID) }
    if (-not (Test-Path -LiteralPath $runDir)) { New-Item -ItemType Directory -Path $runDir -Force | Out-Null }
    $errFile = Join-Path $runDir "lo-$stamp.err.log"
    $outFile = Join-Path $runDir "lo-$stamp.out.log"
    $profile = Join-Path $runDir "lo_profile_$stamp"

    $args = @('--headless', '--norestore', '--nolockcheck', '--nofirststartwizard',
        ('-env:UserInstallation=' + (ConvertTo-Apa7FileUri $profile))) + $Arguments

    # Captura con `&`: stdout por asignacion, stderr a fichero. NO usar
    # 2>&1 | Out-File (lanza NativeCommandError con ErrorActionPreference Stop)
    # ni Start-Process: si el proceso padre ya redirige su salida, la
    # redireccion anidada deja a LibreOffice sin terminar nunca.
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $out = & $bin @args 2> $errFile
    $code = $LASTEXITCODE
    $ErrorActionPreference = $prevEap

    $outText = (($out | ForEach-Object { [string]$_ }) -join "`n")
    Set-Content -LiteralPath $outFile -Value $outText -Encoding UTF8 -ErrorAction SilentlyContinue

    $resultado = @{
        ExitCode   = $code
        StdOut     = $outText
        StdErr     = (Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue)
        TimedOut   = $false
        LogDir     = $runDir
        StdOutPath = $outFile
        StdErrPath = $errFile
        ProfilePath = $profile
    }

    # Ya leimos todo lo que necesitabamos. Si el directorio es nuestro (temporal),
    # se borra aqui para no dejar basura en TEMP.
    if ($runDirTemporal) {
        Remove-Item -LiteralPath $runDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    return $resultado
}

function ConvertTo-Apa7FileUri {
    <#  Convierte una ruta local en URI file:/// para -env:UserInstallation.  #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path).Replace('\', '/')
    if (-not $full.StartsWith('/')) { $full = '/' + $full }
    return 'file://' + $full
}

function ConvertTo-Apa7Arg {
    <#  Entrecomilla un argumento para la linea de comandos de Windows.  #>
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }
    $escaped = $Value -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function Stop-ProcessTree {
    <#  Mata un proceso y sus hijos (soffice -> soffice.bin).  #>
    param([Parameter(Mandatory = $true)][int]$Id)
    if (Test-Apa7IsWindows) {
        $tk = Join-Path $env:SystemRoot 'System32\taskkill.exe'
        if (Test-Path -LiteralPath $tk) {
            try { & $tk /PID $Id /T /F | Out-Null; return } catch { }
        }
    }
    try { Stop-Process -Id $Id -Force -ErrorAction SilentlyContinue } catch { }
}

function Test-SofficeStderrIsBenign {
    <#
        Filtra el ruido conocido de LibreOffice headless y devuelve SOLO las
        lineas que merecen atención (array vacio = todo el stderr era ruido).

        Se descartan:
          * "Could not find platform independent libraries <prefix>": aparece
            cuando el proceso hereda PYTHONHOME/PYTHONPATH. No afecta a la
            conversion.
          * "Warning: failed to launch javaldx", avisos de libpng.
          * El envoltorio que PowerShell anade al redirigir stderr
            (CategoryInfo / FullyQualifiedErrorId / RemoteException), que no es
            salida del proceso sino de PowerShell.
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
        # envoltorio de PowerShell, no del proceso
        if ($l -match 'CategoryInfo|FullyQualifiedErrorId|RemoteException') { continue }
        if ($l -match '^\+|^soffice(\.exe|\.com)?\s*:') { continue }
        if ($l -match '\.ps1:\s*\d+\s+Car') { continue }   # "En <script>.ps1:194 Caracter 12"
        $real += $l
    }
    return $real
}

# ---------------------------------------------------------------------------
# Python con pymupdf: entorno -> candidatos explicitos -> PATH
# ---------------------------------------------------------------------------
function Get-PythonCandidates {
    $c = New-Object System.Collections.Generic.List[string]

    if ($env:APA7_PYTHON) { $c.Add($env:APA7_PYTHON) }

    # venv con nombre convention dentro o cerca de la skill
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

    return $c
}

function Get-PythonPath {
    <#
        Devuelve el primer interprete Python EXISTENTE de los candidatos.
        La comprobacion de pymupdf la hace el preflight con ese interprete.
        Devuelve $null si no hay ninguno.
    #>
    foreach ($p in (Get-PythonCandidates)) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Utilidades compartidas
# ---------------------------------------------------------------------------
function Get-NodeCmd {
    <#  Ruta al ejecutable de Node, o $null. Nunca usa $LASTEXITCODE.  #>
    foreach ($n in @('node', 'node.exe')) {
        $cmd = Get-Command $n -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Get-NpmCmd {
    # En Windows se prefiere npm.cmd: es un ejecutable nativo y fija
    # $LASTEXITCODE, cosa que npm.ps1 (wrapper de PowerShell) no hace.
    $order = if (Test-Apa7IsWindows) { @('npm.cmd', 'npm', 'npm.exe', 'npm.ps1') } else { @('npm', 'npm.cmd') }
    foreach ($n in $order) {
        $cmd = Get-Command $n -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Invoke-Native {
    <#
        Ejecuta un comando nativo y devuelve @{ Output; ExitCode }.
        NO usa Select-Object -First 1 sobre la salida: eso corta el pipeline
        antes de tiempo y deja el proceso nativo colgado o con codigo de
        salida distinto de 0, produciendo FALSOS FALTA.
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
    <#  winget en Windows, brew en macOS/Linux. $null si no hay ninguno.  #>
    $w = Get-WingetCmd
    if ($w) { return @{ Name = 'winget'; Path = $w } }
    $b = Get-BrewCmd
    if ($b) { return @{ Name = 'brew'; Path = $b } }
    return $null
}

function Write-Apa7Log {
    <#  Escribe una linea con timestamp.  #>
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message)
}
