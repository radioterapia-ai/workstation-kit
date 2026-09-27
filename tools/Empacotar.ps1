#requires -Version 5.1
<#
    Empacotar.ps1 - builds the release artifact.

    THE DELIVERABLE IS TWO FILES

        Start_WorkstationKit.cmd   the launcher
        WorkstationKit.ps1         the application

    plus a README.txt written here. No docs/, no tools/ - that is workshop
    material, not a deliverable.

    WHERE IT GOES

        <Destino>\v<version>\WorkstationKit_v<version>.zip
        <Destino>\latest\WorkstationKit_LATEST.zip     byte-for-byte copy
        <Destino>\latest\SHA256.txt

    The LATEST copy exists so an external page can link one URL that never
    changes. The version is inside the package, in VERSION.txt, not only in the
    file name - a fixed name buys a stable link and gives up saying which build
    it is, so the package has to say it itself.

    THE VERSION IS NOT TYPED HERE

    It is read from $script:Versao in the application itself. A packager that
    accepts a typed number lets you ship a .zip whose name does not match what
    the app shows on screen, and then "which version is running on that machine"
    is a guess again.

    AND IT CHECKS WHAT IT DID

    After writing, the package is REOPENED: the file list is compared against
    the expected one, the .ps1 is extracted from inside the .zip and its parse is
    validated, its BOM checked, its SHA256 compared with the source, and latest\
    compared byte for byte. Any failure exits non-zero.
#>

[CmdletBinding()]
param(
    [string]$Destino = '',
    [switch]$Simular
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$meuDir = $PSScriptRoot
if (-not $meuDir) { $meuDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$raiz = Split-Path -Parent $meuDir
if (-not $Destino) { $Destino = Join-Path $env:SystemDrive 'AI_DEPLOY\WORKSTATION_KIT' }

$app = Join-Path $raiz 'WorkstationKit.ps1'
$cmd = Join-Path $raiz 'Start_WorkstationKit.cmd'
$problemas = 0
function Passo { param($t) Write-Host ''; Write-Host "== $t" -ForegroundColor Cyan }
function Ok    { param($t) Write-Host "   ok    $t" -ForegroundColor DarkGray }
function Ruim  { param($t) Write-Host "   FAIL  $t" -ForegroundColor Red; $script:problemas++ }

Passo 'source'
foreach ($f in @($app, $cmd)) { if (-not (Test-Path -LiteralPath $f)) { Ruim "missing: $f"; exit 1 } }
Ok 'both deliverable files are present'

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $meuDir 'Testar-Tudo.ps1') | Out-Null
if ($LASTEXITCODE -ne 0) { Ruim "the gate failed with $LASTEXITCODE - nothing is packaged"; exit $LASTEXITCODE }
Ok 'the full gate passed'

Passo 'version, read from the application'
$err = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($app, [ref]$null, [ref]$err)
if ($err -and $err.Count) { Ruim 'the application does not parse'; exit 1 }
$versao = $null
$ast.FindAll({ param($x)
    $x -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $x.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
    $x.Left.VariablePath.UserPath -match '(?i)^script:Versao$' }, $true) |
    ForEach-Object { $versao = ($_.Right.Extent.Text -replace "^['`"]|['`"]$", '') }
if (-not $versao) { Ruim 'could not read $script:Versao'; exit 1 }
Ok "version $versao"

$cab = (Get-Content -LiteralPath $app -TotalCount 12 -Encoding UTF8) -join "`n"
if ($cab -notmatch [regex]::Escape($versao)) { Ruim "the file header does not mention $versao" }
else { Ok 'the header agrees with the number' }

if ($problemas -gt 0) { Write-Host ''; Write-Host "$problemas problem(s). Not packaged." -ForegroundColor Red; exit $problemas }

Passo 'building'
$montagem = Join-Path $raiz '_montagem\entrega'
$pastaV = Join-Path $Destino "v$versao"
$pastaL = Join-Path $Destino 'latest'
$nome   = "WorkstationKit_v$versao.zip"
$zip    = Join-Path $pastaV $nome
$zipL   = Join-Path $pastaL 'WorkstationKit_LATEST.zip'
Write-Host "   from $raiz"
Write-Host "   to   $zip"
Write-Host "   and  $zipL"
if ($Simular) { Write-Host ''; Write-Host 'Dry run: nothing written.' -ForegroundColor Yellow; exit 0 }

if (Test-Path -LiteralPath $montagem) { Remove-Item -LiteralPath $montagem -Recurse -Force }
New-Item -ItemType Directory -Path $montagem -Force | Out-Null
Copy-Item -LiteralPath $app -Destination $montagem
Copy-Item -LiteralPath $cmd -Destination $montagem

$leiaMe = @"
================================================================
 WORKSTATION KIT                                  version $versao
 Diagnostics, cleanup and session optimisation for Windows.
================================================================

HOW TO USE

  1. Extract BOTH files into the same folder.
  2. Double-click Start_WorkstationKit.cmd.

  They must stay together: the launcher only opens the .ps1 that
  sits beside it.

WHEN TO RUN IT

  Right after you sign in, before opening the work of the day.

  Run in the middle of the day it can close work in progress. The
  app measures that and tells you when the moment is wrong; it does
  not stop you, it says so and lets you decide.

WHAT IT DOES

  Diagnostics and inventory    a portrait of the machine
  Large files and folders      read-only; lists and gives a verdict
  Cleanup                      temporary files, caches, old Downloads
  Optimise the session         closes what is not in use, adjusts
                               CPU priority, compacts memory
  Persistence                  what comes back at every sign-in

NOTHING IS PERMANENT

  Closed processes, stopped tasks, CPU priority and compacted
  memory come back at the next sign-in. Startup items and
  performance tweaks come back through the Undo button. Old
  Downloads go to the Recycle Bin.

  Only the temporary files and caches you selected are gone, and
  that is the cleanup.

  It uninstalls nothing. Where a program looks dispensable it
  closes the process instead.

IT NEEDS NO ADMINISTRATOR

  And because of that it cannot reach system services. What
  protects your database is the scope, not a list.

LANGUAGES

  English, Portuguese, Spanish. The three buttons at the bottom of
  the left panel switch at once.

IF ANTIVIRUS BLOCKS IT

  It is an unsigned PowerShell script that enumerates processes.
  The project does not obfuscate to avoid detection. The route is
  for IT to allow it by hash after reading the source:

  https://github.com/radioterapia-ai/workstation-kit

  Support tool. Not validated for clinical use. See NOTICE.
  Built on $(Get-Date -Format 'yyyy-MM-dd').
"@
[System.IO.File]::WriteAllText((Join-Path $montagem 'README.txt'),
    ($leiaMe -replace "`r`n", "`n" -replace "`n", "`r`n"),
    [System.Text.Encoding]::GetEncoding(1252))
[System.IO.File]::WriteAllText((Join-Path $montagem 'VERSION.txt'),
    "$versao`r`n", (New-Object System.Text.UTF8Encoding($false)))
Copy-Item -LiteralPath (Join-Path $raiz 'NOTICE')  -Destination $montagem
Copy-Item -LiteralPath (Join-Path $raiz 'LICENSE') -Destination $montagem
Ok 'README.txt, VERSION.txt, NOTICE and LICENSE added'

foreach ($d in @($pastaV, $pastaL)) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }
foreach ($z in @($zip, $zipL)) { if (Test-Path -LiteralPath $z) { Remove-Item -LiteralPath $z -Force } }
Compress-Archive -Path (Join-Path $montagem '*') -DestinationPath $zip -CompressionLevel Optimal
Copy-Item -LiteralPath $zip -Destination $zipL -Force
Ok ("$nome written ({0:N0} KB)" -f ((Get-Item $zip).Length / 1KB))

Passo 'checking what was written'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$esperado = @('WorkstationKit.ps1', 'Start_WorkstationKit.cmd', 'README.txt', 'VERSION.txt', 'NOTICE', 'LICENSE')
$extraido = $null
try {
    $arq = [System.IO.Compression.ZipFile]::OpenRead($zip)
    try {
        $dentro = @($arq.Entries | ForEach-Object { $_.FullName })
        $falta = @($esperado | Where-Object { $dentro -notcontains $_ })
        $sobra = @($dentro | Where-Object { $esperado -notcontains $_ })
        if ($falta.Count) { Ruim ('missing from the package: ' + ($falta -join ', ')) }
        if ($sobra.Count) { Ruim ('unexpected in the package: ' + ($sobra -join ', ')) }
        if (-not $falta.Count -and -not $sobra.Count) { Ok ("exactly the {0} expected files" -f $esperado.Count) }
        $e = $arq.Entries | Where-Object { $_.FullName -eq 'WorkstationKit.ps1' }
        if ($e) {
            $extraido = Join-Path $env:TEMP ('wk_check_{0}.ps1' -f $PID)
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, $extraido, $true)
        }
    } finally { $arq.Dispose() }
} catch { Ruim ('could not reopen the package: ' + $_.Exception.Message) }

if ($extraido -and (Test-Path -LiteralPath $extraido)) {
    $e2 = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($extraido, [ref]$null, [ref]$e2)
    if ($e2 -and $e2.Count) { Ruim 'the .ps1 INSIDE the package does not parse' } else { Ok 'the .ps1 inside the package parses' }
    if (([System.IO.File]::ReadAllBytes($extraido)[0..2] -join ',') -ne '239,187,191') { Ruim 'the .ps1 inside the package lost its BOM' }
    else { Ok 'the .ps1 inside the package kept its BOM' }
    $h1 = (Get-FileHash -LiteralPath $app -Algorithm SHA256).Hash
    $h2 = (Get-FileHash -LiteralPath $extraido -Algorithm SHA256).Hash
    if ($h1 -ne $h2) { Ruim 'the .ps1 inside the package differs from the source' }
    else { Ok "identical to the source (SHA256 $($h1.Substring(0,16))...)" }
    Remove-Item -LiteralPath $extraido -Force -ErrorAction SilentlyContinue
}

$hv = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
$hl = (Get-FileHash -LiteralPath $zipL -Algorithm SHA256).Hash
if ($hv -ne $hl) { Ruim 'latest\ is not byte-for-byte identical to the version' } else { Ok 'latest\ is the same file, byte for byte' }

$sha = "{0}  WorkstationKit_LATEST.zip`r`n{0}  {1}`r`n" -f $hv, $nome
[System.IO.File]::WriteAllText((Join-Path $pastaL 'SHA256.txt'), $sha, (New-Object System.Text.UTF8Encoding($false)))
[System.IO.File]::WriteAllText((Join-Path $pastaV 'SHA256.txt'), $sha, (New-Object System.Text.UTF8Encoding($false)))
Ok 'SHA256.txt written next to both copies'

Write-Host ''
if ($problemas -eq 0) {
    Write-Host "PACKAGE READY  ->  $zip" -ForegroundColor Green
    Write-Host "STABLE LINK    ->  $zipL" -ForegroundColor Green
} else {
    Write-Host "$problemas problem(s) in the check. The .zip exists but must not be published." -ForegroundColor Red
}
Write-Host ''
exit $problemas
