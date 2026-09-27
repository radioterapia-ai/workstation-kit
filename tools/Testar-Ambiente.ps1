#requires -Version 5.1
<#
    Testar-Ambiente.ps1 - the three Windows traps, checked mechanically.

    None of these is a matter of style. Each one has broken a shipped file:

      1. .cmd and .bat must be CRLF. cmd.exe reads a batch file BY BYTE OFFSET.
         With LF it loses its place and starts executing whatever follows as a
         prompt command. The symptom does not look like a line-ending problem;
         it looks like wrong code. It breaks `if (...)` blocks and `goto :label`
         first, so a linear script survives and the trap stays LATENT until
         somebody adds a block.

      2. No chcp inside .cmd. Changing the code page mid-batch, with a multibyte
         byte further down, misaligns that same offset reading.

      3. .ps1 with an accented character needs a BOM. PowerShell 5.1 reads a
         BOM-less .ps1 as cp1252; in a comment that is harmless mojibake, but
         INSIDE A STRING it is fatal - an em dash becomes three characters, one
         of which is a quote, and the string closes in the middle with a syntax
         error pointing at an innocent line.

    Plus: no absolute path hard-coded in the application.

    Exit code is the number of failures.
#>

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$meuDir = $PSScriptRoot
if (-not $meuDir) { $meuDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$raiz = Split-Path -Parent $meuDir

$falhas = 0
function Falha { param($t) Write-Host ("  [FAIL] " + $t) -ForegroundColor Red; $script:falhas++ }
function Ok    { param($t) Write-Host ("  [ ok ] " + $t) -ForegroundColor DarkGray }

$arquivos = @(Get-ChildItem -LiteralPath $raiz -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\\\.git\\' -and $_.FullName -notmatch '\\_montagem\\' })

Write-Host ''
Write-Host '=== 1. .cmd and .bat: CRLF, ASCII only, no chcp ===' -ForegroundColor Cyan
$lotes = @($arquivos | Where-Object { $_.Extension -match '^\.(cmd|bat)$' })
if ($lotes.Count -eq 0) { Write-Host '  (none in this repository)' -ForegroundColor DarkGray }
foreach ($f in $lotes) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $t = [System.Text.Encoding]::GetEncoding(1252).GetString($b)
    $rel = $f.FullName.Substring($raiz.Length).TrimStart([char]92)
    if ($t -match "(?<!`r)`n") { Falha "$rel has a bare LF - cmd.exe reads by byte offset and will lose its place" }
    else { Ok "$rel is CRLF" }
    if ($t -match '(?im)^\s*chcp\b') { Falha "$rel contains chcp - it misaligns the byte-offset reading" }
    else { Ok "$rel has no chcp" }
    $fora = @($b | Where-Object { $_ -gt 127 })
    if ($fora.Count) { Falha "$rel has $($fora.Count) byte(s) outside ASCII" }
    else { Ok "$rel is pure ASCII" }
}

Write-Host ''
Write-Host '=== 2. .ps1 with an accent needs a BOM ===' -ForegroundColor Cyan
foreach ($f in @($arquivos | Where-Object { $_.Extension -eq '.ps1' })) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $rel = $f.FullName.Substring($raiz.Length).TrimStart([char]92)
    $bom = ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
    $corpo = if ($bom) { $b[3..($b.Length - 1)] } else { $b }
    $naoAscii = @($corpo | Where-Object { $_ -gt 127 }).Count
    if ($naoAscii -gt 0 -and -not $bom) {
        Falha "$rel has $naoAscii non-ASCII byte(s) and NO BOM - PowerShell 5.1 will read it as cp1252"
    } elseif ($bom) { Ok "$rel has a BOM" }
    else { Ok "$rel is pure ASCII (a BOM is not required)" }
}

Write-Host ''
Write-Host '=== 3. no absolute path in the application ===' -ForegroundColor Cyan
$app = Join-Path $raiz 'WorkstationKit.ps1'
if (Test-Path -LiteralPath $app) {
    $texto = Get-Content -LiteralPath $app -Raw -Encoding UTF8
    # Uma letra de unidade seguida de barra invertida, fora de comentario e fora
    # de padrao de regex. C:\Windows e C:\Program Files sao do proprio sistema e
    # nao sao "caminho de alguem"; o que nao pode e caminho de UMA instalacao.
    $suspeitos = @()
    $n = 0
    foreach ($l in ($texto -split "`r`n")) {
        $n++
        if ($l.TrimStart().StartsWith('#')) { continue }
        foreach ($m in [regex]::Matches($l, "'([A-Za-z]:\\[^']{3,})'")) {
            $v = $m.Groups[1].Value
            if ($v -match '(?i)^[A-Za-z]:\\(Windows|Program Files|ProgramData)\\') { continue }
            if ($v -match '\\\\') { continue }   # regex escapado, nao caminho
            $suspeitos += ("L{0}: {1}" -f $n, $v)
        }
    }
    if ($suspeitos.Count) {
        Falha ("{0} absolute path(s) in the application:" -f $suspeitos.Count)
        $suspeitos | Select-Object -First 8 | ForEach-Object { Write-Host ("        " + $_) -ForegroundColor Yellow }
    } else { Ok 'no absolute path outside the system folders' }
} else { Falha 'WorkstationKit.ps1 not found' }

Write-Host ''
Write-Host '=== 4. .gitattributes declares the line-ending rule ===' -ForegroundColor Cyan
$ga = Join-Path $raiz '.gitattributes'
if (-not (Test-Path -LiteralPath $ga)) { Falha '.gitattributes is missing - a clone would break the launcher' }
else {
    $g = Get-Content -LiteralPath $ga -Raw
    foreach ($r in '\*\.cmd\s+text\s+eol=crlf', '\*\.bat\s+text\s+eol=crlf') {
        if ($g -match $r) { Ok ("declared: " + ($r -replace '\\', '')) }
        else { Falha ("missing rule in .gitattributes: " + ($r -replace '\\', '')) }
    }
}

Write-Host ''
if ($falhas -eq 0) { Write-Host 'ENVIRONMENT OK' -ForegroundColor Green }
else { Write-Host ("{0} FAILURE(S)" -f $falhas) -ForegroundColor Red }
Write-Host ''
exit $falhas
