#requires -Version 5.1
<#
    Testar-Tudo.ps1 - the single gate. Local and CI run exactly this.

    One entry point on purpose: a gate that is spelled out twice drifts, and the
    half that drifts is always the one nobody looks at.

    Exit code is the number of failed checks. Any non-zero blocks the push.
#>

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$meuDir = $PSScriptRoot
if (-not $meuDir) { $meuDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$raiz = Split-Path -Parent $meuDir

$falhas = 0
$resumo = @()

function Etapa {
    param([string]$Nome, [string]$Script, [string]$Porque)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkGray
    Write-Host ("  $Nome") -ForegroundColor Cyan
    Write-Host ("  $Porque") -ForegroundColor DarkGray
    Write-Host ('=' * 78) -ForegroundColor DarkGray
    $p = Join-Path $meuDir $Script
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host ("  FALTA: $Script") -ForegroundColor Red
        $script:falhas++
        $script:resumo += [pscustomobject]@{ Etapa = $Nome; Codigo = 'ausente'; Estado = 'FALHA' }
        return
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $p
    $c = $LASTEXITCODE
    if ($null -eq $c) { $c = 0 }
    $script:resumo += [pscustomobject]@{ Etapa = $Nome; Codigo = $c; Estado = $(if ($c -eq 0) { 'ok' } else { 'FALHA' }) }
    if ($c -ne 0) { $script:falhas += $c }
}

Write-Host ''
Write-Host '  WORKSTATION KIT - full gate' -ForegroundColor Cyan
Write-Host ("  $raiz")
Write-Host ("  PowerShell $($PSVersionTable.PSVersion)")

Etapa 'Protections'   'Testar-Sessao.ps1'      'parse, orphan calls, the eight lists, and the negative cases'
Etapa 'Language'      'Testar-Idioma.ps1'      'every key present, no empty entry, placeholder parity'
Etapa 'Leak'          'Verificar-Vazamento.ps1' 'no third-party network identifier reaches a public repository'
Etapa 'Environment'   'Testar-Ambiente.ps1'    'BOM, line endings, no chcp in .cmd, no absolute path'
Etapa 'Defects'      'Cacar-Defeitos.ps1'     'format arity, colliding variables, dead script-scoped state'

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor DarkGray
$resumo | Format-Table -AutoSize
if ($falhas -eq 0) {
    Write-Host '  GATE PASSED' -ForegroundColor Green
} else {
    Write-Host ("  GATE FAILED - $falhas failure(s). Nothing should be pushed.") -ForegroundColor Red
}
Write-Host ''
exit $falhas
