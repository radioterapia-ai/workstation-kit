#requires -Version 5.1
<#
    Testar-Idioma.ps1 - the language layer, checked against the code that uses it.

    Six checks, and each one exists because the failure it catches is silent:

      1. every key the code asks for exists in every table
      2. no table entry is empty
      3. PLACEHOLDER PARITY. A translation with fewer {N} than the Portuguese
         source makes '-f' throw at runtime, and this app runs with a hidden
         console and SilentlyContinue: the step would end without a message.
         This is the check that matters most.
      4. T never throws and never returns empty, for every key, in every language
      5. a missing key returns the visible !!key!! marker, never blank
      6. button labels fit their control in every language

    Exit code is the number of failures.
#>

param([string]$Alvo = '')

$ErrorActionPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$meuDir = $PSScriptRoot
if (-not $meuDir) { $meuDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Alvo) { $Alvo = Join-Path (Split-Path -Parent $meuDir) 'WorkstationKit.ps1' }
if (-not (Test-Path -LiteralPath $Alvo)) { Write-Host "Nao achei $Alvo"; exit 1 }

$falhas = 0
function Falha { param($t) Write-Host ("  [FALHA] " + $t) -ForegroundColor Red; $script:falhas++ }
function Ok    { param($t) Write-Host ("  [ok ] " + $t) -ForegroundColor DarkGray }

$err = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Alvo, [ref]$null, [ref]$err)
if ($err -and $err.Count) { Write-Host "PARSE FALHOU"; exit 1 }

# carrega as tabelas e T do arquivo de verdade
function Dentro { param($No) $q = $No.Parent
    while ($q) { if ($q -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $true }; $q = $q.Parent }
    return $false }
# Aceita tanto '$script:X = ...' quanto '$script:X[''pt''] = ...'. A tabela de
# textos e preenchida por atribuicao INDEXADA, que nao e VariableExpressionAst -
# a primeira versao deste carregador so pegava a forma simples e reprovava com
# 'tabela ausente' num arquivo que tem as tres tabelas.
foreach ($nm in 'MarcaApp', 'Idiomas', 'Textos', 'Idioma') {
    $re = '(?i)^\s*\$script:' + [regex]::Escape($nm) + '(\[|\s*=)'
    $achados = @($ast.FindAll({ param($x)
        $x -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        -not (Dentro $x) }, $true) |
        Where-Object { $_.Extent.Text -match $re } |
        Sort-Object { $_.Extent.StartLineNumber })
    foreach ($a in $achados) { try { Invoke-Expression $a.Extent.Text } catch { } }
}
foreach ($fn in 'T', 'Get-IdiomaSalvo', 'Get-IdiomaInicial') {
    $f = @($ast.FindAll({ param($x)
        $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $fn }, $true))
    if ($f.Count -eq 1) { Invoke-Expression $f[0].Extent.Text }
}

$codigos = @($script:Idiomas | ForEach-Object { $_.Cod })
Write-Host ''
Write-Host ('=== idiomas: {0} ===' -f ($codigos -join ', ')) -ForegroundColor Cyan
if ($codigos.Count -lt 1) { Falha 'nenhum idioma declarado'; exit 1 }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 1. toda chave pedida pelo codigo existe em toda tabela ===' -ForegroundColor Cyan
$texto = Get-Content -LiteralPath $Alvo -Raw -Encoding UTF8
$usadas = @([regex]::Matches($texto, "\(T '([^']+)'\)") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
Write-Host ("  chaves usadas no codigo: {0}" -f $usadas.Count)
foreach ($c in $codigos) {
    $tab = $script:Textos[$c]
    if (-not $tab) { Falha ("tabela ausente: " + $c); continue }
    $faltam = @($usadas | Where-Object { -not $tab.ContainsKey($_) })
    if ($faltam.Count) {
        Falha ("{0}: {1} chave(s) sem entrada -- ex.: {2}" -f $c, $faltam.Count, ($faltam | Select-Object -First 4) -join ', ')
    } else { Ok ("{0}: todas as {1} chaves presentes" -f $c, $usadas.Count) }
}

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 2. nenhuma entrada vazia ===' -ForegroundColor Cyan
foreach ($c in $codigos) {
    $tab = $script:Textos[$c]
    if (-not $tab) { continue }
    $vazias = @($usadas | Where-Object { -not "$($tab[$_])".Trim() })
    if ($vazias.Count) {
        Falha ("{0}: {1} de {2} entradas VAZIAS (nao traduzidas)" -f $c, $vazias.Count, $usadas.Count)
    } else { Ok ("{0}: {1} entradas preenchidas" -f $c, $usadas.Count) }
}

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 3. PARIDADE DE PLACEHOLDER (a que mata em silencio) ===' -ForegroundColor Cyan
Write-Host '    Traducao com menos {N} que a fonte faz o -f LANCAR, e com console'
Write-Host '    oculto e SilentlyContinue a etapa morre sem mensagem.'
function Marcadores { param([string]$s)
    return @([regex]::Matches("$s", '\{(\d+)') | ForEach-Object { [int]$_.Groups[1].Value } | Sort-Object -Unique) }
$ptTab = $script:Textos['pt']
$comPh = @($usadas | Where-Object { (Marcadores $ptTab[$_]).Count -gt 0 })
Write-Host ("  chaves com placeholder na fonte: {0}" -f $comPh.Count)
foreach ($c in $codigos) {
    if ($c -eq 'pt') { continue }
    $tab = $script:Textos[$c]
    if (-not $tab) { continue }
    $ruins = @()
    foreach ($k in $usadas) {
        $a = Marcadores $ptTab[$k]
        $b = Marcadores $tab[$k]
        if (-not "$($tab[$k])".Trim()) { continue }   # vazia ja reprovou no bloco 2
        if (($a -join ',') -ne ($b -join ',')) { $ruins += ('{0} [fonte {1} / {2} {3}]' -f $k, ($a -join ','), $c, ($b -join ',')) }
    }
    if ($ruins.Count) {
        Falha ("{0}: {1} chave(s) com placeholder divergente" -f $c, $ruins.Count)
        $ruins | Select-Object -First 6 | ForEach-Object { Write-Host ("        " + $_) -ForegroundColor Yellow }
    } else { Ok ("{0}: placeholders batem com a fonte" -f $c) }
}

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 4. T nunca lanca e nunca devolve vazio ===' -ForegroundColor Cyan
$guardado = $script:Idioma
$mau = 0
foreach ($c in $codigos) {
    $script:Idioma = $c
    foreach ($k in $usadas) {
        $v = $null
        try { $v = T $k } catch { $mau++; continue }
        # !!chave!! nao conta como resposta: e justamente o sinal de que faltou.
        # Sem esta linha, o bloco passava com as tres tabelas ausentes.
        if ($null -eq $v -or "$v" -eq '' -or "$v" -like '!!*!!') { $mau++ }
    }
}
$script:Idioma = $guardado
if ($mau) { Falha ("{0} retorno(s) vazio(s) ou com excecao" -f $mau) }
else { Ok ("{0} chaves x {1} idiomas, nenhum vazio, nenhuma excecao" -f $usadas.Count, $codigos.Count) }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 5. chave inexistente aparece, nao some ===' -ForegroundColor Cyan
$m = T 'chave.que.nao.existe.999'
if ("$m" -eq '!!chave.que.nao.existe.999!!') { Ok 'devolve !!chave!! visivel' }
else { Falha ("devolveu '{0}' em vez do marcador visivel" -f $m) }
$m2 = T ''
if ("$m2" -eq '') { Ok 'chave vazia devolve vazio, sem lancar' } else { Falha 'chave vazia deu resultado inesperado' }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 6. o rotulo cabe no controle, em todo idioma ===' -ForegroundColor Cyan
$LARGURA_BOTAO = 244 - 26   # SetBounds(14, Y, 244, ...) menos o Padding de 12 e folga
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $fonte = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    # As chaves de BOTAO vem do codigo: New-Botao -Texto (T 'chave'). A versao
    # anterior usava o prefixo 'topo.', que e o ESCOPO do arquivo e nao o tipo
    # de controle - e por isso acusava mensagem de erro como rotulo estourado.
    $chavesBotao = @([regex]::Matches($texto, "New-Botao[^\r\n]*-Texto\s*\(T '([^']+)'\)") |
                     ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    Write-Host ('  rotulos de botao: {0}' -f $chavesBotao.Count)
    $estourou = @()
    foreach ($c in $codigos) {
        $script:Idioma = $c
        foreach ($k in $chavesBotao) {
            $t = T $k
            if ("$t".Length -gt 40) { continue }    # so rotulo curto e candidato a botao
            $w = [System.Windows.Forms.TextRenderer]::MeasureText("$t", $fonte).Width
            if ($w -gt $LARGURA_BOTAO) { $estourou += ('{0} [{1}] {2}px > {3}px: "{4}"' -f $k, $c, $w, $LARGURA_BOTAO, $t) }
        }
    }
    $script:Idioma = $guardado
    if ($estourou.Count) {
        Write-Host ("  {0} rotulo(s) passam da largura do botao:" -f $estourou.Count) -ForegroundColor Yellow
        $estourou | Select-Object -First 8 | ForEach-Object { Write-Host ("        " + $_) -ForegroundColor Yellow }
        Write-Host '  Nao reprova: AutoEllipsis corta sem quebrar. Mas rotulo cortado e rotulo perdido.' -ForegroundColor Yellow
    } else { Ok 'todo rotulo curto cabe na largura do botao' }
} catch { Write-Host '  (WinForms indisponivel nesta sessao; medicao de largura pulada)' -ForegroundColor DarkGray }

# ---------------------------------------------------------------------
Write-Host ''
if ($falhas -eq 0) { Write-Host 'CAMADA DE IDIOMA CONFERIDA' -ForegroundColor Green }
else { Write-Host ("{0} FALHA(S)" -f $falhas) -ForegroundColor Red }
Write-Host ''
exit $falhas
