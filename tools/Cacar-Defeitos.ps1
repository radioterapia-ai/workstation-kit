#requires -Version 5.1
<#
    Cacar-Defeitos.ps1 - static defect hunt over the application.

    Five hunts, each aimed at a failure this code base has actually suffered:

      1. FORMAT ARITY. '(T key) -f a, b' where the source string has fewer
         placeholders than there are arguments, or more. Too few makes -f throw;
         too many silently drops a value the message promised. With a hidden
         console and SilentlyContinue, the first ends the step with no message.

      2. CASE-COLLIDING VARIABLES. PowerShell does not distinguish $P from $p.
         A loop variable that collides with an outer one overwrites it, and the
         symptom appears far from the cause. This has bitten this project three
         times.

      3. $script: ASSIGNED AND NEVER READ. Either dead state, or a name typo on
         one side that quietly split one variable into two.

      4. A KEY WHOSE VALUE IS ANOTHER KEY'S NAME, and the nested 'T (T x)' call
         that goes with it. It works while the other tables are empty, because
         an empty entry falls back to Portuguese and the indirection resolves.
         Translated, the outer key stops existing and the screen shows the
         missing-key marker -- a defect that appears only AFTER the language
         check goes green.

      5. EMPTY CATCH. Counted, not failed: many are deliberate here, because a
         hidden-console app must not die on a non-terminating error. But a catch
         that swallows around a line that DECIDES something is a different thing,
         so the count is printed for a human to look at.

    Exit code is the number of real defects (hunts 1-4). Hunt 5 only reports.
#>

param([string]$Alvo = '')

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$meuDir = $PSScriptRoot
if (-not $meuDir) { $meuDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Alvo) { $Alvo = Join-Path (Split-Path -Parent $meuDir) 'WorkstationKit.ps1' }

$err = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Alvo, [ref]$null, [ref]$err)
if ($err -and $err.Count) { Write-Host 'PARSE FAILED'; exit 1 }
$texto = Get-Content -LiteralPath $Alvo -Raw -Encoding UTF8

$defeitos = 0
function Defeito { param($t) Write-Host ("  [DEFECT] " + $t) -ForegroundColor Red; $script:defeitos++ }
function Ok      { param($t) Write-Host ("  [  ok  ] " + $t) -ForegroundColor DarkGray }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 1. format arity: (T key) -f ... ===' -ForegroundColor Cyan

# A tabela pt e a fonte, e ela e carregada EXECUTANDO a atribuicao, nunca por
# regex de uma linha. Tres valores do aplicativo sao strings de multiplas linhas,
# e um carregador de uma-linha-por-entrada simplesmente nao os ve: as chaves
# somem da fonte, e toda conferencia de aridade que dependa delas passa por
# ausencia, nao por acerto. Foi assim que este cacador deixou 3 chaves de fora.
$fonte = @{}
$atrib = @($ast.FindAll({ param($x)
    $x -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true) |
    Where-Object { $_.Extent.Text -match "(?i)^\s*\`$script:Textos\['pt'\]\s*=" })
if ($atrib.Count -ne 1) {
    Write-Host ("  [DEFECT] the pt table was found {0} time(s), expected exactly 1" -f $atrib.Count) -ForegroundColor Red
    exit 1
}
$script:Textos = @{}
Invoke-Expression $atrib[0].Extent.Text
foreach ($k in $script:Textos['pt'].Keys) { $fonte[$k] = [string]$script:Textos['pt'][$k] }
Write-Host ("  keys in the pt table: {0}" -f $fonte.Count)

$bin = @($ast.FindAll({ param($x)
    $x -is [System.Management.Automation.Language.BinaryExpressionAst] -and
    $x.Operator -eq 'Format' }, $true))
Write-Host ("  -f expressions: {0}" -f $bin.Count)

$ruins = 0
foreach ($b in $bin) {
    $esq = $b.Left.Extent.Text
    $k = [regex]::Match($esq, "T '([^']+)'")
    if (-not $k.Success) { continue }
    $chave = $k.Groups[1].Value
    if (-not $fonte.ContainsKey($chave)) { continue }
    $ph = @([regex]::Matches($fonte[$chave], '\{(\d+)') | ForEach-Object { [int]$_.Groups[1].Value } | Sort-Object -Unique)
    # argumentos: um ArrayLiteral conta os elementos; qualquer outra coisa conta 1
    $dir = $b.Right
    $nArg = 1
    if ($dir -is [System.Management.Automation.Language.ArrayLiteralAst]) { $nArg = $dir.Elements.Count }
    elseif ($dir -is [System.Management.Automation.Language.ParenExpressionAst] -and
            $dir.Pipeline.PipelineElements.Count -eq 1 -and
            $dir.Pipeline.PipelineElements[0].Expression -is [System.Management.Automation.Language.ArrayLiteralAst]) {
        $nArg = $dir.Pipeline.PipelineElements[0].Expression.Elements.Count
    }
    $maior = if ($ph.Count) { ($ph | Measure-Object -Maximum).Maximum + 1 } else { 0 }
    if ($maior -gt $nArg) {
        Defeito ("L{0}  key {1}: uses {{{2}}} but only {3} argument(s) are passed -- -f would THROW" -f
                 $b.Extent.StartLineNumber, $chave, ($maior - 1), $nArg)
        $ruins++
    } elseif ($maior -lt $nArg -and $maior -gt 0) {
        Defeito ("L{0}  key {1}: {2} argument(s) passed, highest placeholder is {{{3}}} -- {4} value(s) silently dropped" -f
                 $b.Extent.StartLineNumber, $chave, $nArg, ($maior - 1), ($nArg - $maior))
        $ruins++
    } elseif ($maior -eq 0 -and $nArg -ge 1) {
        Defeito ("L{0}  key {1}: no placeholder in the source, but {2} argument(s) passed" -f
                 $b.Extent.StartLineNumber, $chave, $nArg)
        $ruins++
    }
}
if ($ruins -eq 0) { Ok ("all {0} -f expressions match their source string" -f $bin.Count) }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 2. case-colliding variables inside a function ===' -ForegroundColor Cyan
$col = 0
foreach ($f in $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    $nomes = @{}
    foreach ($v in $f.FindAll({ param($x) $x -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
        $n = $v.VariablePath.UserPath
        if ($n -match ':') { continue }
        $b = $n.ToLower()
        if (-not $nomes.ContainsKey($b)) { $nomes[$b] = @{} }
        $nomes[$b][$n] = $true
    }
    foreach ($b in $nomes.Keys) {
        if ($nomes[$b].Count -gt 1) {
            Defeito ("{0}: `${1} written {2} different ways -- PowerShell treats them as ONE variable" -f
                     $f.Name, $b, (@($nomes[$b].Keys) -join ', '))
            $col++
        }
    }
}
if ($col -eq 0) { Ok 'no variable written two ways inside the same function' }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 3. $script: assigned and never read ===' -ForegroundColor Cyan
$decl = @{}
foreach ($a in $ast.FindAll({ param($x)
    $x -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $x.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
    $x.Left.VariablePath.UserPath -like 'script:*' }, $true)) {
    $decl[($a.Left.VariablePath.UserPath -replace '(?i)^script:', '')] = $true
}
$orfas = 0
foreach ($n in @($decl.Keys)) {
    $c = ([regex]::Matches($texto, '\$script:' + [regex]::Escape($n) + '\b', 'IgnoreCase')).Count
    if ($c -le 1) { Defeito ("`$script:{0} appears only once -- assigned and never read" -f $n); $orfas++ }
}
if ($orfas -eq 0) { Ok ("all {0} script-scoped variables are read somewhere" -f $decl.Count) }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 4. a key whose VALUE is the name of another key ===' -ForegroundColor Cyan
# Um extrator que roda duas vezes pode tratar como texto uma string que ja era
# nome de chave, e reescrever a chamada para T (T 'x'). Isso funciona enquanto
# as outras tabelas estao vazias, porque o vazio cai de volta no portugues e a
# indirecao se resolve. Traduzida, a chave de fora deixa de existir e a tela
# passa a mostrar o marcador de chave ausente - um defeito que so aparece
# DEPOIS de traduzir, ou seja, depois de a conferencia de idioma ficar verde.
$ind = 0
foreach ($k in $fonte.Keys) {
    if ($fonte.ContainsKey($fonte[$k])) {
        Defeito ("key {0} = '{1}', which is itself a key -- indirection breaks once translated" -f $k, $fonte[$k])
        $ind++
    }
}
foreach ($m in [regex]::Matches($texto, '\bT\s*\(\s*T\s')) {
    $ln = ($texto.Substring(0, $m.Index) -split "`n").Count
    Defeito ("L{0}  nested T (T ...): the inner call returns TEXT, and text is not a key" -f $ln)
    $ind++
}
if ($ind -eq 0) { Ok 'no key resolves to another key, and no nested T call' }

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 5. empty catch blocks (reported, not failed) ===' -ForegroundColor Cyan
$vazios = @($ast.FindAll({ param($x)
    $x -is [System.Management.Automation.Language.CatchClauseAst] -and
    $x.Body.Statements.Count -eq 0 }, $true))
Write-Host ("  empty catch: {0}" -f $vazios.Count)
Write-Host '  Deliberate here: a hidden-console app must not die on a non-terminating'
Write-Host '  error. Worth a human look when one wraps a line that DECIDES something.'

# ---------------------------------------------------------------------
Write-Host ''
if ($defeitos -eq 0) { Write-Host 'NO DEFECTS FOUND' -ForegroundColor Green }
else { Write-Host ("{0} DEFECT(S)" -f $defeitos) -ForegroundColor Red }
Write-Host ''
exit $defeitos
