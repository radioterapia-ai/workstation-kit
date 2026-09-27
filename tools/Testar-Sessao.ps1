#requires -Version 5.1
<#
    Testar-Sessao.ps1 - as protecoes do Modulo 5, exercitadas de verdade.

    Herdado do WORKSTATION_RT no fork de 27/09/2026, e o que mais valeu a pena
    herdar: os seis casos abaixo nasceram de seis defeitos reais, e a natureza
    deles e de ARQUITETURA, nao de clinica. Valem igual aqui.

    POR QUE ESTE ARQUIVO EXISTE

    O Modulo 5 ganhou instantaneos para ficar rapido: em vez de consultar o
    Windows uma vez por processo, consulta uma vez e reaproveita. A conta de
    tempo fechou - o Aplicar caiu de 26,6 s para 2,6 s -, e junto vieram seis
    defeitos, todos da mesma familia: DADO CONGELADO USADO PARA DECIDIR.

    Nenhum deles aparecia em leitura de codigo nem em teste positivo. O que os
    expoe e o teste NEGATIVO: montar o caso em que a protecao PRECISA falhar se
    estiver quebrada, e exigir que ela nao falhe. Foi assim que se descobriu, em
    outra ocasiao, que Test-CaminhoEcossistema devolvia falso para todo caminho
    do Windows - a expressao regular compilava, o parse passava, e a protecao
    simplesmente nao acontecia.

    Cada bloco abaixo nomeia o defeito que ele pega. Se um teste passar a falhar,
    leia o nome antes de mexer no teste.

    COMO RODAR
        powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Sessao.ps1

    Nao abre janela, nao encerra processo de ninguem e nao escreve no projeto.
    Cria e apaga um em_uso.json de mentira em %LOCALAPPDATA%\RADIOTERAPIA_AI\,
    e sobe um cmd.exe proprio para ter uma arvore de processos real - os dois
    saem no fim, inclusive se um teste estourar.

    Sai com 0 se tudo passou, ou com o numero de falhas.
#>

$ErrorActionPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$raiz  = Split-Path -Parent $PSScriptRoot
$alvo  = Join-Path $raiz 'WorkstationKit.ps1'
if (-not (Test-Path -LiteralPath $alvo)) { Write-Host "Nao achei $alvo"; exit 1 }

# ---------------------------------------------------------------------
# 1. O PARSE, e a integridade basica do arquivo
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '=== 1. integridade do arquivo ===' -ForegroundColor Cyan
$err = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($alvo, [ref]$null, [ref]$err)
if ($err -and $err.Count) {
    Write-Host 'PARSE FALHOU:' -ForegroundColor Red
    $err | ForEach-Object { Write-Host ('  L{0}: {1}' -f $_.Extent.StartLineNumber, $_.Message) }
    exit 1
}
$fns = $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$dup = @($fns | Group-Object Name | Where-Object { $_.Count -gt 1 })
Write-Host ('  parse OK       funcoes: {0}    duplicadas: {1}' -f $fns.Count, $dup.Count)
if ($dup.Count) { $dup | ForEach-Object { Write-Host ('  DUPLICADA: ' + $_.Name) -ForegroundColor Red }; exit 1 }
$tres = [System.IO.File]::ReadAllBytes($alvo)[0..2] -join ','
if ($tres -ne '239,187,191') {
    # Regra de ambiente do ecossistema: .ps1 com acento e sem BOM e lido como
    # cp1252 pelo PowerShell 5.1, e um travessao dentro de string fecha a string.
    Write-Host '  BOM AUSENTE - .ps1 com acento sem BOM quebra no PowerShell 5.1' -ForegroundColor Red
    exit 1
}
Write-Host '  BOM presente'

# ---------------------------------------------------------------------
# 2. Carrega as listas e as funcoes REAIS do aplicativo
#    (nao copias: o teste tem de morrer junto com o codigo que testa)
# ---------------------------------------------------------------------
$listas = 'MarcaEcossistema','PastaClinica','Protegidos','SessaoAudio','SessaoEnfeites','SessaoRemoto',
          'SessaoPreservar','SessaoNaoCompactar','RaizesClinicas','Lim','ArquivoEmUso',
          'EmUsoMaxMin','SnapProc','SnapSvc','EmUsoOperacao','SnapProcEm','EmUsoOperacaoEm',
          'SnapMaxSeg','SnapSessao','SnapGrupos'
function Test-DentroDeFuncao {
    # Sobe a arvore: se algum ancestral e uma definicao de funcao, o no esta
    # dentro de um corpo e nao deve ser executado na carga.
    param($No)
    $p = $No.Parent
    while ($p) {
        if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $true }
        $p = $p.Parent
    }
    return $false
}

# CARREGA NA ORDEM DO ARQUIVO, nao na ordem dos nomes.
#
# As listas dependem umas das outras: $script:SessaoPreservar e
# $script:SessaoNaoCompactar terminam em '+ $script:SessaoRemoto'. Iterar nome
# por nome fazia a ordem de carga depender da ordem em que os nomes estao
# escritos aqui - que estava certa por sorte. Na ordem do arquivo, que e a
# ordem em que o aplicativo executa, a dependencia resolve por construcao.
$alvos = @{}
foreach ($nm in $listas) { $alvos[$nm.ToLower()] = $true }
$aCarregar = @($ast.FindAll({ param($x)
    $x -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $x.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
    $x.Left.VariablePath.UserPath -like 'script:*' -and
    -not (Test-DentroDeFuncao $x) }, $true) |
    Where-Object { $alvos.ContainsKey(($_.Left.VariablePath.UserPath -replace '(?i)^script:', '').ToLower()) } |
    Sort-Object { $_.Extent.StartLineNumber })
foreach ($a in $aCarregar) { Invoke-Expression $a.Extent.Text }

# GUARDA: alternativa vazia casa com TUDO.
#
# 'Spotify' -match 'a||b' devolve True. Numa lista de protecao isso vira
# 'protege tudo', e esta suite passaria inteira pelo motivo errado - porque
# nada poderia ser encerrado. E o mesmo modo de falha do [\/] que virou [/]:
# o regex compila, o parse passa, e a protecao para de discriminar. So que
# aqui ela para na direcao SEGURA, e por isso nada quebra e os testes ficam
# verdes. Uma suite que nao confere as proprias listas nao vale nada.
$listasRegex = 'Protegidos', 'SessaoPreservar', 'SessaoNaoCompactar', 'SessaoAudio', 'SessaoEnfeites', 'SessaoRemoto', 'RaizesClinicas'
$ruim = 0
foreach ($nm in $listasRegex) {
    $v = ''
    try { $v = [string](Get-Variable -Scope Script -Name $nm -ValueOnly -ErrorAction Stop) } catch { }
    if (-not $v) {
        Write-Host ('  LISTA VAZIA OU NULA: $script:' + $nm + ' - -match casaria com tudo') -ForegroundColor Red
        $ruim++
        continue
    }
    if ($v -match '\|\|' -or $v -match '^\|' -or $v -match '\|$') {
        Write-Host ('  ALTERNATIVA VAZIA em $script:' + $nm + ' - -match casaria com tudo') -ForegroundColor Red
        $ruim++
        continue
    }
    try { [void][regex]::new($v) } catch {
        Write-Host ('  REGEX NAO COMPILA: $script:' + $nm + ' - ' + $_.Exception.Message) -ForegroundColor Red
        $ruim++
        continue
    }
    # prova viva: um nome que nenhuma lista deveria casar
    if ('zzzQualquerCoisaQueNaoExiste999' -match $v) {
        Write-Host ('  $script:' + $nm + ' CASA COM QUALQUER COISA') -ForegroundColor Red
        $ruim++
    }
}
if ($ruim -gt 0) {
    Write-Host ''
    Write-Host ('  ' + $ruim + ' lista(s) carregada(s) errado. Os testes abaixo nao valeriam nada.') -ForegroundColor Red
    exit 1
}
Write-Host ('  {0} listas de protecao conferidas: nenhuma casa com qualquer coisa' -f $listasRegex.Count)
$funcoes = 'Get-SnapshotProc','Get-SnapshotSvc','Clear-SnapshotProc','Clear-SnapshotsOperacao',
           'Get-CaminhosProcesso','Get-CaminhoAoVivo','Get-EmUsoOperacao','Get-AppEmUso',
           'Test-CaminhoEcossistema','Test-PodeEncerrarSessao','Get-ProcessosSessao',
           'Test-DentroDaPasta','Get-VeredictoArquivo','Get-VeredictoPasta',
           'Test-Padrao'
foreach ($n in $funcoes) {
    $f = $ast.FindAll({ param($x)
        $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
    if ($f.Count -ne 1) { Write-Host ("  FALTOU $n (achei {0})" -f $f.Count) -ForegroundColor Red; exit 1 }
    Invoke-Expression $f[0].Extent.Text
}

# dubles de interface: o que so serve para desenhar na tela
$script:Avisos = @()
function Write-Log  { param($Texto, $Nivel = 'DADO') ; $script:Avisos += "$Nivel|$Texto" }
function Set-Status { param($t) }
function Pump {}
function Format-Bytes { param($Bytes) ; '{0:N0} MB' -f ($Bytes / 1MB) }

$script:Falhas = 0
function Vale {
    param($Rotulo, $Esperado, $Obtido)
    $ok = ($Esperado -eq $Obtido)
    $cor = if ($ok) { 'DarkGray' } else { 'Red' }
    Write-Host ('  [{0}] {1,-58} esperado={2} obtido={3}' -f $(if ($ok) { 'ok ' } else { 'FALHA' }), $Rotulo, $Esperado, $Obtido) -ForegroundColor $cor
    if (-not $ok) { $script:Falhas++ }
}

# em_uso.json de mentira, e uma arvore de processos de verdade.
#
# ISOLAMENTO. O teste NUNCA toca o em_uso.json de verdade: redireciona o caminho
# para o TEMP depois de ter carregado as funcoes. Elas leem $script:ArquivoEmUso
# em tempo de execucao, entao a redirecao vale para o teste inteiro e nao alcanca
# o aplicativo, que roda no proprio processo com o caminho proprio.
#
# A primeira versao desta suite RECUSAVA rodar quando o arquivo de verdade
# existia. A intencao estava certa - nao mexer no que e do Local Suite -, mas o
# efeito era uma suite que nao roda justamente quando ha o que testar. E foi o
# que aconteceu na pratica: o Local Suite deixou um arquivo para tras, com quatro
# PIDs ja mortos, e a suite desistiu.
$script:ArquivoEmUsoReal = $script:ArquivoEmUso
$script:ArquivoEmUso = Join-Path $env:TEMP ('kit_teste_em_uso_{0}.json' -f $PID)
$dir = Split-Path $script:ArquivoEmUso -Parent
$criouDir = $false
if (Test-Path -LiteralPath $script:ArquivoEmUsoReal) {
    Write-Host ''
    Write-Host '  Ha um em_uso.json de verdade nesta maquina, e nao vou toca-lo.' -ForegroundColor DarkGray
    Write-Host ('  O teste trabalha em ' + $script:ArquivoEmUso) -ForegroundColor DarkGray
}
function Grava-EmUso {
    param([int[]]$Pids)
    $j = '{"versao":1,"app":"teste-do-kit","atualizado":"' + (Get-Date).ToString('o') + '","pids":[' + ($Pids -join ',') + '],"porta":8777,"segmentando":true}'
    [System.IO.File]::WriteAllText($script:ArquivoEmUso, $j, (New-Object System.Text.UTF8Encoding($false)))
}

$pai = Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', 'ping -n 30 127.0.0.1 > nul' -PassThru -WindowStyle Hidden
Start-Sleep -Milliseconds 1200
$filhos = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$($pai.Id)" | ForEach-Object { [int]$_.ProcessId })

try {
    Write-Host ''
    Write-Host ('=== 2. caminho ao vivo vence mapa congelado ===' ) -ForegroundColor Cyan
    Write-Host '    (defeito: processo nascido depois do instantaneo ficava sem caminho,'
    Write-Host '     Test-CaminhoEcossistema respondia "nao sei" e a marca nao protegia)'
    $vivo = Get-Process -Id $PID
    Vale 'processo fora do mapa devolve o caminho ao vivo' $vivo.Path (Get-CaminhoAoVivo -Processo $vivo -Mapa @{})
    $semPath = [pscustomobject]@{ Id = 424242; Path = $null }
    Vale 'sem .Path, o mapa e o reforco' 'C:\RADIOTERAPIA_AI\bin\python.exe' (Get-CaminhoAoVivo -Processo $semPath -Mapa @{ 424242 = 'C:\RADIOTERAPIA_AI\bin\python.exe' })
    Vale 'sem .Path e sem mapa devolve vazio, nunca lanca' '' (Get-CaminhoAoVivo -Processo $semPath -Mapa @{})
    $nova = [pscustomobject]@{ Id = 777001; Path = 'C:\RADIOTERAPIA_AI\_montagem\radai_x\python.exe' }
    Vale 'COMO ERA: inferencia nova nao era reconhecida' $false (Test-CaminhoEcossistema ([string]@{}[777001]))
    Vale 'COMO FICA: inferencia nova e reconhecida pela marca' $true (Test-CaminhoEcossistema (Get-CaminhoAoVivo -Processo $nova -Mapa @{}))

    Write-Host ''
    Write-Host '=== 3. a sinalizacao do Local Suite so sabe ACRESCENTAR protecao ===' -ForegroundColor Cyan
    Write-Host '    (defeito: o filtro por "vivos" do instantaneo descartava PID que o'
    Write-Host '     arquivo listava explicitamente, se ele nasceu depois do retrato)'
    $fantasma = 999731
    Grava-EmUso @($PID, $fantasma)
    Clear-SnapshotsOperacao
    $r = Get-AppEmUso
    Vale 'arquivo fresco' $true $r.Fresco
    Vale 'PID vivo listado esta protegido' $true ($r.Protegidos -contains [int]$PID)
    Vale 'PID fora do instantaneo NAO e descartado' $true ($r.Protegidos -contains $fantasma)
    Vale 'e o executor nao o encerra' $false (Test-PodeEncerrarSessao -Nome 'python' -Caminho 'C:\Users\x\python.exe' -ProcId $fantasma -EmUso $r)

    Write-Host ''
    Write-Host '=== 4. uniao: o arquivo relido AGORA, mais a descendencia ===' -ForegroundColor Cyan
    Write-Host '    (defeito: o gate "Protegidos.Count -eq 0" trocava uma leitura pela'
    Write-Host '     outra so quando o arquivo estava vazio, isto e, quando nao havia'
    Write-Host '     nada a proteger. Com o arquivo cheio, a descendencia sumia)'
    Write-Host ('    pai cmd.exe PID {0}, filho(s): {1}' -f $pai.Id, $(if ($filhos.Count) { $filhos -join ',' } else { '(nenhum)' }))
    Grava-EmUso @($pai.Id)
    Clear-SnapshotsOperacao
    $u = Get-EmUsoOperacao -Renovar
    Vale 'pai protegido' $true ($u.Protegidos -contains [int]$pai.Id)
    if ($filhos.Count) {
        Vale 'FILHO protegido pela descendencia' $true ($u.Protegidos -contains [int]$filhos[0])
        Vale 'executor nao encerra o filho' $false (Test-PodeEncerrarSessao -Nome 'PING' -Caminho 'C:\Windows\System32\PING.EXE' -ProcId ([int]$filhos[0]) -EmUso $u)
    }
    # PID que entra no arquivo DEPOIS, sem renovar a descendencia
    $novo = 999741
    Grava-EmUso @($pai.Id, $novo)
    $u2 = Get-EmUsoOperacao
    Vale 'PID recem-listado entra (veio do -Rapido)' $true ($u2.Protegidos -contains $novo)
    if ($filhos.Count) {
        Vale 'e a descendencia nao se perdeu (veio do completo)' $true ($u2.Protegidos -contains [int]$filhos[0])
        $soRapido = Get-AppEmUso -Rapido
        Vale 'COMO ERA: -Rapido sozinho perdia o filho' $false ($soRapido.Protegidos -contains [int]$filhos[0])
        Vale 'COMO ERA: o gate nunca disparava com arquivo cheio' $false ($soRapido.Protegidos.Count -eq 0)
    }

    Write-Host ''
    Write-Host '=== 5. o instantaneo tem prazo, e fracasso nao se guarda ===' -ForegroundColor Cyan
    Write-Host '    (defeito: a guarda era "if ($script:SnapProc)", e um pscustomobject'
    Write-Host '     vazio e verdadeiro: consulta que falhasse ficava guardada para a'
    Write-Host '     operacao inteira, desligando em silencio a descendencia)'
    Clear-SnapshotsOperacao
    $null = Get-SnapshotProc
    Vale 'carimbo gravado depois de uma leitura boa' $true ($null -ne $script:SnapProcEm)
    $marca = $script:SnapProcEm
    $null = Get-SnapshotProc
    Vale 'dentro do prazo reaproveita' $true ($script:SnapProcEm -eq $marca)
    $script:SnapProcEm = (Get-Date).AddSeconds(-($script:SnapMaxSeg + 5))
    $null = Get-SnapshotProc
    Vale 'vencido o prazo, reconstroi' $true ($script:SnapProcEm -ne $marca)
    $script:SnapProc = [pscustomobject]@{ Caminhos = @{}; Pais = @{}; Vivos = @{} }
    $script:SnapProcEm = Get-Date
    Vale 'instantaneo vazio nao vira resposta' $true ((Get-SnapshotProc).Vivos.Count -gt 0)

    Write-Host ''
    Write-Host '=== 6. a lista de servicos nao pode emudecer ===' -ForegroundColor Cyan
    Write-Host '    (defeito: -ErrorAction Stop virava terminante no primeiro servico'
    Write-Host '     ilegivel, o catch engolia, a lista saia vazia e o aviso CRITICO de'
    Write-Host '     papel de servidor desaparecia sem dizer que a leitura falhou)'
    $script:SnapSvc = $null
    $script:Avisos = @()
    $svc = Get-SnapshotSvc
    Vale 'leu servicos' $true ($svc.Count -gt 0)
    Vale 'achou o Spooler pelo nome' $true (@($svc | Where-Object { $_.Name -eq 'Spooler' }).Count -ge 1)
    Vale 'tem Name e State' $true ([bool]($svc[0].Name -and $svc[0].State))
    Vale 'nenhum alerta quando deu certo' 0 (@($script:Avisos).Count)
    $script:SnapSvc = @()
    $script:Avisos = @()
    Vale 'lista vazia guardada e refeita' $true ((Get-SnapshotSvc).Count -gt 0)
    Vale 'e sem ALERTA falso' 0 (@($script:Avisos | Where-Object { $_ -like 'ALERTA*' }).Count)

    Write-Host ''
    if (-not $BS) { $BS = [string][char]92 }

    Write-Host '=== 7. o que nunca pode mudar ===' -ForegroundColor Cyan
    Grava-EmUso @($pai.Id)
    Clear-SnapshotsOperacao
    $g = Get-EmUsoOperacao -Renovar
    Vale 'carga do ecossistema pela marca: nao encerra' $false (Test-PodeEncerrarSessao -Nome 'python' -Caminho 'C:\RADIOTERAPIA_AI\bin\python.exe' -ProcId 12 -EmUso $g)
    Vale 'marca no meio do caminho tambem' $false (Test-PodeEncerrarSessao -Nome 'python' -Caminho 'D:\dados\RADIOTERAPIA_AI\x\python.exe' -ProcId 12 -EmUso $g)
    Vale 'nome que so PARECE a marca nao protege' $true (Test-PodeEncerrarSessao -Nome 'python' -Caminho 'C:\RADIOTERAPIA_AI_ANTIGO\python.exe' -ProcId 12 -EmUso $g)
    Vale 'Citrix nao e encerrado' $false (Test-PodeEncerrarSessao -Nome 'wfica32' -Caminho 'C:\Program Files\Citrix\wfica32.exe' -ProcId 13 -EmUso $g)
    Vale 'banco do Vitrea nao e encerrado' $false (Test-PodeEncerrarSessao -Nome 'sqlservr' -Caminho 'C:\SQL\sqlservr.exe' -ProcId 14 -EmUso $g)
    Vale 'acesso remoto do TI nao e encerrado' $false (Test-PodeEncerrarSessao -Nome 'TeamViewer' -Caminho 'C:\TV\TeamViewer.exe' -ProcId 15 -EmUso $g)
    Vale 'dispensavel continua encerravel' $true (Test-PodeEncerrarSessao -Nome 'Spotify' -Caminho 'C:\Users\x\Spotify.exe' -ProcId 16 -EmUso $g)

    # ---------------------------------------------------------------
    # VERSAO UNIVERSAL: o que a lista precisa cobrir, e o que ainda nao cobre.
    #
    # Estes casos entram ANTES das entradas correspondentes entrarem na lista de
    # protecao. E deliberado: escrever o caso primeiro e o que faz aparecer a
    # entrada que nao casa - nome curto demais, ancora faltando, grafia diferente.
    # Lista primeiro e teste depois so confirma o que ja se acreditava.
    # ---------------------------------------------------------------
    Write-Host ''
    Write-Host '=== 9. fabricantes de radioterapia, alem dos da clinica de origem ===' -ForegroundColor Cyan
    # Caminhos montados por concatenacao com [char]92. A versao anterior usava '@'
    # como marcador e trocava depois. O resultado tinha a forma de um endereco de
    # e-mail - fabricante, arroba, produto - e acendia o varredor de vazamento em
    # oito linhas. Marcador tem de ser algo que nao pareca outra coisa.
    $BS = [string][char]92
    $fabricantes = @(
        @{ n = 'ARIA';       d = 'Varian';    q = 'Varian ARIA (record and verify)' }
        @{ n = 'Eclipse';    d = 'Varian';    q = 'Varian Eclipse (planejamento)' }
        @{ n = 'MOSAIQ';     d = 'Elekta';    q = 'Elekta MOSAIQ' }
        @{ n = 'Monaco';     d = 'Elekta';    q = 'Elekta Monaco' }
        @{ n = 'RayStation'; d = 'RaySearch'; q = 'RaySearch RayStation' }
        @{ n = 'Velocity';   d = 'Varian';    q = 'Varian Velocity' }
        @{ n = 'MIM';        d = 'MIM';       q = 'MIM Software' }
        @{ n = 'Vitrea';     d = 'Vital';     q = 'Canon Vitrea' }
    )
    foreach ($f in $fabricantes) {
        $cam = 'C:' + $BS + $f.d + $BS + $f.n + '.exe'
        Vale ('nao encerra: ' + $f.q) $false (Test-PodeEncerrarSessao -Nome $f.n -Caminho $cam -ProcId 20 -EmUso $g)
    }

    Write-Host ''
    Write-Host '=== 10. navegadores, alem do mercado de origem ===' -ForegroundColor Cyan
    foreach ($nav in @('msedge', 'chrome', 'firefox', 'iexplore', 'opera', 'vivaldi', 'brave')) {
        $cam = 'C:' + $BS + 'nav' + $BS + $nav + '.exe'
        Vale ('nao encerra navegador: ' + $nav) $false (Test-PodeEncerrarSessao -Nome $nav -Caminho $cam -ProcId 30 -EmUso $g)
    }

    Write-Host ''
    Write-Host '=== 11. o que AINDA NAO esta coberto, registrado em vez de escondido ===' -ForegroundColor Cyan
    Write-Host '    Nao e falha da suite: e a lista de trabalho da versao universal.'
    Write-Host '    Cada linha abaixo vira um Vale acima quando a entrada entrar na lista.'
    $pendentes = @(
        @{ g = 'acesso remoto';  n = @('AnyDesk', 'RustDesk', 'ScreenConnect', 'Splashtop', 'Parsec', 'ToDesk') }
        @{ g = 'navegador';      n = @('browser', 'UCBrowser', 'QQBrowser', 'Whale', 'Maxthon', 'Sleipnir') }
        @{ g = 'planejamento';   n = @('Pinnacle', 'Oncentra', 'iPlan', 'Precision', 'RayCare', 'XiO') }
        @{ g = 'prontuario';     n = @('Epic', 'PowerChart', 'MEDITECH', 'Soarian', 'Sectra') }
        @{ g = 'escritorio';     n = @('soffice', 'wps', 'et', 'wpp', 'Hwp') }
    )
    $faltando = 0
    foreach ($p in $pendentes) {
        foreach ($n in $p.n) {
            $cam = 'C:' + $BS + 'app' + $BS + $n + '.exe'
            if (Test-PodeEncerrarSessao -Nome $n -Caminho $cam -ProcId 40 -EmUso $g) {
                $faltando++
                Write-Host ('     [FALTA] {0,-16} {1}' -f $n, $p.g) -ForegroundColor Yellow
            } else {
                Write-Host ('     [ja ok] {0,-16} {1}' -f $n, $p.g) -ForegroundColor DarkGray
            }
        }
    }
    Write-Host ('    {0} entrada(s) por cobrir. Nao conta como falha: conta como trabalho.' -f $faltando) -ForegroundColor Yellow

    Write-Host ''
    Write-Host '=== 12. configuracao vazia nao casa com tudo ===' -ForegroundColor Cyan
    Write-Host '    (defeito: -like (''*'' + $vazio + ''*'') vira -like ''**'' e casa com TODO'
    Write-Host '     caminho. Medido: todo arquivo grande saia MANTER e toda pasta saia'
    Write-Host '     NAO APAGAR, com o motivo "Esta na pasta clinica" - que e falso)'

    $guardado = $script:PastaClinica
    try {
        foreach ($vazio in @('', '   ', $null)) {
            $script:PastaClinica = $vazio
            $rot = if ($null -eq $vazio) { '$null' } elseif (-not $vazio) { 'vazia' } else { 'so espacos' }

            Vale ('Test-DentroDaPasta com pasta ' + $rot) $false (Test-DentroDaPasta -Caminho 'C:\Qualquer\Coisa.txt' -Pasta $vazio)

            $v = Get-VeredictoArquivo -Caminho 'C:\Users\x\Downloads\enorme.msi' -Ext '.msi' -Bytes 500MB -Idade 200
            Vale ('arquivo comum nao vira MANTER com pasta ' + $rot) $true ($v.V -ne 'MANTER')

            $vp = Get-VeredictoPasta -Caminho 'C:\Users\x\Music'
            Vale ('pasta comum nao vira NAO APAGAR com pasta ' + $rot) $true ($vp.V -ne 'NAO APAGAR')
        }

        # E com a pasta PREENCHIDA continua funcionando, inclusive recusando o
        # vizinho de nome parecido - prefixo de texto casaria, componente nao.
        $script:PastaClinica = 'C:\PASTA CLINICA'
        Vale 'dentro da pasta e reconhecido'        $true  (Test-DentroDaPasta -Caminho 'C:\PASTA CLINICA\sub\a.xlsb' -Pasta $script:PastaClinica)
        Vale 'a propria pasta e reconhecida'        $true  (Test-DentroDaPasta -Caminho 'C:\PASTA CLINICA' -Pasta $script:PastaClinica)
        Vale 'com barra final tambem'               $true  (Test-DentroDaPasta -Caminho 'C:\PASTA CLINICA\' -Pasta $script:PastaClinica)
        Vale 'vizinho de nome parecido NAO entra'   $false (Test-DentroDaPasta -Caminho 'C:\PASTA CLINICA_ANTIGA\a.txt' -Pasta $script:PastaClinica)
        Vale 'outra pasta nao entra'                $false (Test-DentroDaPasta -Caminho 'C:\Outra\a.txt' -Pasta $script:PastaClinica)
        Vale 'caixa diferente entra'                $true  (Test-DentroDaPasta -Caminho 'c:\pasta clinica\a.txt' -Pasta $script:PastaClinica)
    }
    finally { $script:PastaClinica = $guardado }

    Write-Host ''
    Write-Host '=== 8. o retrato da sessao serve para MOSTRAR, e nao mente ===' -ForegroundColor Cyan
    Clear-SnapshotsOperacao
    $lista = @(Get-ProcessosSessao)
    Vale 'listou processos da sessao' $true ($lista.Count -gt 0)
    $cacheado = @(Get-ProcessosSessao)
    Vale 'a segunda chamada devolve o mesmo retrato' $true ($cacheado.Count -eq $lista.Count)
    Clear-SnapshotsOperacao
    Vale 'e a limpeza por operacao o descarta' 0 (@($script:SnapSessao.Keys).Count)

    Write-Host ''
    Write-Host '=== 13. as decisoes valem em qualquer cultura ===' -ForegroundColor Cyan
    Write-Host '    (defeito: -match herda IgnoreCase da cultura da THREAD. Em turco e azeri'
    Write-Host "     o I maiusculo baixa para i SEM PONTO, e 'CITRIX' deixava de casar"
    Write-Host "     'Citrix'. Fail-open: a linha e 'if (-match protegidos) { continue }',"
    Write-Host '     entao match falho significa NAO pula, ENCERRA)'


    Write-Host ''
    Write-Host '=== 14. agente de backup: protecao desenhada, nao acidental ===' -ForegroundColor Cyan
    Write-Host "    (antes: 'AcronisAgent' -match 'Onis' devolvia True por coincidencia de"
    Write-Host '     tres letras com o visualizador DICOM Onis. Funcionava, e sumiria no'
    Write-Host "     dia em que alguem tirasse 'Onis' da lista)"
    foreach ($b in @('AcronisAgent', 'AcronisCyberProtect', 'mms', 'VeeamAgent',
                     'CommVaultCvd', 'cvd', 'TrueImageMonitor', 'MacriumService',
                     'ReflectMonitor', 'ShadowProtectSvc', 'Arcserve', 'Datto')) {
        Vale ('backup nao e encerrado: ' + $b) $false (Test-PodeEncerrarSessao -Nome $b -Caminho '' -ProcId 0 -EmUso $null)
    }
    # Os curtos tem de ser ANCORADOS: sem ancora, tres letras casam meio Windows.
    foreach ($n in @('mmc', 'msdtc', 'svchost', 'WmiPrvSE', 'cmd', 'conhost')) {
        Vale ('ancora do curto nao pega ' + $n) $false ((Test-Padrao $n '^mms$') -or (Test-Padrao $n '^cvd$'))
    }
    # E a cobertura nao pode mais DEPENDER do token acidental.
    Vale 'Acronis coberto por entrada propria, nao por Onis' $true (Test-Padrao 'AcronisAgent' 'Acronis')
    $culturaVelha = [System.Threading.Thread]::CurrentThread.CurrentCulture
    try {
        foreach ($cult in 'pt-BR', 'en-US', 'de-DE', 'tr-TR', 'az-Latn-AZ') {
            try { [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::new($cult) }
            catch { Write-Host ('  (cultura ' + $cult + ' indisponivel nesta maquina)') -ForegroundColor DarkGray; continue }

            # Caixa TROCADA de proposito: e a unica forma que expoe o I turco.
            # Nome igual a entrada da lista passa em qualquer cultura.
            Vale ($cult + ': CITRIX maiusculo nao e encerrado')  $false (Test-PodeEncerrarSessao -Nome 'CITRIX'   -Caminho '' -ProcId 0 -EmUso $null)
            Vale ($cult + ': WFICA32 maiusculo nao e encerrado') $false (Test-PodeEncerrarSessao -Nome 'WFICA32'  -Caminho '' -ProcId 0 -EmUso $null)
            Vale ($cult + ': MOSAIQ maiusculo nao e encerrado')  $false (Test-PodeEncerrarSessao -Nome 'MOSAIQ'   -Caminho '' -ProcId 0 -EmUso $null)
            Vale ($cult + ': vitrea minusculo nao e encerrado')  $false (Test-PodeEncerrarSessao -Nome 'vitrea'   -Caminho '' -ProcId 0 -EmUso $null)
            Vale ($cult + ': dispensavel continua encerravel')   $true  (Test-PodeEncerrarSessao -Nome 'Spotify'  -Caminho '' -ProcId 0 -EmUso $null)
            Vale ($cult + ': marca do ecossistema protege')      $false (Test-PodeEncerrarSessao -Nome 'python' -Caminho ('C:' + $BS + $script:MarcaEcossistema + $BS + 'p.exe') -ProcId 0 -EmUso $null)
        }
    }
    finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $culturaVelha }
}
finally {
    Stop-Process -Id $pai.Id -Force -ErrorAction SilentlyContinue
    foreach ($f in $filhos) { Stop-Process -Id $f -Force -ErrorAction SilentlyContinue }
    [System.IO.File]::Delete($script:ArquivoEmUso)
    if ($criouDir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ''
if ($script:Falhas -eq 0) {
    Write-Host 'TODOS OS TESTES PASSARAM' -ForegroundColor Green
} else {
    Write-Host ('{0} FALHA(S) - leia o nome do bloco antes de mexer no teste' -f $script:Falhas) -ForegroundColor Red
}
Write-Host ''
exit $script:Falhas
