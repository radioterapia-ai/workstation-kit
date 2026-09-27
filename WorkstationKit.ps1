#requires -Version 5.1
<#
=====================================================================
 WORKSTATION KIT - CLINICAL AND OFFICE WORKSTATIONS        version 1.0
---------------------------------------------------------------------
 A versao UNIVERSAL do Kit de Suporte. Serve qualquer computador - de
 clinica, de consultorio, de recepcao ou pessoal -, em qualquer pais,
 preservando software de radioterapia de qualquer fabricante.

 Nasceu por fork do WORKSTATION_RT, feito para um hospital do Brasil.
 O que era daquele hospital saiu; ver docs/PROVENIENCIA.md.

 A interface esta em transicao para ingles, portugues e espanhol. O
 codigo e a documentacao seguem em portugues, que e a regra do
 ecossistema - ver CLAUDE.md.
---------------------------------------------------------------------
 Modulo 1 - Preparar ambiente         (rotina embutida neste arquivo)
 Modulo 2 - Inventario e diagnostico  (retrato + o que da para resolver)
 Modulo 3 - Limpeza segura            (analisa, marca o seguro, aplica)
 Modulo 4 - Arquivos e pastas grandes (somente leitura, so o C:)
 Modulo 5 - Otimizar sessao atual     (reversivel no proximo logon)

 O diagnostico do Citrix e do Tasy roda dentro do Modulo 2.

 Tudo roda com o usuario comum. Nenhuma acao exige administrador.
 Nada e apagado fora das pastas de cache do proprio perfil.
=====================================================================
#>

# A janela roda com o console oculto. Erro nao-terminante escrito no fluxo de erro
# nessa condicao derruba o aplicativo com "O pipeline foi interrompido".
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 } catch { }

# =====================================================================
# 1. CONFIGURACAO  (edite somente esta secao)
# =====================================================================
$script:Versao       = '1.0'
$script:BaseDir      = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$script:VersaoPreparador = ''
# MODULO 1. A versao da clinica trazia um batch de 358 linhas embutido aqui, que
# montava a arvore de trabalho daquele hospital: portais em Excel, atalhos, fonte
# de codigo de barras, mapeamento de unidade e o passo 10 que copiava a pasta da
# rede para o C:. Nada disso e universal, e sair levou embora tambem o dominio, o
# fileserver e os dois IPs internos.
#
# O Modulo 1 continua existindo e ja degradava bem: sem rotina embutida, ele pede
# um .bat de preparacao. Na versao universal e isso que ele e - "rode o seu script
# de preparacao" -, e essa e uma funcao honesta para qualquer maquina.
$script:ArquivoBatExterno = ''

# PASTA DE TRABALHO CLINICA. Opcional, e vazia por padrao. Preenchida, o kit
# confere se os arquivos dela estao em dia contra uma copia de referencia.
$script:PastaClinica = ''
$script:PastaClinicaRede = ''

# Onde gravar a planilha do inventario. Local por padrao, e resolvido pela API do
# Windows em vez de por nome: o caminho do perfil nao e adivinhavel e USERPROFILE
# pode estar redirecionado. Pode receber um caminho de rede para juntar maquinas.
$script:PastaRelatorios = try { [Environment]::GetFolderPath('MyDocuments') } catch { $env:TEMP }

# Marca das cargas do ecossistema radioterapia.ai (AUTO_CONTORNO e afins).
# E NOME DE PASTA, nunca caminho absoluto: a raiz muda de maquina - C: nesta,
# D:\RADIOTERAPIA_AI\LOCAL_SUITE em outra estacao - e o _instalados.json
# pode estar desatualizado. A marca viaja no caminho do proprio processo em
# execucao, entao nao depende de raiz declarada nem de registro em dia.
#
# Por que por caminho e nao por nome: todo processo do AUTO_CONTORNO e python.exe.
# O projeto chama o TotalSegmentator como 'sys.executable -m ...' de proposito,
# para ser relocavel. Proteger o nome 'python' protegeria tambem o .venv de
# desenvolvimento, que e descartavel; o caminho separa os dois.
#
# NA VERSAO UNIVERSAL isto e um PARAMETRO. O conceito e geral e vale para qualquer
# maquina: "existe uma arvore de trabalho cuja carga nunca deve ser encerrada, e
# ela se reconhece pelo CAMINHO do processo, nao pelo nome". Um consultorio que
# roda o proprio lote de processamento tem a mesma necessidade.
#
# Vazio desliga a protecao por marca sem quebrar nada: Test-CaminhoEcossistema
# passa a responder 'nao sei' para todo caminho, que e o padrao seguro dela.
$script:MarcaEcossistema = 'RADIOTERAPIA_AI'

# SUFIXOS DE DNS para tentar quando um nome CURTO de servidor nao resolve. Vazio
# por padrao: o que o kit usa sozinho e o dominio real da maquina, descoberto em
# tempo de execucao por Get-SufixosDns.
#
# A versao da clinica tinha o dominio dela chumbado, em dois lugares. Em qualquer
# outra rede isso e uma consulta de DNS que sempre falha - e, pior, o nome de um
# servidor de terceiro sendo consultado por uma maquina que nao e dele.
$script:SufixosDnsExtra = @()

# Arvores de DADO CLINICO. Governa o veredicto dos Modulos 3 e 4: o que casar
# aqui nunca e sugerido para apagar, nem como arquivo nem como pasta.
#
# Estava duplicado em Get-VeredictoArquivo e Get-VeredictoPasta, e as duas
# listas tinham DIVERGIDO: a de pasta nao tinha \Patients\, \ARIA, MOSAIQ,
# VitreaData nem Monaco. Uma pasta de paciente do Vitrea caia no veredicto
# generico 'confira o conteudo antes de mover ou apagar'. Agora e uma lista so.
#
# Levantado com a sessao do sincronizador Vitrea em 15/09/2026. Se uma raiz
# nova de exame entrar em uso, ACRESCENTE AQUI antes de ela existir no disco.
$script:RaizesClinicas =
    'Vitrea|VitreaData|\\Patients\\|\\Patients$' +
    '|\\radioterapia\\TC_DATA' +
    '|Mirada|Medis|INVIA|Corridor4DM|NeuroQ|OleaSphere|TomTec' +
    '|\\ARIA|MOSAIQ|Monaco|Eclipse|RayStation|Velocity' +
    '|VspApp|VspMgmt' +
    '|Digitalcore|\\Onis'

# PORTAIS. Vazios de proposito: os enderecos sao de cada servico, e os da clinica
# de origem sairam junto com o resto da rede dela.
#
# CitrixStores vazio nao desliga o diagnostico de Citrix: ele passa a relatar o
# que encontra na maquina sem ter destino para comparar. TasyUrls vazio faz o kit
# procurar sozinho nos atalhos da area de trabalho, no menu Iniciar e nos
# favoritos dos navegadores - caminho que ja existia, e que e o certo aqui porque
# nao presume nome de prontuario nenhum.
#
# Tasy e o prontuario da Philips usado no Brasil. Numa versao universal ele e UM
# caso entre muitos: Epic, Cerner, MEDITECH, Soarian, Sectra. Ver PENDENCIAS.md.
$script:CitrixStores = @()
$script:TasyUrls = @()
$script:TasyDescobertas = $null

$script:Lim = @{
    DiscoCritPct    = 10     # % livre em C: abaixo disso = critico
    DiscoAlertaPct  = 15
    RamCritPct      = 90
    RamAlertaPct    = 80
    UptimeAlertaH   = 72     # horas sem reiniciar
    UptimeCritH     = 168
    CitrixMs        = 3000   # resposta aceitavel do portal Citrix
    ArquivoGrandeMB = 100
    # Processo de fundo que o kit nao reconhece e que ocupa mais que isso vem
    # DESMARCADO: acima deste tamanho e mais provavel ser trabalho em andamento
    # que lixo. Medido nesta maquina: o maior grupo dispensavel desconhecido tem
    # 18 MB e o maior conhecido 52 MB, entao sobra folga larga para o lixo real.
    SessaoDesconhecidoMB = 300
    # Item de cache mexido nos ultimos N minutos nao e apagado, mesmo em alvo sem
    # filtro de idade. Nao muda a politica de retencao - continua limpando o lixo
    # de hoje -, so recusa apagar o que esta sendo escrito neste instante.
    CacheFrescoMin = 15
}

# =====================================================================
# 2. ESTADO INTERNO
# =====================================================================
$script:Achados      = New-Object System.Collections.ArrayList
$script:AlvosAtuais  = @()
$script:PlanoAtual   = @()
$script:GrandesItens = @()
$script:ModoPlano    = 'LIMPEZA'
$script:UltimoGanhoMemoria = 0
$script:SessaoEncerrados   = @()
$script:UltimaPrioridade   = $null
$script:ExplorerAtualizado = $false
$script:LogonSegundos      = 0
$script:StatusSessao       = $null
$script:DiasCorte    = 7    # idade minima para TODA a rotina: 3, 7, 15, 30 ou 0 (tudo)
$script:Cancelar     = $false
# Instantaneos de uma operacao. Descartados no inicio de cada modulo, nunca
# reaproveitados entre cliques: decisao sobre encerrar processo clinico nao
# pode usar dado de uma operacao anterior.
$script:SnapProc       = $null
$script:SnapSvc        = $null
$script:EmUsoOperacao  = $null
# Quando cada instantaneo foi tirado, e por quanto tempo ele vale. O prazo e em
# SEGUNDOS de propósito: um Aplicar longo passa minutos entre o primeiro e o
# ultimo encerramento, e worker de inferencia nasce nesse meio.
$script:SnapProcEm       = $null
$script:EmUsoOperacaoEm  = $null
$script:SnapMaxSeg       = 15
# Lista de processos JA CLASSIFICADA, por operacao. Duas fatias, porque
# -TodasAsSessoes responde outra pergunta. Ver Get-ProcessosSessao: e retrato
# para MOSTRAR, e nao entra em decisao de encerrar.
$script:SnapSessao       = @{}
$script:SnapGrupos       = $null
$script:Ocupado      = $false

# Paleta: escala de dose (azul frio -> verde -> amarelo -> vermelho quente)
$script:Cor = @{
    Fundo    = [System.Drawing.Color]::FromArgb(18, 22, 28)
    Painel   = [System.Drawing.Color]::FromArgb(28, 34, 43)
    Borda    = [System.Drawing.Color]::FromArgb(45, 54, 66)
    Texto    = [System.Drawing.Color]::FromArgb(214, 222, 232)
    Fraco    = [System.Drawing.Color]::FromArgb(130, 143, 158)
    Titulo   = [System.Drawing.Color]::FromArgb(96, 190, 255)
    Ok       = [System.Drawing.Color]::FromArgb(86, 214, 140)
    Alerta   = [System.Drawing.Color]::FromArgb(248, 196, 74)
    Critico  = [System.Drawing.Color]::FromArgb(255, 108, 104)
    Acao     = [System.Drawing.Color]::FromArgb(186, 156, 255)
    Botao    = [System.Drawing.Color]::FromArgb(38, 47, 59)
}

# =====================================================================
# 3. FUNCOES DE APOIO
# =====================================================================
function Pump { [System.Windows.Forms.Application]::DoEvents() }

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -lt 0)   { return 'n/d' }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} bytes' -f $Bytes)
}

function Write-Log {
    param(
        [string]$Texto = '',
        [ValidateSet('INFO','OK','ALERTA','CRITICO','TITULO','DADO','ACAO','BRUTO')]
        [string]$Nivel = 'INFO'
    )
    if (-not $script:Log) { return }

    switch ($Nivel) {
        'TITULO'  { $cor = $script:Cor.Titulo;  $marca = '' }
        'OK'      { $cor = $script:Cor.Ok;      $marca = '[ OK ] ' }
        'ALERTA'  { $cor = $script:Cor.Alerta;  $marca = '[ !! ] ' }
        'CRITICO' { $cor = $script:Cor.Critico; $marca = '[ XX ] ' }
        'ACAO'    { $cor = $script:Cor.Acao;    $marca = '' }
        'DADO'    { $cor = $script:Cor.Fraco;   $marca = '' }
        'BRUTO'   { $cor = $script:Cor.Fraco;   $marca = '' }
        default   { $cor = $script:Cor.Texto;   $marca = '[ .. ] ' }
    }

    if ($Nivel -eq 'DADO' -or $Nivel -eq 'ACAO' -or $Nivel -eq 'BRUTO') {
        $linha = '           ' + $marca + $Texto
    } elseif ([string]::IsNullOrWhiteSpace($Texto)) {
        $linha = ''
    } else {
        $linha = ('[{0}] ' -f (Get-Date -Format 'HH:mm:ss')) + $marca + $Texto
    }

    $script:Log.SelectionStart  = $script:Log.TextLength
    $script:Log.SelectionLength = 0
    $script:Log.SelectionColor  = $cor
    $script:Log.AppendText($linha + "`r`n")
    $script:Log.SelectionColor  = $script:Cor.Texto
    $script:Log.ScrollToCaret()
    Pump
}

function Write-Titulo {
    param([string]$Texto)
    Write-Log ''
    Write-Log ('== {0} {1}' -f $Texto.ToUpper(), ('=' * [Math]::Max(4, 66 - $Texto.Length))) 'TITULO'
}

function Add-Achado {
    param(
        [ValidateSet('OK','ALERTA','CRITICO')][string]$Severidade,
        [string]$Titulo,
        [string]$Recomendacao = '',
        [string]$Categoria = 'Geral',
        [ValidateSet('Alto','Medio','Baixo')][string]$Impacto = 'Medio'
    )
    Write-Log $Titulo $Severidade
    if ($Recomendacao) { Write-Log ('-> ' + $Recomendacao) 'ACAO' }
    if ($Severidade -ne 'OK') {
        [void]$script:Achados.Add([pscustomobject]@{
            Severidade   = $Severidade
            Titulo       = $Titulo
            Recomendacao = $Recomendacao
            Categoria    = $Categoria
            Impacto      = $Impacto
        })
    }
}

function Set-Status {
    param([string]$Texto)
    if ($script:LblStatus) { $script:LblStatus.Text = $Texto; Pump }
}

function Set-Ocupado {
    param([bool]$Valor)
    $script:Ocupado = $Valor
    foreach ($b in @($script:BtnMod1, $script:BtnMod2, $script:BtnMod3,
                     $script:BtnGrandes, $script:BtnSessao, $script:BtnReiniciar,
                     $script:BtnLimparSel, $script:BtnDesfazer, $script:BtnRestaurar,
                     $script:BtnReverterSes)) {
        if ($b) { $b.Enabled = -not $Valor }
    }
    if ($script:BtnCancelar) { $script:BtnCancelar.Enabled = $Valor }
    $script:Barra.Style = if ($Valor) { 'Marquee' } else { 'Blocks' }
    if (-not $Valor) { $script:Barra.Value = 0 }
    $script:Form.Cursor = if ($Valor) { [System.Windows.Forms.Cursors]::AppStarting } else { [System.Windows.Forms.Cursors]::Default }
    Pump
}

# --- medicao de tamanho (robocopy e rapido e nao precisa de admin) ---
function Get-TamanhoPasta {
    param([string]$Caminho, [int]$TimeoutSeg = 90)
    if ([string]::IsNullOrWhiteSpace($Caminho)) { return -1 }
    try { if (-not (Test-Path -LiteralPath $Caminho -ErrorAction Stop)) { return -1 } } catch { return -1 }

    $bytes  = -1
    $saida  = Join-Path $env:TEMP ('kitrt_{0}.txt' -f ([guid]::NewGuid().ToString('N')))
    try {
        $argumentos = '"{0}" NULL /L /S /NJH /BYTES /NC /NFL /NDL /XJ /R:0 /W:0' -f $Caminho.TrimEnd('\')
        $proc = Start-Process -FilePath 'robocopy.exe' -ArgumentList $argumentos -WindowStyle Hidden `
                              -PassThru -RedirectStandardOutput $saida -WorkingDirectory $env:TEMP
        $inicio = Get-Date
        while (-not $proc.HasExited) {
            Pump
            Start-Sleep -Milliseconds 120
            if (((Get-Date) - $inicio).TotalSeconds -gt $TimeoutSeg) { try { $proc.Kill() } catch { }; break }
        }
        Start-Sleep -Milliseconds 100
        $texto = Get-Content -LiteralPath $saida -Raw -ErrorAction SilentlyContinue
        if ($texto -and ($texto -match '(?m)^\s*Bytes\s*:\s*(\d+)')) { $bytes = [double]$Matches[1] }
    } catch { $bytes = -1 }
    finally { Remove-Item -LiteralPath $saida -Force -ErrorAction SilentlyContinue }

    if ($bytes -lt 0) { $bytes = Get-TamanhoPastaNet -Caminho $Caminho }
    return $bytes
}

function Get-TamanhoPastaNet {
    param([string]$Caminho)
    $total = 0.0
    $conta = 0
    $pilha = New-Object System.Collections.Stack
    $pilha.Push($Caminho)
    while ($pilha.Count -gt 0) {
        if ($script:Cancelar) { break }
        $dir = $pilha.Pop()
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
                try { $total += (New-Object System.IO.FileInfo $f).Length } catch { }
            }
        } catch { }
        try {
            foreach ($d in [System.IO.Directory]::EnumerateDirectories($dir)) { $pilha.Push($d) }
        } catch { }
        $conta++
        if ($conta % 60 -eq 0) { Pump }
    }
    return $total
}

function Get-MaioresArquivos {
    param([string]$Raiz, [int]$Top = 15, [int]$TimeoutSeg = 60, [int]$MinimoMB = 50)
    $lista  = New-Object System.Collections.ArrayList
    $inicio = Get-Date
    $conta  = 0
    $pilha  = New-Object System.Collections.Stack
    $pilha.Push($Raiz)
    while ($pilha.Count -gt 0) {
        if ($script:Cancelar) { break }
        if (((Get-Date) - $inicio).TotalSeconds -gt $TimeoutSeg) { break }
        $dir = $pilha.Pop()
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
                try {
                    $fi = New-Object System.IO.FileInfo $f
                    if ($fi.Length -ge ($MinimoMB * 1MB)) { [void]$lista.Add($fi) }
                } catch { }
            }
        } catch { }
        try {
            foreach ($d in [System.IO.Directory]::EnumerateDirectories($dir)) { $pilha.Push($d) }
        } catch { }
        $conta++
        if ($conta % 50 -eq 0) { Pump }
    }
    return ($lista | Sort-Object Length -Descending | Select-Object -First $Top)
}

function Test-PortaTcp {
    param([string]$Alvo, [int]$Porta = 445, [int]$TimeoutMs = 1500)
    $ok  = $false
    $rel = [System.Diagnostics.Stopwatch]::StartNew()
    $cli = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $cli.BeginConnect($Alvo, $Porta, $null, $null)
        if ($ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            try { $cli.EndConnect($ar); $ok = $true } catch { $ok = $false }
        }
    } catch { $ok = $false }
    finally { try { $cli.Close() } catch { }; $rel.Stop() }
    return [pscustomobject]@{ Ok = $ok; Ms = [int]$rel.ElapsedMilliseconds }
}

function Get-ValorReg {
    param([string]$Caminho, [string]$Nome)
    try { return (Get-ItemProperty -Path $Caminho -Name $Nome -ErrorAction Stop).$Nome } catch { return $null }
}

# =====================================================================
# 4. MODULO 1 - PREPARAR AMBIENTE
# =====================================================================
function Get-ConteudoPreparador {
    # ------------------------------------------------------------------
    # Rotina de preparacao de ambiente, embutida no aplicativo.
    # Para editar: altere aqui dentro. E gravada em ANSI (1252) antes de rodar.
    # ------------------------------------------------------------------
    # VAZIO NA VERSAO UNIVERSAL, e de proposito.
    #
    # Aqui viviam 358 linhas de batch que preparavam a arvore de trabalho de UM
    # hospital: portais em Excel copiados de um compartilhamento, atalhos na area
    # de trabalho, fonte de codigo de barras, mapeamento de unidade de rede e
    # testes de alcance de tres servidores por nome. Saiu junto a rede daquele
    # hospital: o dominio, o fileserver e dois IPs internos.
    #
    # Devolver vazio NAO pode virar sucesso silencioso, e por isso Save-Preparador
    # ganhou uma guarda: string vazia faz ele devolver $null em vez de gravar um
    # .bat de zero byte. Sem ela, o Modulo 1 executaria um batch vazio, que sai
    # com codigo 0 sem imprimir nada, e o log diria "Executando:" e mais nada -
    # exatamente o defeito que este projeto persegue desde o xcopy silenciado.
    #
    # Com $null, Invoke-PrepararAmbiente explica e pede um .bat ao usuario. O
    # Modulo 1 da versao universal e "rode o seu script de preparacao", que e uma
    # funcao legitima em qualquer maquina.
    #
    # Se um dia houver rotina universal de preparacao, ela nasce aqui.
    return ''
}

function Save-Preparador {
    try {
        if (-not (Test-Path -LiteralPath $script:PastaEstado)) {
            New-Item -ItemType Directory -Path $script:PastaEstado -Force | Out-Null
        }
        # Sem conteudo, NAO grava. Um .bat de zero byte roda, sai com codigo 0 e
        # nao imprime nada: o log diria "Executando:" e depois nada, e quem le
        # concluiria que a preparacao rodou. Devolver $null faz o Modulo 1
        # explicar e pedir um script ao usuario.
        $conteudo = Get-ConteudoPreparador
        if (-not "$conteudo".Trim()) { return $null }
        $destino = Join-Path $script:PastaEstado 'Preparar_Workstation.bat'
        [System.IO.File]::WriteAllText($destino, $conteudo, [System.Text.Encoding]::GetEncoding(1252))
        return $destino
    } catch {
        Write-Log ('Nao foi possivel gravar a rotina de preparacao: {0}' -f $_.Exception.Message) 'CRITICO'
        return $null
    }
}

function Test-ArquivoEmUso {
    param([string]$Caminho)
    if (-not (Test-Path -LiteralPath $Caminho)) { return $false }
    try {
        $fs = [System.IO.File]::Open($Caminho, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch { return $true }
}

function Test-PortalEmDia {
    param([string]$Nome)
    $local = Join-Path $script:PastaClinica $Nome
    if (-not (Test-Path -LiteralPath $local)) {
        Add-Achado 'ALERTA' ('{0} ausente no disco local' -f $Nome) 'Rode o Modulo 1 (Preparar ambiente).' 'Ambiente' 'Alto'
        return
    }
    $fl = Get-Item -LiteralPath $local -ErrorAction SilentlyContinue
    $aberto = Test-ArquivoEmUso -Caminho $local

    $rede = Join-Path $script:PastaClinicaRede $Nome
    $fr = $null
    try { if (Test-Path -LiteralPath $rede -ErrorAction Stop) { $fr = Get-Item -LiteralPath $rede -ErrorAction Stop } } catch { }

    if (-not $fr) {
        Write-Log ('{0}: {1}, alterado em {2:dd/MM/yyyy HH:mm} (rede indisponivel para comparar)' -f $Nome, (Format-Bytes $fl.Length), $fl.LastWriteTime) 'DADO'
        return
    }

    Write-Log ('{0}' -f $Nome) 'DADO'
    Write-Log ('   C: ...: {0,10}  {1:dd/MM/yyyy HH:mm}{2}' -f (Format-Bytes $fl.Length), $fl.LastWriteTime, $(if ($aberto) { '   [ABERTO AGORA]' } else { '' })) 'DADO'
    Write-Log ('   rede .: {0,10}  {1:dd/MM/yyyy HH:mm}' -f (Format-Bytes $fr.Length), $fr.LastWriteTime) 'DADO'

    $difMin = [Math]::Round(($fr.LastWriteTime - $fl.LastWriteTime).TotalMinutes)

    if ($aberto) {
        Add-Achado 'ALERTA' ('{0} esta aberto agora nesta maquina' -f $Nome) 'Enquanto estiver aberto, a preparacao nao consegue substituir a copia local. Feche o Excel e rode o Modulo 1 de novo.' 'Ambiente' 'Alto'
        return
    }
    if ($difMin -gt 2) {
        Add-Achado 'ALERTA' ('{0}: a copia do C: esta {1} atras da rede' -f $Nome, $(if ($difMin -ge 1440) { "$([Math]::Round($difMin/1440)) dia(s)" } else { "$difMin minuto(s)" })) 'A copia local nao foi atualizada. Se o arquivo estava aberto durante a preparacao, feche o Excel e rode o Modulo 1 de novo.' 'Ambiente' 'Alto'
    } elseif ($difMin -lt -2) {
        Add-Achado 'ALERTA' ('{0}: a copia do C: esta MAIS NOVA que a da rede' -f $Nome) 'Alguem abriu e salvou o arquivo nesta maquina. A preparacao nao sobrescreve por seguranca. Se a versao boa e a da rede, renomeie a copia local e rode o Modulo 1 de novo.' 'Ambiente' 'Alto'
    } else {
        Add-Achado 'OK' ('{0} em dia com a rede' -f $Nome) '' 'Ambiente'
    }
}

function Invoke-PrepararAmbiente {
    Write-Titulo 'Modulo 1 - Preparar ambiente de radioterapia'

    $bat = $null
    if ($script:ArquivoBatExterno -and (Test-Path -LiteralPath $script:ArquivoBatExterno)) {
        $bat = $script:ArquivoBatExterno
        Write-Log ('Usando a versao externa configurada: {0}' -f $bat) 'DADO'
    } else {
        $bat = Save-Preparador
        if (-not $bat) {
            # Esta versao nao traz rotina embutida: a que existia era da arvore de
            # trabalho de um hospital so. Dizer isso e melhor que abrir um dialogo
            # sem explicar por que ele apareceu.
            Write-Log 'Esta versao nao traz rotina de preparacao embutida.' 'DADO'
            Write-Log 'O Modulo 1 roda o script de preparacao que VOCE indicar.' 'DADO'
            Write-Log 'Selecione um arquivo .bat ou .cmd de preparacao.' 'ACAO'
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = 'Script de preparacao (*.bat;*.cmd)|*.bat;*.cmd'
            $dlg.Title  = 'Selecione o script de preparacao de ambiente'
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                Write-Log 'Preparacao cancelada.' 'ALERTA'
                return
            }
            $bat = $dlg.FileName
        } else {
            Write-Log ('Rotina embutida no aplicativo (versao {0}), gravada em ANSI 1252.' -f $script:VersaoPreparador) 'DADO'
        }
    }

    Write-Log ('Executando: {0}' -f $bat)
    Write-Log 'A saida original do script aparece abaixo, linha a linha.' 'DADO'
    Write-Log ''

    $saida = Join-Path $env:TEMP ('kitrt_mod1_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $comando = '/c ""{0}" < nul > "{1}" 2>&1"' -f $bat, $saida
    $impressas = 0

    try {
        $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList $comando -WindowStyle Hidden -PassThru
    } catch {
        Write-Log ('Falha ao iniciar o script: {0}' -f $_.Exception.Message) 'CRITICO'
        return
    }

    $lerSaida = {
        if (-not (Test-Path -LiteralPath $saida)) { return }
        try {
            $fs = New-Object System.IO.FileStream($saida, [System.IO.FileMode]::Open,
                                                  [System.IO.FileAccess]::Read,
                                                  [System.IO.FileShare]::ReadWrite)
            $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::GetEncoding(1252))
            $todo = $sr.ReadToEnd()
            $sr.Close(); $fs.Close()
            $linhas = $todo -split "`r?`n"
            for ($i = $impressas; $i -lt $linhas.Count; $i++) {
                $l = $linhas[$i]
                if ($i -eq ($linhas.Count - 1) -and -not $proc.HasExited) { break }  # linha pode estar incompleta
                if ($l -match '\[ERRO|\[AVISO|\[ALERTA') { Write-Log $l.Trim() 'ALERTA' }
                elseif ($l -match '\[SUCESSO|\[ATUALIZADO|sucesso')       { Write-Log $l.Trim() 'DADO' }
                elseif ($l.Trim())                                        { Write-Log $l.TrimEnd() 'BRUTO' }
                $script:impressasTmp = $i + 1
            }
        } catch { }
    }

    $script:impressasTmp = 0
    while (-not $proc.HasExited) {
        & $lerSaida
        $impressas = $script:impressasTmp
        Pump
        Start-Sleep -Milliseconds 250
        if ($script:Cancelar) {
            try { $proc.Kill() } catch { }
            Write-Log 'Preparacao interrompida pelo usuario.' 'ALERTA'
            break
        }
    }
    Start-Sleep -Milliseconds 300
    & $lerSaida

    Write-Log ''
    if ($script:Cancelar) { return }
    Write-Log 'Preparacao de ambiente concluida.' 'OK'
    Write-Log ('Saida completa salva em: {0}' -f $saida) 'DADO'

    # conferencia rapida do resultado
    $pops = Join-Path $script:PastaClinica 'POPs - RADIOTERAPIA'
    if (Test-Path -LiteralPath $pops) {
        $arq = @(Get-ChildItem -LiteralPath $pops -File -Recurse -ErrorAction SilentlyContinue)
        if ($arq.Count -gt 0) {
            $recente = ($arq | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
            Write-Log ('POPs - RADIOTERAPIA: {0} arquivo(s), {1}, mais recente de {2:dd/MM/yyyy}' -f $arq.Count, (Format-Bytes (($arq | Measure-Object Length -Sum).Sum)), $recente.LastWriteTime) 'OK'
        } else {
            Write-Log 'POPs - RADIOTERAPIA: pasta criada, mas vazia. Confira o caminho na rede.' 'ALERTA'
        }
    }

    $kit = Join-Path $script:PastaClinica 'UTILITARIOS\PREPARAR WORKSTATION RADIOTERAPIA'
    if (Test-Path -LiteralPath $kit) {
        $arq = @(Get-ChildItem -LiteralPath $kit -File -Recurse -ErrorAction SilentlyContinue)
        if ($arq.Count -gt 0) {
            $recente = ($arq | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
            Write-Log ('Kit de Suporte no C:: {0} arquivo(s), mais recente {1} de {2:dd/MM/yyyy HH:mm}' -f $arq.Count, $recente.Name, $recente.LastWriteTime) 'OK'
        }
    }

    Write-Log '' 'DADO'
    Write-Log 'Conferencia dos portais (C: contra a rede):' 'DADO'
    foreach ($p in @('PORTAL - RADIOTERAPIA.xlsb', 'PORTAL - FICHA TÉCNICA.xlsb')) {
        Test-PortalEmDia -Nome $p
    }
}

# =====================================================================
# 4B. CATALOGO DO QUE E SEGURO ENCERRAR E DESATIVAR
# =====================================================================

# Nunca encerrar nem desativar: clinico, seguranca, rede corporativa,
# drivers e qualquer programa que possa ter documento aberto.
$script:Protegidos = 'wfica32|wfcrun32|CDViewer|SelfService|Receiver|concentr|CtxWebHelper|AuthManSvr|redirector|HdxRtcEngine|CtxCFRUI|Citrix|tasy|TasyAgent|CentBrowser|javaw|^java$|jp2launcher|Wheb|Philips|CcmExec|ntrtscan|tmlisten|TMBM|PccNTMon|ShowMsg|smartscreen|unsecapp|SearchProtocolHost|SearchFilterHost|DSASvc|QualysAgent|stAgent|Cortex|cyserver|cytray|cyvera|traps|LsAgent|Quest|OnDemand|ODMActiveDirectory|SecureConnector|ARIA|Eclipse|Varian|Vitrea|MIM|MOSAIQ|RayStation|Monaco|Velocity|Osirix|Horos|Weasis|dicom|PACS|EXCEL|WINWORD|POWERPNT|OUTLOOK|MSACCESS|onenote|StickyNot|msedge|chrome|firefox|iexplore|notepad|wordpad|Acrobat|AcroRd32|AnyConnect|GlobalProtect|FortiClient|Pulse|CcmExec|CmRcService|ccmsetup|CSFalcon|CSAgent|Sophos|SAVService|^mfe|masvc|macmnsvc|McShield|ccSvcHst|SepMaster|ZSA|stAgent|nsdiag|ivanti|LANDesk|Forcepoint|splunk|nxlog|MsMpEng|NisSrv|SecurityHealth|Realtek|RtkAud|IDTNC|Synaptics|igfx|nvcontainer|audiodg|System|Idle|Registry|smss|csrss|wininit|winlogon|^services$|lsass|svchost|fontdrvhost|dwm|explorer|RuntimeBroker|sihost|ctfmon|taskhostw|dllhost|conhost|WmiPrvSE|powershell|pwsh|LogonUI|SearchIndexer'

function Get-CatalogoProcessos {
    @(
        [pscustomobject]@{ Padrao = '^OneDrive$|^FileCoAuth$';                                    Rotulo = 'OneDrive (sincronismo)';        Nivel = 'Sempre';   Nota = 'Volta ao abrir o OneDrive ou no proximo logon.' }
        [pscustomobject]@{ Padrao = '^Teams$|^ms-teams$|^msteams';                                 Rotulo = 'Microsoft Teams';               Nivel = 'Sempre';   Nota = 'As conversas ficam no servidor. Volta ao abrir.' }
        [pscustomobject]@{ Padrao = 'Update$|Updater$|^GoogleUpdate|EdgeUpdate|^AdobeARM$|^AdobeGCClient$|^armsvc$|^jusched$|^SquirrelUpdate|^OfficeC2RClient$'; Rotulo = 'Atualizadores automaticos'; Nivel = 'Sempre'; Nota = 'Voltam sozinhos quando houver atualizacao.' }
        [pscustomobject]@{ Padrao = '^CCXProcess$|^Creative Cloud|^CCLibrary|^AdobeIPCBroker$|^AdobeNotificationClient$'; Rotulo = 'Adobe Creative Cloud (fundo)'; Nivel = 'Sempre'; Nota = '' }
        [pscustomobject]@{ Padrao = '^GameBar|^XboxApp|^GamingServices|^XboxGame|^GameBarPresence'; Rotulo = 'Xbox e Game Bar';              Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^NVIDIA Share$|^NVIDIA Web Helper';                           Rotulo = 'Sobreposicao NVIDIA';           Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^YourPhone$|^PhoneExperienceHost$';                           Rotulo = 'Vincular ao telefone';          Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^Widgets$|^WidgetService$';                                   Rotulo = 'Widgets do Windows';            Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^SearchApp$|^Cortana$|^Copilot';                              Rotulo = 'Pesquisa e Copilot';            Nivel = 'Sempre';   Nota = 'O Windows recarrega sozinho quando precisar.' }
        [pscustomobject]@{ Padrao = '^SupportAssist|^DellSupport|^HPSupport|^HPPrintScan|^Vantage|^ImController$|^McUICnt$|^HPNotifications'; Rotulo = 'Assistentes do fabricante'; Nivel = 'Sempre'; Nota = '' }
        [pscustomobject]@{ Padrao = '^iTunesHelper$|^AppleMobileDevice|^iPodService$';             Rotulo = 'Servicos Apple';                Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^Spotify';                                                    Rotulo = 'Spotify';                       Nivel = 'SemJanela'; Nota = '' }
        [pscustomobject]@{ Padrao = '^Dropbox$|^Box$|^GoogleDriveFS$';                             Rotulo = 'Nuvem pessoal';                 Nivel = 'SemJanela'; Nota = '' }
        [pscustomobject]@{ Padrao = '^Zoom$|^Slack$|^Discord$|^Skype';                             Rotulo = 'Comunicadores';                 Nivel = 'SemJanela'; Nota = 'Encerrado apenas se nao houver janela aberta.' }
        [pscustomobject]@{ Padrao = '^Steam|^EpicGames';                                           Rotulo = 'Lojas de jogos';                Nivel = 'SemJanela'; Nota = '' }
    )
}

function Get-ProcessosEncerraveis {
    $saida = @()
    try { $todos = @(Get-Process -ErrorAction SilentlyContinue) } catch { return $saida }

    foreach ($c in (Get-CatalogoProcessos)) {
        $ps = @($todos | Where-Object { $_.ProcessName -match $c.Padrao -and $_.ProcessName -notmatch $script:Protegidos })
        if ($ps.Count -eq 0) { continue }
        $comJanela = @($ps | Where-Object { $_.MainWindowHandle -ne 0 })
        if ($c.Nivel -eq 'SemJanela' -and $comJanela.Count -gt 0) { continue }
        $saida += [pscustomobject]@{
            Rotulo  = $c.Rotulo
            Nomes   = @($ps | Select-Object -ExpandProperty ProcessName -Unique)
            Ids     = @($ps | Select-Object -ExpandProperty Id)
            Qtd     = $ps.Count
            Memoria = ($ps | Measure-Object WorkingSet64 -Sum).Sum
            Nota    = $c.Nota
        }
    }
    return $saida
}

function Test-PastasNoOneDrive {
    foreach ($k in @('Desktop', 'Personal')) {
        $v = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' $k
        if ($v -and ([Environment]::ExpandEnvironmentVariables($v)) -like '*OneDrive*') { return $true }
    }
    return $false
}

function Get-CatalogoInicializacao {
    @(
        [pscustomobject]@{ Padrao = 'OneDrive';                          Rotulo = 'OneDrive' }
        [pscustomobject]@{ Padrao = 'Teams';                             Rotulo = 'Microsoft Teams' }
        [pscustomobject]@{ Padrao = 'Spotify';                           Rotulo = 'Spotify' }
        [pscustomobject]@{ Padrao = 'Skype';                             Rotulo = 'Skype' }
        [pscustomobject]@{ Padrao = 'Zoom';                              Rotulo = 'Zoom' }
        [pscustomobject]@{ Padrao = 'Slack|Discord';                     Rotulo = 'Comunicador' }
        [pscustomobject]@{ Padrao = 'Steam|Epic';                        Rotulo = 'Loja de jogos' }
        [pscustomobject]@{ Padrao = 'Dropbox|Box Sync|GoogleDrive';      Rotulo = 'Nuvem pessoal' }
        [pscustomobject]@{ Padrao = 'Adobe|CCX|Creative|Acro';           Rotulo = 'Componente Adobe' }
        [pscustomobject]@{ Padrao = 'Update|Updater|jusched|Java';       Rotulo = 'Atualizador automatico' }
        [pscustomobject]@{ Padrao = 'iTunes|Apple|QuickTime';            Rotulo = 'Componente Apple' }
        [pscustomobject]@{ Padrao = 'SupportAssist|Dell|HP |Vantage|Lenovo'; Rotulo = 'Assistente do fabricante' }
        [pscustomobject]@{ Padrao = 'Cortana|Copilot|Widgets|YourPhone'; Rotulo = 'Extra do Windows' }
        [pscustomobject]@{ Padrao = 'Steam|Xbox|GameBar|NVIDIA';         Rotulo = 'Jogos e sobreposicoes' }
        [pscustomobject]@{ Padrao = 'Spark|Grammarly|CCleaner|Toolbar';  Rotulo = 'Utilitario opcional' }
    )
}

function Get-ItensInicializacao {
    $itens = @()
    $kfm = Test-PastasNoOneDrive

    try {
        $p = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -ErrorAction Stop
        foreach ($n in $p.PSObject.Properties.Name) {
            if ($n -like 'PS*') { continue }
            $itens += [pscustomobject]@{ Nome = $n; Tipo = 'Run'; Comando = "$($p.$n)" }
        }
    } catch { }
    try {
        foreach ($f in (Get-ChildItem -LiteralPath ([Environment]::GetFolderPath('Startup')) -File -ErrorAction Stop)) {
            $itens += [pscustomobject]@{ Nome = $f.Name; Tipo = 'StartupFolder'; Comando = $f.FullName }
        }
    } catch { }

    $cat = Get-CatalogoInicializacao
    $saida = @()
    foreach ($i in $itens) {
        $texto = ($i.Nome + ' ' + $i.Comando)
        if ($i.Nome -match 'MicrosoftEdgeAutoLaunch') {
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Seguro'; Rotulo = 'Edge abrindo sozinho no logon'; Seguro = $true }
            continue
        }
        if ($texto -match $script:Protegidos) {
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Protegido'; Rotulo = 'protegido (seguranca, rede ou clinico)'; Seguro = $false }
            continue
        }
        $achou = $null
        foreach ($c in $cat) { if ($texto -match $c.Padrao) { $achou = $c; break } }
        if ($achou) {
            $seguro = $true
            $rot = $achou.Rotulo
            if ($achou.Rotulo -eq 'OneDrive' -and $kfm) {
                $rot = 'OneDrive - ATENCAO: suas pastas estao dentro dele; a sincronizacao passa a ocorrer so quando voce abrir'
            }
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Seguro'; Rotulo = $rot; Seguro = $seguro }
        } else {
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Desconhecido'; Rotulo = 'nao reconhecido'; Seguro = $false }
        }
    }
    return $saida
}

function Get-AjustesPendentes {
    $lista = @()

    $vfx = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting'
    if ($vfx -ne 2) {
        $lista += [pscustomobject]@{ Id = 'EFEITOS'; Rotulo = 'Desligar animacoes, sombras e transparencia'; Nota = 'Deixa a resposta imediata em maquina fraca.' }
    }
    $bg = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' 'GlobalUserDisabled'
    if ($bg -ne 1) {
        $lista += [pscustomobject]@{ Id = 'SEGUNDOPLANO'; Rotulo = 'Impedir apps da Store de rodar em segundo plano'; Nota = 'Tira dezenas de processos do fundo.' }
    }
    $ss = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' '01'
    if ($ss -ne 1) {
        $lista += [pscustomobject]@{ Id = 'STORAGE'; Rotulo = 'Ligar a limpeza automatica do Windows'; Nota = 'Evita que o disco encha de novo.' }
    }
    $zonas = Get-ZonasCitrixPendentes
    if ($zonas.Count -gt 0) {
        $lista += [pscustomobject]@{ Id = 'CITRIXZONA'; Rotulo = ('Confiar nos enderecos do Citrix ({0})' -f ($zonas -join ', ')); Nota = 'Faz o login unico funcionar e o .ica abrir sozinho.' }
    }

    # barra de tarefas: pesquisa, visao de tarefas, widgets e botao de chat
    $sb = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' 'SearchboxTaskbarMode'
    $tv = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowTaskViewButton'
    $wd = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarDa'
    if ($sb -ne 0 -or $tv -ne 0 -or $wd -ne 0) {
        $lista += [pscustomobject]@{ Id = 'BARRATAREFAS'; Rotulo = 'Limpar a barra de tarefas (pesquisa, visao de tarefas, widgets)'; Nota = 'Some da barra e para de consumir memoria. Reversivel.' }
    }

    # icone do OneDrive no painel esquerdo do Explorer
    $odIco = Get-ValorReg 'HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' 'System.IsPinnedToNameSpaceTree'
    if ($odIco -ne 0) {
        $lista += [pscustomobject]@{ Id = 'ONEDRIVEICONE'; Rotulo = 'Tirar o icone do OneDrive do Explorer'; Nota = 'So esconde o atalho. Os arquivos e a conta continuam intactos.' }
    }

    $sug = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' 'SilentInstalledAppsEnabled'
    if ($sug -ne 0) {
        $lista += [pscustomobject]@{ Id = 'SUGESTOES'; Rotulo = 'Desligar sugestoes e apps promocionais do Windows'; Nota = 'Para de instalar aplicativo sozinho.' }
    }
    return $lista
}

function Get-MapeamentosMortos {
    $mortos = @()
    try {
        foreach ($c in (Get-CimInstance Win32_NetworkConnection -ErrorAction Stop)) {
            if ("$($c.ConnectionState)" -notmatch 'Connected|Conectado') {
                $mortos += [pscustomobject]@{ Letra = $c.LocalName; Destino = $c.RemoteName }
            }
        }
    } catch { }
    return $mortos
}

# =====================================================================
# 4C. CITRIX - ARIA, MOSAIQ E MONACO PUBLICADOS
# =====================================================================
function Get-InfoUrl {
    param([string]$Url)
    try { $u = [uri]$Url } catch { return $null }
    [pscustomobject]@{
        Url      = $Url
        Servidor = $u.Host
        Porta    = $u.Port
        Https    = ($u.Scheme -eq 'https')
        EhIp     = ($u.Host -match '^\d{1,3}(\.\d{1,3}){3}$')
    }
}

function Get-CitrixInstalado {
    # 1. processo em execucao entrega o caminho real do cliente
    foreach ($n in @('SelfService', 'SelfServicePlugin', 'Receiver', 'wfica32', 'wfcrun32', 'concentr', 'CDViewer')) {
        try {
            $p = Get-Process -Name $n -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($p -and $p.Path -and (Test-Path -LiteralPath $p.Path)) {
                $v = (Get-Item -LiteralPath $p.Path).VersionInfo
                return [pscustomobject]@{ Instalado = $true; Versao = "$($v.ProductVersion)"; Caminho = $p.Path }
            }
        } catch { }
    }
    # 2. arquivos conhecidos
    $arquivos = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Citrix\ICA Client\wfica32.exe'),
        (Join-Path $env:ProgramFiles 'Citrix\ICA Client\wfica32.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Citrix\ICA Client\SelfServicePlugin\SelfService.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Citrix\SelfServicePlugin\SelfService.exe')
    )
    foreach ($a in $arquivos) {
        if ($a -and (Test-Path -LiteralPath $a)) {
            try {
                $v = (Get-Item -LiteralPath $a).VersionInfo
                return [pscustomobject]@{ Instalado = $true; Versao = "$($v.ProductVersion)"; Caminho = $a }
            } catch { }
        }
    }
    # 3. registro
    foreach ($k in @('HKLM:\SOFTWARE\WOW6432Node\Citrix\InstallDetect\*', 'HKLM:\SOFTWARE\Citrix\InstallDetect\*',
                     'HKLM:\SOFTWARE\WOW6432Node\Citrix\ICA Client', 'HKLM:\SOFTWARE\Citrix\ICA Client')) {
        try {
            $r = Get-ItemProperty -Path $k -ErrorAction Stop | Select-Object -First 1
            if ($r) {
                $ver = "$($r.DisplayVersion)"
                if (-not $ver) { $ver = "$($r.Version)" }
                return [pscustomobject]@{ Instalado = $true; Versao = $ver; Caminho = 'registro' }
            }
        } catch { }
    }
    foreach ($k in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                     'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')) {
        try {
            $r = Get-ItemProperty -Path $k -ErrorAction SilentlyContinue |
                 Where-Object { "$($_.DisplayName)" -match 'Citrix (Workspace|Receiver|Online Plug)' } | Select-Object -First 1
            if ($r) { return [pscustomobject]@{ Instalado = $true; Versao = "$($r.DisplayVersion)"; Caminho = 'registro' } }
        } catch { }
    }
    return [pscustomobject]@{ Instalado = $false; Versao = ''; Caminho = '' }
}

function Get-CitrixProcessos {
    $nomes = @('wfica32', 'wfcrun32', 'CDViewer', 'SelfService', 'SelfServicePlugin', 'Receiver',
               'concentr', 'CtxWebHelper', 'AuthManSvr', 'redirector', 'HdxRtcEngine', 'CtxCFRUI')
    $ps = @(Get-Process -Name $nomes -ErrorAction SilentlyContinue)
    [pscustomobject]@{
        Rodando  = ($ps.Count -gt 0)
        Sessao   = (@($ps | Where-Object { $_.ProcessName -match 'wfica32|CDViewer|wfcrun32' }).Count -gt 0)
        Qtd      = $ps.Count
        Memoria  = $(if ($ps.Count -gt 0) { ($ps | Measure-Object WorkingSet64 -Sum).Sum } else { 0 })
        Nomes    = @($ps | Select-Object -ExpandProperty ProcessName -Unique)
    }
}

function Get-CitrixStoresConfigurados {
    $achados = @()
    $chaves = @(
        'HKCU:\Software\Citrix\Dazzle\Sites',
        'HKCU:\Software\Citrix\Receiver\SR\Store',
        'HKCU:\Software\Citrix\StoreFront\Stores',
        'HKCU:\Software\Citrix\Receiver\StoreFront\Stores'
    )
    foreach ($base in $chaves) {
        if (-not (Test-Path $base)) { continue }
        try {
            foreach ($k in (Get-ChildItem -Path $base -Recurse -ErrorAction SilentlyContinue)) {
                $p = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                foreach ($prop in $p.PSObject.Properties) {
                    if ($prop.Name -like 'PS*') { continue }
                    $v = "$($prop.Value)"
                    if ($v -match '^https?://') { $achados += $v }
                }
            }
        } catch { }
    }
    return ($achados | Select-Object -Unique)
}

function Test-UrlHttp {
    param([string]$Url, [int]$TimeoutSeg = 8)
    $rel = [System.Diagnostics.Stopwatch]::StartNew()
    $codigo = 0
    $erro = ''
    $final = $Url
    $certRuim = $false

    $tentativa = {
        param($u, $t)
        Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec $t -UseDefaultCredentials -MaximumRedirection 5 -ErrorAction Stop
    }

    try {
        $r = & $tentativa $Url $TimeoutSeg
        $codigo = [int]$r.StatusCode
        try { $final = "$($r.BaseResponse.ResponseUri)" } catch { }
    } catch {
        $msg = "$($_.Exception.Message)"
        if ($msg -match 'SSL|TLS|certificad|certificate|secure channel|confianca|trust') {
            # repete ignorando o certificado, so para saber se o servidor responde
            $antes = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
            try {
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
                $r = & $tentativa $Url $TimeoutSeg
                $codigo = [int]$r.StatusCode
                $certRuim = $true
                try { $final = "$($r.BaseResponse.ResponseUri)" } catch { }
            } catch {
                try { if ($_.Exception.Response) { $codigo = [int]$_.Exception.Response.StatusCode } } catch { }
                $erro = "$($_.Exception.Message)"
            } finally {
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $antes
            }
        } else {
            try {
                if ($_.Exception.Response) {
                    $codigo = [int]$_.Exception.Response.StatusCode
                    $final = "$($_.Exception.Response.ResponseUri)"
                }
            } catch { }
            $erro = $msg
        }
    }
    $rel.Stop()
    return [pscustomobject]@{
        Codigo = $codigo; Ms = [int]$rel.ElapsedMilliseconds; Erro = $erro
        UrlFinal = $final; CertificadoInvalido = $certRuim
    }
}

function Test-HttpsAlternativo {
    param([string]$Url)
    try { $u = [uri]$Url } catch { return $null }
    if ($u.Scheme -eq 'https') { return $null }
    $t = Test-PortaTcp -Alvo $u.Host -Porta 443 -TimeoutMs 1500
    if (-not $t.Ok) { return [pscustomobject]@{ Disponivel = $false; Url = ''; Codigo = 0 } }
    $alt = 'https://' + $u.Host + $u.AbsolutePath
    $r = Test-UrlHttp -Url $alt -TimeoutSeg 8
    return [pscustomobject]@{ Disponivel = ($r.Codigo -gt 0 -and $r.Codigo -lt 500); Url = $alt; Codigo = $r.Codigo }
}

function Test-ZonaIntranet {
    param([string]$Servidor)
    if ([string]::IsNullOrWhiteSpace($Servidor)) { return $true }
    if ($Servidor -match '^\d{1,3}(\.\d{1,3}){3}$') {
        $base = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges'
        if (-not (Test-Path $base)) { return $false }
        try {
            foreach ($k in (Get-ChildItem -Path $base -ErrorAction Stop)) {
                $p = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                $r = $p.PSObject.Properties[':Range']
                if ($r -and "$($r.Value)" -eq $Servidor) { return $true }
            }
        } catch { }
        return $false
    }
    if ($Servidor -notmatch '\.') { return $true }   # nome simples ja cai na Intranet
    $partes  = $Servidor.Split('.')
    $dominio = ($partes[-2..-1] -join '.')
    $sub     = $(if ($partes.Count -gt 2) { ($partes[0..($partes.Count - 3)] -join '.') } else { '' })
    $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\' + $dominio
    if ($sub) { $k = $k + '\' + $sub }
    return (Test-Path $k)
}

function Add-ZonaIntranet {
    param([string]$Servidor)
    $criadas = @()
    try {
        if ($Servidor -match '^\d{1,3}(\.\d{1,3}){3}$') {
            $base = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges'
            if (-not (Test-Path $base)) { New-Item -Path $base -Force | Out-Null }
            $n = 1
            while (Test-Path (Join-Path $base ('Range' + $n))) { $n++ }
            $k = Join-Path $base ('Range' + $n)
            New-Item -Path $k -Force | Out-Null
            New-ItemProperty -Path $k -Name ':Range' -Value $Servidor -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $k -Name 'http'   -Value 1 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $k -Name 'https'  -Value 1 -PropertyType DWord -Force | Out-Null
            $criadas += $k
        } else {
            $partes  = $Servidor.Split('.')
            if ($partes.Count -lt 2) { return @() }   # nome simples ja e Intranet
            $dominio = ($partes[-2..-1] -join '.')
            $sub     = $(if ($partes.Count -gt 2) { ($partes[0..($partes.Count - 3)] -join '.') } else { '' })
            $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\' + $dominio
            if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null; $criadas += $k }
            if ($sub) {
                $k = $k + '\' + $sub
                if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null; $criadas += $k }
            }
            New-ItemProperty -Path $k -Name 'http'  -Value 1 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $k -Name 'https' -Value 1 -PropertyType DWord -Force | Out-Null
        }
    } catch { }
    return $criadas
}

function Get-CitrixCaches {
    $c = @(
        (Join-Path $env:APPDATA 'ICAClient\Cache'),
        (Join-Path $env:APPDATA 'Citrix\SelfService'),
        (Join-Path $env:LOCALAPPDATA 'Citrix\SelfService'),
        (Join-Path $env:LOCALAPPDATA 'Citrix\ICA Client\Cache'),
        (Join-Path $env:APPDATA 'Citrix\Receiver\Cache')
    )
    return @($c | Where-Object { Test-Path -LiteralPath $_ })
}

function Get-ZonasCitrixPendentes {
    $pend = @()
    foreach ($s in $script:CitrixStores) {
        $i = Get-InfoUrl -Url $s.Url
        if (-not $i) { continue }
        if (-not (Test-ZonaIntranet -Servidor $i.Servidor)) { $pend += $i.Servidor }
    }
    return ($pend | Select-Object -Unique)
}

function Set-AjusteZonasCitrix {
    $criadas = @()
    foreach ($srv in (Get-ZonasCitrixPendentes)) {
        $r = Add-ZonaIntranet -Servidor $srv
        $criadas += $r
        Write-Log ('{0} adicionado aos sites da Intranet.' -f $srv) 'DADO'
    }
    if ($criadas.Count -gt 0) {
        $e = Read-Estado
        $ant = @()
        if ($e.ContainsKey('zonas_citrix')) { $ant = @($e['zonas_citrix']) }
        Save-Estado 'zonas_citrix' ($ant + $criadas)
    }
}

function Invoke-DiagCitrix {
    Write-Titulo 'Diagnostico do Citrix (ARIA, MOSAIQ e Monaco)'

    # ---- cliente ----
    $cli = Get-CitrixInstalado
    if ($cli.Instalado) {
        $maior = 0
        try { $maior = [int](("$($cli.Versao)" -split '\.')[0]) } catch { }
        if ($maior -gt 0 -and $maior -lt 19) {
            Add-Achado 'ALERTA' ('Citrix Receiver antigo: versao {0}' -f $cli.Versao) 'O Receiver 4.x saiu de suporte em 2018 e e substituido pelo Citrix Workspace. Atualizar exige o TI, mas resolve boa parte das falhas de abertura e de login unico.' 'Citrix' 'Alto'
        } else {
            Add-Achado 'OK' ('Citrix Workspace instalado (versao {0})' -f $cli.Versao) '' 'Citrix'
        }
    } else {
        Add-Achado 'CRITICO' 'Citrix Workspace nao encontrado nesta maquina' 'Sem o cliente instalado o ARIA/MOSAIQ/Monaco nao abrem. Instalacao exige o TI.' 'Citrix' 'Alto'
    }

    # ---- sessao em andamento ----
    $pr = Get-CitrixProcessos
    if ($pr.Rodando) {
        Write-Log ('Citrix em execucao: {0} processo(s), {1} - {2}' -f $pr.Qtd, (Format-Bytes $pr.Memoria), ($pr.Nomes -join ', ')) 'DADO'
        if ($pr.Sessao) {
            Add-Achado 'ALERTA' 'Ha sessao publicada aberta agora (ARIA, MOSAIQ ou Monaco)' 'O kit nunca encerra o Citrix. Feche a sessao antes de limpar o cache do Workspace.' 'Citrix' 'Medio'
        }
    } else {
        Write-Log 'Nenhum processo do Citrix em execucao.' 'DADO'
    }

    # ---- stores configurados x esperados ----
    $conf = Get-CitrixStoresConfigurados
    if ($conf.Count -gt 0) {
        Write-Log 'Stores ja configurados no Workspace deste usuario:' 'DADO'
        foreach ($c in $conf) { Write-Log ('- ' + $c) 'DADO' }
    } else {
        Write-Log 'Nenhum store gravado no perfil do usuario (acesso so pelo navegador).' 'DADO'
    }

    # ---- teste de cada destino ----
    foreach ($s in $script:CitrixStores) {
        if ($script:Cancelar) { return }
        $i = Get-InfoUrl -Url $s.Url
        if (-not $i) { Write-Log ('Endereco invalido na configuracao: {0}' -f $s.Url) 'ALERTA'; continue }

        Write-Log '' 'DADO'
        Write-Log ('--- {0} ---' -f $s.Nome) 'DADO'
        Write-Log ('Aplicativos: {0}' -f $s.Apps) 'DADO'
        Write-Log ('Endereco ..: {0}' -f $s.Url) 'DADO'
        Set-Status ('Testando ' + $s.Nome + '...')

        # nome -> IP, com tentativa pelo nome completo quando o curto falha
        $alvo    = $i.Servidor
        $urlTeste = $s.Url
        if (-not $i.EhIp) {
            $candidatos = @($i.Servidor)
            if ($i.Servidor -notmatch '\.') {
                foreach ($sf in (Get-SufixosDns)) {
                    $candidatos += ($i.Servidor + '.' + $sf.ToLower())
                }
            }
            $resolvido = $false
            foreach ($c in $candidatos) {
                try {
                    $rel = [System.Diagnostics.Stopwatch]::StartNew()
                    $ips = [System.Net.Dns]::GetHostAddresses($c) | Select-Object -ExpandProperty IPAddressToString
                    $rel.Stop()
                    Write-Log ('DNS .......: {0} -> {1} ({2} ms)' -f $c, ($ips -join ', '), $rel.ElapsedMilliseconds) 'DADO'
                    $alvo = $c
                    $resolvido = $true
                    if ($c -ne $i.Servidor) {
                        $urlTeste = $s.Url.Replace($i.Servidor, $c)
                        Add-Achado 'ALERTA' ('{0}: o nome curto {1} nao resolve, so o completo {2}' -f $s.Nome, $i.Servidor, $c) ('Troque o endereco do atalho e da configuracao para {0} . Sem isso o navegador nao acha o portal nesta estacao.' -f $urlTeste) 'Citrix' 'Alto'
                    }
                    break
                } catch { }
            }
            if (-not $resolvido) {
                Add-Achado 'CRITICO' ('{0}: o nome {1} nao resolve nesta rede, nem com o dominio completo' -f $s.Nome, $i.Servidor) 'A estacao nao esta enxergando o DNS interno desse servidor. Confirme cabo/VPN e o sufixo DNS da placa antes de abrir chamado.' 'Citrix' 'Alto'
                continue
            }
        }

        # portas
        $portas = @(80, 443)
        $abertas = @()
        foreach ($p in $portas) {
            $t = Test-PortaTcp -Alvo $alvo -Porta $p -TimeoutMs 1500
            if ($t.Ok) { $abertas += ('{0} ({1} ms)' -f $p, $t.Ms) }
        }
        # portas do protocolo ICA, usadas depois que o app e lancado
        $ica = Test-PortaTcp -Alvo $alvo -Porta 1494 -TimeoutMs 1200
        $rel2598 = Test-PortaTcp -Alvo $alvo -Porta 2598 -TimeoutMs 1200
        Write-Log ('Portas web : {0}' -f $(if ($abertas.Count -gt 0) { $abertas -join ' · ' } else { 'nenhuma respondeu' })) 'DADO'
        $notaIca = $(if ($ica.Ok -or $rel2598.Ok) { '' } else { '   (normal: quem atende o ICA e o servidor da aplicacao, nao o portal)' })
        Write-Log ('ICA 1494 ..: {0}   ·   Confiabilidade 2598: {1}{2}' -f $(if ($ica.Ok) { 'responde' } else { 'sem resposta' }), $(if ($rel2598.Ok) { 'responde' } else { 'sem resposta' }), $notaIca) 'DADO'

        if ($abertas.Count -eq 0) {
            Add-Achado 'CRITICO' ('{0}: servidor {1} sem resposta nas portas web' -f $s.Nome, $alvo) 'Enquanto isso nao voltar, os aplicativos desse destino nao abrem. Verifique a rede da estacao e, se as outras estacoes tambem falharem, avise o TI.' 'Citrix' 'Alto'
            continue
        }

        # pagina do store
        $r = Test-UrlHttp -Url $urlTeste
        if ($r.Codigo -ge 200 -and $r.Codigo -lt 400) {
            if ($r.Ms -gt $script:Lim.CitrixMs) {
                Add-Achado 'ALERTA' ('{0}: portal responde em {1} ms' -f $s.Nome, $r.Ms) 'Acima de 3 segundos a lista de aplicativos demora a aparecer. Limpar o cache do Workspace no Modulo 3 costuma resolver do lado da estacao.' 'Citrix' 'Medio'
            } else {
                Add-Achado 'OK' ('{0}: portal respondendo (HTTP {1}, {2} ms)' -f $s.Nome, $r.Codigo, $r.Ms) '' 'Citrix'
            }
        } elseif ($r.Codigo -eq 401 -or $r.Codigo -eq 403) {
            Add-Achado 'OK' ('{0}: portal no ar, pedindo login (HTTP {1})' -f $s.Nome, $r.Codigo) '' 'Citrix'
        } elseif ($r.Codigo -gt 0) {
            Add-Achado 'CRITICO' ('{0}: portal respondeu HTTP {1}' -f $s.Nome, $r.Codigo) 'O servidor esta no ar mas o store nao. Registre o codigo no chamado.' 'Citrix' 'Alto'
        } else {
            Add-Achado 'CRITICO' ('{0}: portal nao respondeu' -f $s.Nome) ('Detalhe: ' + $r.Erro) 'Citrix' 'Alto'
        }

        # zona de seguranca (sem isso o login unico falha e o .ica nao abre sozinho)
        if (Test-ZonaIntranet -Servidor $i.Servidor) {
            Write-Log 'Zona ......: ja esta na Intranet (login unico funciona)' 'DADO'
        } else {
            Add-Achado 'ALERTA' ('{0}: {1} nao esta na zona de Intranet' -f $s.Nome, $i.Servidor) 'E o que faz o navegador pedir senha de novo e nao abrir o arquivo .ica sozinho. O Modulo 3 corrige com um clique.' 'Citrix' 'Alto'
        }

        # HTTP x HTTPS: o atalho pode estar desatualizado
        if (-not $i.Https) {
            if ("$($r.UrlFinal)" -match '^https://') {
                Write-Log 'Seguranca .: o servidor redireciona sozinho para HTTPS. O atalho pode ficar como esta.' 'DADO'
            } else {
                $alt = Test-HttpsAlternativo -Url $urlTeste
                if ($alt -and $alt.Disponivel) {
                    Add-Achado 'ALERTA' ('{0}: o atalho usa HTTP, mas o servidor tambem atende em HTTPS' -f $s.Nome) ('Troque o atalho para {0} . Em HTTP o navegador bloqueia parte do login unico e o trafego vai sem criptografia.' -f $alt.Url) 'Citrix' 'Medio'
                } else {
                    Write-Log 'Seguranca .: endereco so em HTTP. O trafego vai sem criptografia.' 'DADO'
                }
            }
        }
        if ($r.CertificadoInvalido) {
            Add-Achado 'ALERTA' ('{0}: certificado HTTPS nao confiavel nesta estacao' -f $s.Nome) 'O navegador mostra aviso de site nao seguro e pode bloquear download e impressao. Peca ao TI a instalacao do certificado da autoridade interna.' 'Citrix' 'Alto'
        }
    }

    # ---- proxy ----
    $proxyOn = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyEnable'
    if ($proxyOn -eq 1) {
        $srv = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyServer'
        $exc = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyOverride'
        Write-Log '' 'DADO'
        Write-Log ('Proxy ligado: {0}' -f $srv) 'DADO'
        Write-Log ('Excecoes ...: {0}' -f $exc) 'DADO'
        $faltando = @()
        foreach ($s in $script:CitrixStores) {
            $i = Get-InfoUrl -Url $s.Url
            if ($i -and ("$exc" -notlike ('*' + $i.Servidor + '*'))) { $faltando += $i.Servidor }
        }
        if ($faltando.Count -gt 0) {
            Add-Achado 'ALERTA' ('Proxy sem excecao para: {0}' -f ($faltando -join ', ')) 'O trafego do Citrix esta passando pelo proxy sem necessidade, o que deixa a abertura lenta. Peca ao TI para incluir nas excecoes.' 'Citrix' 'Medio'
        }
    }

    # ---- cache local ----
    $caches = Get-CitrixCaches
    $tot = 0.0
    foreach ($c in $caches) {
        $b = Get-TamanhoPasta -Caminho $c -TimeoutSeg 45
        if ($b -gt 0) { $tot += $b; Write-Log ('{0,10}  {1}' -f (Format-Bytes $b), $c) 'DADO' }
    }
    $icas = @()
    foreach ($p in @($env:TEMP, (Join-Path $env:USERPROFILE 'Downloads'))) {
        try { $icas += @(Get-ChildItem -LiteralPath $p -Filter '*.ica' -File -Force -ErrorAction SilentlyContinue) } catch { }
    }
    if ($icas.Count -gt 0) { Write-Log ('Arquivos .ica soltos: {0}' -f $icas.Count) 'DADO' }

    if ($tot -gt 0) {
        Add-Achado 'OK' ('Cache do Citrix com {0} - este kit nunca o limpa' -f (Format-Bytes $tot)) '' 'Citrix'
    }
    if ($icas.Count -gt 5) {
        Add-Achado 'ALERTA' ('{0} arquivos .ica soltos' -f $icas.Count) 'Sao apenas lancadores baixados pelo navegador. O Modulo 3 remove os de Downloads dentro do periodo escolhido; o cache do Citrix fica intacto.' 'Citrix' 'Medio'
    }

    Write-Log '' 'DADO'
    Write-Log 'Se um destino falha e os outros funcionam, o problema e daquele servidor, nao da estacao.' 'DADO'
    Write-Log 'Se todos falham, e rede ou perfil da estacao.' 'DADO'
}

# =====================================================================
# 4D. INVENTARIO DA ESTACAO (retrato, somente leitura)
# =====================================================================
$script:ItensInv = New-Object System.Collections.ArrayList

function Add-ItemInv {
    param([string]$Categoria, [string]$Nome, [string]$Valor = '', [string]$Detalhe = '')
    [void]$script:ItensInv.Add([pscustomobject]@{
        Maquina = $env:COMPUTERNAME; Usuario = $env:USERNAME
        Data = (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Categoria = $Categoria; Nome = $Nome; Valor = $Valor; Detalhe = $Detalhe
    })
}

function Get-CatalogoClinico {
    @(
        [pscustomobject]@{ Chave = 'ARIA';        Padrao = '\bARIA\b|Varian|VMS\.|Vision.*Varian' }
        [pscustomobject]@{ Chave = 'Eclipse';     Padrao = 'Eclipse.*Varian|Varian.*Eclipse|ExternalBeam' }
        [pscustomobject]@{ Chave = 'MOSAIQ';      Padrao = '\bMOSAIQ\b|\bIMPAC\b' }
        [pscustomobject]@{ Chave = 'Monaco';      Padrao = '\bMonaco\b.*Elekta|Elekta.*\bMonaco\b|\bFocal\b.*Elekta' }
        [pscustomobject]@{ Chave = 'Vitrea';      Padrao = '\bVitrea|Vital Images|VspApp|VspMgmt' }
        [pscustomobject]@{ Chave = 'Mirada';      Padrao = '\bMirada\b' }
        [pscustomobject]@{ Chave = 'MIM';         Padrao = 'MIM Software|MIMSoftware|MIMMaestro' }
        [pscustomobject]@{ Chave = 'RayStation';  Padrao = 'RayStation|RaySearch' }
        [pscustomobject]@{ Chave = 'Visor DICOM'; Padrao = 'RadiAnt|Weasis|Horos|OsiriX|MicroDicom|\bOnis\b|K-PACS' }
        [pscustomobject]@{ Chave = 'Citrix';      Padrao = '\bCitrix\b' }
        [pscustomobject]@{ Chave = 'Prontuario';  Padrao = '\bTasy\b|Soul MV|CentBrowser' }
        [pscustomobject]@{ Chave = 'Assinatura';  Padrao = 'SafeSign|SafeNet|Watchdata|eToken|Bird ?ID|VIDaaS' }
        [pscustomobject]@{ Chave = 'PDF';         Padrao = 'Acrobat|Foxit|PDFtk|Nitro|PDF-XChange' }
        [pscustomobject]@{ Chave = 'Office';      Padrao = 'Microsoft 365|Microsoft Office|LibreOffice' }
    )
}

function Get-CatalogoServidor {
    @(
        [pscustomobject]@{ Chave = 'SQL Server'; Padrao = '^sqlservr$|^SQLAgent|^MSSQL' }
        [pscustomobject]@{ Chave = 'Tomcat';     Padrao = 'Tomcat|catalina' }
        [pscustomobject]@{ Chave = 'IIS';        Padrao = '^w3wp$|^inetinfo$|^W3SVC$' }
        [pscustomobject]@{ Chave = 'Apache';     Padrao = '^httpd$|^Apache' }
        [pscustomobject]@{ Chave = 'Banco local'; Padrao = '^postgres|^mysqld$|^oracle|^firebird' }
    )
}

function Get-ProgramasInstalados {
    $chaves = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $prog = @()
    foreach ($k in $chaves) {
        try {
            $prog += Get-ItemProperty -Path $k -ErrorAction SilentlyContinue |
                     Where-Object { $_.DisplayName -and -not $_.SystemComponent } |
                     Select-Object @{n = 'Nome'; e = { "$($_.DisplayName)".Trim() } },
                                   @{n = 'Versao'; e = { "$($_.DisplayVersion)" } },
                                   @{n = 'Fabricante'; e = { "$($_.Publisher)" } }
        } catch { }
    }
    return ($prog | Sort-Object Nome -Unique)
}

function Invoke-InvIdentificacao {
    Write-Titulo 'Inventario · identificacao'
    try {
        $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $cp = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $bi = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue

        Write-Log ('Maquina .......: {0}   usuario: {1}\{2}' -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME) 'DADO'
        Write-Log ('Sistema .......: {0} build {1}' -f $so.Caption, $so.BuildNumber) 'DADO'
        Write-Log ('Equipamento ...: {0} {1}   serie {2}' -f $cs.Manufacturer, $cs.Model, $bi.SerialNumber) 'DADO'
        Write-Log ('Processador ...: {0} ({1} nucleos)' -f "$($cp.Name)".Trim(), $cp.NumberOfCores) 'DADO'
        Write-Log ('Memoria .......: {0}' -f (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))) 'DADO'
        Write-Log ('Windows desde .: {0:dd/MM/yyyy}' -f $so.InstallDate) 'DADO'

        Add-ItemInv 'Identificacao' 'Maquina' $env:COMPUTERNAME
        Add-ItemInv 'Identificacao' 'Sistema' ("$($so.Caption) build $($so.BuildNumber)")
        Add-ItemInv 'Identificacao' 'Modelo' ("$($cs.Manufacturer) $($cs.Model)")
        Add-ItemInv 'Identificacao' 'Serie' ("$($bi.SerialNumber)")
        Add-ItemInv 'Identificacao' 'Processador' ("$($cp.Name)".Trim())
        Add-ItemInv 'Identificacao' 'Memoria' (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))

        if (($so.TotalVisibleMemorySize * 1KB) -lt 8GB) {
            Add-Achado 'ALERTA' ('Memoria total de {0}' -f (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))) 'Abaixo de 8 GB o Excel em rede mais o navegador ja saturam a maquina. Registre no chamado.' 'Hardware' 'Medio'
        }
    } catch { }

    try {
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)) {
            if (-not $d.Size) { continue }
            Add-ItemInv 'Disco' $d.DeviceID ('{0}% livre' -f [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1)) ('{0} de {1}' -f (Format-Bytes $d.FreeSpace), (Format-Bytes $d.Size))
        }
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=4' -ErrorAction SilentlyContinue)) {
            Write-Log ('{0} -> {1}' -f $d.DeviceID, $d.ProviderName) 'DADO'
            Add-ItemInv 'Unidade de rede' $d.DeviceID "$($d.ProviderName)"
        }
    } catch { }
}

function Invoke-InvSoftware {
    Write-Titulo 'Inventario · software instalado'
    $prog = Get-ProgramasInstalados
    Write-Log ('{0} programas instalados' -f $prog.Count) 'DADO'
    foreach ($p in $prog) {
        Write-Log ('{0,-56} {1,-18} {2}' -f $p.Nome, $p.Versao, $p.Fabricante) 'DADO'
        Add-ItemInv 'Programa' $p.Nome $p.Versao $p.Fabricante
    }

    $procs = @(Get-Process -ErrorAction SilentlyContinue)
    $svcs = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)

    Write-Log '' 'DADO'
    Write-Log 'Software clinico:' 'DADO'
    foreach ($c in (Get-CatalogoClinico)) {
        $achou = @()
        $achou += @($prog  | Where-Object { $_.Nome -match $c.Padrao } | ForEach-Object { 'instalado: ' + $_.Nome })
        $achou += @($procs | Where-Object { $_.ProcessName -match $c.Padrao } | Select-Object -ExpandProperty ProcessName -Unique | ForEach-Object { 'processo: ' + $_ })
        $achou += @($svcs  | Where-Object { $_.Name -match $c.Padrao -or $_.DisplayName -match $c.Padrao } | Select-Object -ExpandProperty Name -Unique | ForEach-Object { 'servico: ' + $_ })
        $achou = @($achou | Select-Object -Unique)
        if ($achou.Count -gt 0) {
            Write-Log ('[X] {0,-14} {1}' -f $c.Chave, (($achou | Select-Object -First 4) -join ' · ')) 'DADO'
            Add-ItemInv 'Clinico' $c.Chave 'presente' (($achou | Select-Object -First 6) -join ' | ')
        } else {
            Write-Log ('[ ] {0}' -f $c.Chave) 'DADO'
            Add-ItemInv 'Clinico' $c.Chave 'ausente'
        }
    }

    Write-Log '' 'DADO'
    $papel = $false
    foreach ($c in (Get-CatalogoServidor)) {
        $ps = @($procs | Where-Object { $_.ProcessName -match $c.Padrao })
        $sv = @($svcs  | Where-Object { $_.Name -match $c.Padrao -or $_.DisplayName -match $c.Padrao })
        if ($ps.Count -gt 0 -or $sv.Count -gt 0) {
            $papel = $true
            Write-Log ('[X] Papel de servidor: {0} ({1} processo(s), {2} servico(s))' -f $c.Chave, $ps.Count, $sv.Count) 'DADO'
            Add-ItemInv 'Servidor' $c.Chave 'presente' ('processos=' + $ps.Count + ' servicos=' + $sv.Count)
        }
    }
    if ($papel) {
        Add-Achado 'ALERTA' 'Esta maquina hospeda servico (banco de dados ou aplicacao)' 'Limpeza e reinicio precisam de combinado previo com quem depende do servico.' 'Servidor' 'Alto'
    } else {
        Add-ItemInv 'Servidor' 'Nenhum' 'estacao comum'
    }

    # agentes corporativos, so para registro
    $agentes = @($svcs | Where-Object { $_.State -eq 'Running' -and ($_.Name -match 'CcmExec|DSASvc|ntrtscan|tmlisten|TMBM|QualysAgent|stAgent|Cortex|cyserver|cyvera|LsAgent|Quest|OnDemand') })
    if ($agentes.Count -gt 0) {
        Write-Log ('Agentes corporativos em execucao: {0}' -f (($agentes | Select-Object -ExpandProperty Name) -join ', ')) 'DADO'
        Add-ItemInv 'Agentes' 'Em execucao' ("$($agentes.Count)") ((($agentes | Select-Object -ExpandProperty Name) -join ', '))
    }
}

function Invoke-InvInicializacaoCompleta {
    Write-Titulo 'Inventario · tudo que abre com o Windows'
    $itens = @()
    foreach ($k in @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
                     'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                     'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run')) {
        try {
            $p = Get-ItemProperty -Path $k -ErrorAction Stop
            foreach ($n in $p.PSObject.Properties.Name) {
                if ($n -like 'PS*') { continue }
                $origem = if ($k -like 'HKCU*') { 'usuario' } else { 'maquina' }
                $itens += [pscustomobject]@{ Nome = $n; Origem = $origem; Comando = "$($p.$n)" }
            }
        } catch { }
    }
    foreach ($pasta in @([Environment]::GetFolderPath('Startup'), [Environment]::GetFolderPath('CommonStartup'))) {
        try {
            foreach ($f in (Get-ChildItem -LiteralPath $pasta -File -ErrorAction Stop)) {
                $origem = if ($pasta -like "*$env:USERNAME*") { 'pasta usuario' } else { 'pasta maquina' }
                $itens += [pscustomobject]@{ Nome = $f.Name; Origem = $origem; Comando = $f.FullName }
            }
        } catch { }
    }
    Write-Log ('{0} itens (usuario + maquina)' -f $itens.Count) 'DADO'
    foreach ($i in ($itens | Sort-Object Origem, Nome)) {
        Write-Log ('[{0,-13}] {1,-36} {2}' -f $i.Origem, $i.Nome, $i.Comando) 'DADO'
        Add-ItemInv 'Inicializacao' $i.Nome $i.Origem $i.Comando
    }
    $daMaquina = @($itens | Where-Object { $_.Origem -like '*maquina*' }).Count
    if ($daMaquina -ge 8) {
        Add-Achado 'ALERTA' ('{0} itens de inicializacao sao da maquina, nao do seu usuario' -f $daMaquina) 'Esses so o TI desativa. Registre no chamado se o logon estiver longo.' 'Inicializacao' 'Medio'
    }
}

function Invoke-InvServicos {
    Write-Titulo 'Inventario · servicos em execucao'
    try {
        $svcs = @(Get-CimInstance Win32_Service -ErrorAction Stop)
        $rod = @($svcs | Where-Object { $_.State -eq 'Running' } | Sort-Object DisplayName)
        Write-Log ('{0} em execucao de {1} instalados' -f $rod.Count, $svcs.Count) 'DADO'
        foreach ($sv in $rod) {
            Write-Log ('{0,-40} {1,-12} {2}' -f $sv.Name, $sv.StartMode, $sv.DisplayName) 'DADO'
            Add-ItemInv 'Servico' $sv.Name $sv.StartMode "$($sv.DisplayName)"
        }
    } catch { }
}

function Invoke-InvOfficeExcel {
    Write-Titulo 'Inventario · Office e Excel'
    $ver = Get-ValorReg 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' 'VersionToReport'
    if ($ver) { Write-Log ('Versao do Office: {0}' -f $ver) 'DADO'; Add-ItemInv 'Office' 'Versao' "$ver" }

    $locais = @()
    try {
        foreach ($k in (Get-ChildItem -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security\Trusted Locations' -ErrorAction Stop)) {
            $cam = (Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue).Path
            if ($cam) { $locais += $cam; Write-Log ('Local confiavel: {0}' -f $cam) 'DADO'; Add-ItemInv 'Excel' 'Local confiavel' $cam }
        }
    } catch { }
    $temLocal = @($locais | Where-Object { $_ -like '*PRESCRI*' }).Count
    if ($temLocal -eq 0) {
        Add-Achado 'ALERTA' 'A pasta clinica nao esta nos locais confiaveis do Excel' 'Rode o Modulo 1 (Preparar ambiente) nesta maquina.' 'Excel' 'Alto'
    }

    foreach ($k in @('HKCU:\Software\Microsoft\Office\Excel\Addins', 'HKLM:\Software\Microsoft\Office\Excel\Addins')) {
        try {
            foreach ($a in (Get-ChildItem -Path $k -ErrorAction Stop)) {
                $lb = (Get-ItemProperty -Path $a.PSPath -ErrorAction SilentlyContinue).LoadBehavior
                Write-Log ('Suplemento: {0} (LoadBehavior {1})' -f $a.PSChildName, $lb) 'DADO'
                Add-ItemInv 'Excel' 'Suplemento' $a.PSChildName "LoadBehavior=$lb"
                if ($lb -eq 3 -and $a.PSChildName -match 'PDFMaker|Acrobat') {
                    Add-Achado 'ALERTA' ('Suplemento {0} carrega junto com o Excel' -f $a.PSChildName) 'Arquivo > Opcoes > Suplementos > Suplementos de COM > desmarque. Atrasa a abertura de toda planilha.' 'Excel' 'Medio'
                }
            }
        } catch { }
    }
}

function Invoke-InvPastaClinica {
    Write-Titulo 'Inventario · pasta clinica'
    if (-not (Test-Path -LiteralPath $script:PastaClinica)) {
        Add-Achado 'ALERTA' ('{0} nao existe' -f $script:PastaClinica) 'Rode o Modulo 1 (Preparar ambiente).' 'Ambiente' 'Alto'
        return
    }
    $tot = Get-TamanhoPasta -Caminho $script:PastaClinica -TimeoutSeg 120
    Write-Log ('Total: {0}' -f (Format-Bytes $tot)) 'DADO'
    Add-ItemInv 'Pasta clinica' 'Total' (Format-Bytes $tot)

    try {
        foreach ($d in (Get-ChildItem -LiteralPath $script:PastaClinica -Directory -ErrorAction Stop)) {
            if ($script:Cancelar) { break }
            $b = Get-TamanhoPasta -Caminho $d.FullName -TimeoutSeg 45
            Write-Log ('{0,-40} {1,10}' -f $d.Name, (Format-Bytes $b)) 'DADO'
            Add-ItemInv 'Pasta clinica' $d.Name (Format-Bytes $b)
        }
        foreach ($f in (Get-ChildItem -LiteralPath $script:PastaClinica -File -Filter '*.xls*' -ErrorAction Stop)) {
            Write-Log ('{0,-40} {1,10}  {2:dd/MM/yyyy HH:mm}' -f $f.Name, (Format-Bytes $f.Length), $f.LastWriteTime) 'DADO'
            Add-ItemInv 'Portal' $f.Name (Format-Bytes $f.Length) ('{0:yyyy-MM-dd HH:mm}' -f $f.LastWriteTime)
        }
    } catch { }

    $fonte = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts\3OF9_NEW.TTF'
    $pdftk = Join-Path $script:PastaClinica 'UTILITARIOS\PDFtk\pdftk.exe'
    $temFonte = Test-Path -LiteralPath $fonte
    $temPdftk = Test-Path -LiteralPath $pdftk
    Add-ItemInv 'Ambiente' 'Fonte 3OF9' $(if ($temFonte) { 'instalada' } else { 'ausente' })
    Add-ItemInv 'Ambiente' 'PDFtk' $(if ($temPdftk) { 'instalado' } else { 'ausente' })
    if (-not $temFonte -or -not $temPdftk) {
        $falta = @()
        if (-not $temFonte) { $falta += 'fonte de codigo de barras' }
        if (-not $temPdftk) { $falta += 'PDFtk' }
        Add-Achado 'ALERTA' ('Preparacao incompleta: falta {0}' -f ($falta -join ' e ')) 'Rode o Modulo 1 (Preparar ambiente) nesta maquina.' 'Ambiente' 'Alto'
    }
}

function Invoke-InvRedeImpressoras {
    Write-Titulo 'Inventario · rede e impressoras'
    try {
        foreach ($n in (Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' })) {
            $ip = (Get-NetIPAddress -InterfaceIndex $n.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress -join ', '
            Write-Log ('{0,-32} {1,-12} {2}' -f $n.InterfaceDescription, $n.LinkSpeed, $ip) 'DADO'
            Add-ItemInv 'Rede' $n.Name $n.LinkSpeed $ip
            if ($n.InterfaceDescription -match 'Wi-?Fi|Wireless|802\.11' -or $n.Name -match 'Wi-?Fi') {
                Add-Achado 'ALERTA' ('Estacao conectada por Wi-Fi ({0})' -f $n.LinkSpeed) 'Os portais tem 40 MB e sao abertos da rede. Cabo reduz de forma sensivel o tempo de abertura e evita corrupcao do arquivo.' 'Rede' 'Alto'
            } elseif ($n.LinkSpeed -match '^(10|100) Mbps') {
                Add-Achado 'ALERTA' ('Placa de rede negociando apenas {0}' -f $n.LinkSpeed) 'A placa e o switch sao de 1 Gbps: 100 Mbps quase sempre e cabo velho, mal crimpado ou porta ruim. Trocar o cabo custa nada e melhora Tasy, Citrix e a abertura dos portais de 40 MB.' 'Rede' 'Alto'
            }
        }
    } catch { }
    try {
        $sufixo = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
        Write-Log ('Sufixo DNS: {0}' -f $sufixo) 'DADO'
        Add-ItemInv 'Rede' 'Sufixo DNS' "$sufixo"
    } catch { }
    try {
        foreach ($imp in (Get-CimInstance Win32_Printer -ErrorAction Stop)) {
            $marca = if ($imp.Default) { '[padrao] ' } else { '' }
            Write-Log ('{0}{1,-42} {2}' -f $marca, $imp.Name, $imp.PortName) 'DADO'
            Add-ItemInv 'Impressora' $imp.Name $(if ($imp.Default) { 'padrao' } else { '' }) "$($imp.PortName)"
            if ($imp.Default -and ($imp.WorkOffline -or $imp.PrinterStatus -eq 7)) {
                Add-Achado 'ALERTA' ('Impressora padrao offline: {0}' -f $imp.Name) 'O Excel e o Word congelam ao abrir arquivo quando a impressora padrao nao responde. Troque a padrao para "Microsoft Print to PDF" enquanto nao resolvem.' 'Impressora' 'Alto'
            }
        }
    } catch { }
}

# ---- residuos de migracao de dominio ------------------------------
function Get-CredenciaisOrfas {
    $orfas = @()
    try {
        $saida = & cmdkey /list 2>$null
        foreach ($l in $saida) {
            if ("$l" -notmatch 'Destino:|Target:') { continue }
            $alvo = ("$l" -replace '.*(Destino|Target):\s*', '').Trim()
            if (-not $alvo) { continue }
            if ($alvo -match '\*|MicrosoftAccount|WindowsLive|virtualapp|SSO_POP|LegacyGeneric:target=Microsoft') { continue }
            $servidor = ($alvo -replace '^[A-Za-z]+:target=', '') -replace '\\.*$', ''
            $servidor = $servidor.Split('/')[0]
            if (-not $servidor -or $servidor -match '^\s*$') { continue }
            if ($servidor -match '^\d{1,3}(\.\d{1,3}){3}$') { continue }
            $resolve = $false
            $tentativas = @($servidor)
            foreach ($sf in (Get-SufixosDns)) { $tentativas += ($servidor + '.' + $sf) }
            foreach ($tent in $tentativas) {
                try { [void][System.Net.Dns]::GetHostAddresses($tent); $resolve = $true; break } catch { }
            }
            if (-not $resolve) { $orfas += [pscustomobject]@{ Alvo = $alvo; Servidor = $servidor } }
        }
    } catch { }
    return $orfas
}

function Invoke-InvMigracao {
    Write-Titulo 'Inventario · residuos de migracao de dominio'

    $prog = Get-ProgramasInstalados
    $ag = @($prog | Where-Object { $_.Nome -match 'On Demand Migration|Quest.*Migration|Binary Tree|ADMT' })
    if ($ag.Count -gt 0) {
        foreach ($a in $ag) { Write-Log ('Agente de migracao instalado: {0} {1}' -f $a.Nome, $a.Versao) 'DADO'; Add-ItemInv 'Migracao' 'Agente' $a.Nome $a.Versao }
        Add-Achado 'ALERTA' ('Agente de migracao de dominio ainda instalado ({0})' -f $ag[0].Nome) 'A migracao terminou ha anos. Remover exige administrador: peca ao TI a desinstalacao em massa. Ele roda servico e varre o perfil a cada logon.' 'Migracao' 'Alto'
    } else {
        Write-Log 'Nenhum agente de migracao instalado.' 'DADO'
    }

    # pasta do perfil com nome de outra conta: heranca classica de migracao
    $pastaPerfil = Split-Path $env:USERPROFILE -Leaf
    if ($pastaPerfil -and ($pastaPerfil -ne $env:USERNAME)) {
        Write-Log ('Usuario logado .....: {0}' -f $env:USERNAME) 'DADO'
        Write-Log ('Pasta do perfil ....: {0}' -f $env:USERPROFILE) 'DADO'
        Add-ItemInv 'Migracao' 'Perfil com nome de outra conta' $pastaPerfil $env:USERNAME
        Add-Achado 'ALERTA' ('A pasta do perfil nao tem o nome do usuario: {0} usa a pasta {1}' -f $env:USERNAME, $pastaPerfil) 'Sinal de perfil herdado de outra conta, tipico de migracao de dominio. Funciona, mas confunde script, politica de grupo e permissao. Se esta maquina der problema estranho de perfil, e por aqui que se comeca.' 'Migracao' 'Medio'
    }

    if ($env:USERPROFILE -match '\.[A-Za-z0-9]+$') {
        Write-Log ('Perfil com sufixo de dominio: {0}' -f $env:USERPROFILE) 'DADO'
        Add-ItemInv 'Migracao' 'Perfil com sufixo' $env:USERPROFILE
        Write-Log 'Sinal de perfil recriado na migracao. Recriar de novo exige administrador.' 'DADO'
    }

    $orfas = Get-CredenciaisOrfas
    if ($orfas.Count -gt 0) {
        Write-Log ('{0} credencial(is) salva(s) para servidores que nao existem mais:' -f $orfas.Count) 'DADO'
        foreach ($o in $orfas) { Write-Log ('   ' + $o.Alvo) 'DADO'; Add-ItemInv 'Migracao' 'Credencial orfa' $o.Servidor $o.Alvo }
        Add-Achado 'ALERTA' ('{0} credencial(is) salva(s) apontando para servidor inexistente' -f $orfas.Count) 'O Windows tenta usar cada uma ao abrir o Explorer e espera o tempo limite. O Modulo 3 remove - o item vem DESMARCADO porque nao tem desfazer.' 'Migracao' 'Alto'
    } else {
        Write-Log 'Nenhuma credencial orfa no Gerenciador de Credenciais.' 'DADO'
    }
}

# ---- gravacao do inventario ---------------------------------------
function Save-Inventario {
    Write-Titulo 'Relatorio'
    $carimbo = Get-Date -Format 'yyyyMMdd_HHmm'
    $base = 'Inventario_{0}_{1}' -f $env:COMPUTERNAME, $carimbo
    $pasta = $script:PastaRelatorios

    # sem fallback: se a pasta de logs nao estiver acessivel, nao grava nada
    $acessivel = $false
    try {
        if ($pasta -like '\\*') {
            $srv = ($pasta.TrimStart('\') -split '\\')[0]
            $t = Test-PortaTcp -Alvo $srv -Porta 445 -TimeoutMs 1500
            if (-not $t.Ok) {
                Write-Log ('Servidor de logs fora de alcance: {0}' -f $srv) 'DADO'
                Write-Log 'A planilha nao foi gravada. O relatorio e este log ("Copiar log" ou "Salvar log").' 'DADO'
                return
            }
        }
        if (Test-Path -LiteralPath $pasta) {
            $acessivel = $true
        } else {
            $pai = Split-Path -Path $pasta -Parent
            if ($pai -and (Test-Path -LiteralPath $pai)) {
                New-Item -ItemType Directory -Path $pasta -Force | Out-Null
                $acessivel = (Test-Path -LiteralPath $pasta)
            }
        }
    } catch { $acessivel = $false }

    if (-not $acessivel) {
        Write-Log ('Pasta de logs indisponivel: {0}' -f $pasta) 'DADO'
        Write-Log 'A planilha nao foi gravada. O relatorio e este log ("Copiar log" ou "Salvar log").' 'DADO'
        return
    }
    $csv = Join-Path $pasta ($base + '.csv')
    try {
        $script:ItensInv | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Force
        Write-Log ('Planilha ..: {0}' -f $csv) 'OK'
        Write-Log 'Gravada na pasta de LOGs da rede. O relatorio de texto e este proprio log.' 'DADO'
        Write-Log 'Junte os .csv de varias maquinas numa planilha so para comparar o parque.' 'ACAO'
    } catch {
        Write-Log ('Nao foi possivel gravar o relatorio: {0}' -f $_.Exception.Message) 'ALERTA'
    }
}

function Get-VeredictoArquivo {
    param($Arquivo)
    $p = "$($Arquivo.FullName)"
    $e = "$($Arquivo.Extension)".ToLower()

    # ---- MANTER: arquivo de sistema, dado clinico, banco, programa ----
    if ($Arquivo.Name -match '^(hiberfil|pagefile|swapfile)\.sys$')      { return @{ V = 'MANTER';  M = 'Arquivo de sistema do Windows. Apagar quebra a maquina.' } }
    if ($e -match '^\.(mdf|ldf|ndf|bak|trn|dbf)$')                       { return @{ V = 'MANTER';  M = 'Arquivo de banco de dados. Apagar derruba o sistema.' } }
    if (Test-DentroDaPasta -Caminho $p -Pasta $script:PastaClinica)      { return @{ V = 'MANTER';  M = 'Esta na pasta clinica.' } }
    if ($p -match $script:RaizesClinicas) { return @{ V = 'MANTER'; M = 'Dado de sistema clinico.' } }
    if ($p -match '\\CentBrowser')      { return @{ V = 'MANTER'; M = 'CentBrowser e o navegador do prontuario da clinica.' } }
    if ($p -match 'Digitalcore|\\Onis') { return @{ V = 'MANTER'; M = 'Dados do visualizador DICOM Onis.' } }
    if ($p -match '\\Citrix\\')         { return @{ V = 'MANTER'; M = 'Cache do Citrix. Este kit nunca o toca.' } }
    if ($e -match '^\.(dcm|dicom|nii|nrrd|mha|raw)$')                    { return @{ V = 'MANTER';  M = 'Imagem medica. Mova para a rede ou PACS, nao apague.' } }
    if ($e -match '^\.(ost|pst)$')                                       { return @{ V = 'MANTER';  M = 'Caixa de e-mail. Reduza pelo Outlook, nunca apague.' } }
    if ($p -match '\\Program Files|\\Windows\\|\\ProgramData\\|\\Microsoft SQL Server\\|\\Tomcat') { return @{ V = 'MANTER'; M = 'Parte de um programa instalado.' } }
    if ($e -match '^\.(exe|dll|msi|sys)$' -and $p -match '\\AppData\\Local\\' -and $p -notmatch '\\Temp\\|\\Downloads\\|[Cc]ache') { return @{ V = 'MANTER'; M = 'Programa instalado dentro do seu perfil.' } }
    if ($e -match '^\.(xlsb|xlsm)$')                                     { return @{ V = 'MANTER';  M = 'Planilha de trabalho.' } }

    # ---- SEGURO: descartavel ----
    if ($e -match '^\.(dmp|mdmp|hdmp)$')                                 { return @{ V = 'SEGURO';  M = 'Despejo de travamento. Nao tem uso.' } }
    if ($e -match '^\.(iso|img|vhd|vhdx|wim|esd)$')                      { return @{ V = 'SEGURO';  M = 'Imagem de instalacao, serve so uma vez.' } }
    if ($e -match '^\.(msi|msp|exe)$' -and $p -match '\\Downloads\\|\\Temp\\|\\Temporar|C:\\Temp') { return @{ V = 'SEGURO'; M = 'Instalador ja usado.' } }
    if ($p -match 'component_crx_cache|GrShaderCache|ShaderCache|ProvenanceData|EBWebView') { return @{ V = 'SEGURO'; M = 'Cache interno do navegador. E recriado sozinho.' } }
    if ($p -match '\\Temp\\|\\Temporar|[Cc]ache|\\CrashDumps\\|\\WER\\|\\INetCache\\|\\Downloaded Installations\\') { return @{ V = 'SEGURO'; M = 'Esta em pasta de cache ou temporaria.' } }
    if ($e -match '^\.(log|etl|old|tmp|chk|gid)$')                       { return @{ V = 'SEGURO';  M = 'Log ou sobra de instalacao.' } }

    # ---- resto: decisao do usuario ----
    if ($e -match '^\.(zip|rar|7z|cab)$')                                { return @{ V = 'REVISAR'; M = 'Compactado. Se ja foi extraido, pode apagar.' } }
    if ($e -match '^\.(mp4|avi|mkv|mov|wmv|mp3|wav)$')                   { return @{ V = 'REVISAR'; M = 'Midia. Mova para a rede se for do servico.' } }
    if ($e -match '^\.(pdf|docx|doc|xlsx|pptx|accdb)$')                  { return @{ V = 'REVISAR'; M = 'Documento. Confirme se ja existe copia na rede.' } }
    return @{ V = 'REVISAR'; M = 'Nao reconhecido. Confira antes de apagar.' }
}


function Test-EnderecoInterno {
    param([string]$Servidor)
    if ($Servidor -match '^10\.|^192\.168\.|^172\.(1[6-9]|2\d|3[01])\.') { return $true }
    try {
        $ips = @([System.Net.Dns]::GetHostAddresses($Servidor) |
                 Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                 Select-Object -ExpandProperty IPAddressToString)
        foreach ($ip in $ips) {
            if ($ip -match '^10\.|^192\.168\.|^172\.(1[6-9]|2\d|3[01])\.') { return $true }
        }
        return $false
    } catch {
        # nao resolveu: so trata como interno se o nome pertencer ao dominio da rede
        $sufixo = ''
        try { $sufixo = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName } catch { }
        if (-not $sufixo) { $sufixo = "$env:USERDNSDOMAIN" }
        if ($Servidor -notmatch '\.') { return $true }
        if ($sufixo -and $Servidor.ToLower().EndsWith($sufixo.ToLower())) { return $true }
        return $false
    }
}

function Measure-RespostaHttp {
    param([string]$Url, [string]$Alvo, [int]$Porta = 0, [int]$Amostras = 4)
    if ($Porta -le 0) { try { $Porta = ([uri]$Url).Port } catch { $Porta = 80 } }

    $tcp = @()
    $http = @()
    $codigo = 0
    $urlFinal = $Url
    $certRuim = $false
    for ($i = 0; $i -lt $Amostras; $i++) {
        if ($script:Cancelar) { break }
        $t = Test-PortaTcp -Alvo $Alvo -Porta $Porta -TimeoutMs 2500
        if ($t.Ok) { $tcp += $t.Ms }
        $r = Test-UrlHttp -Url $Url -TimeoutSeg 12
        if ($r.Codigo -gt 0) { $http += $r.Ms; $codigo = $r.Codigo }
        $urlFinal = $r.UrlFinal
        if ($r.CertificadoInvalido) { $certRuim = $true }
        Pump
        Start-Sleep -Milliseconds 250
    }

    # a mediana ignora o pico unico da primeira conexao, que a media inflava
    $tcpMed = -1; $tcpMin = -1; $tcpMax = -1
    if ($tcp.Count -gt 0) {
        $ord = @($tcp | Sort-Object)
        $tcpMed = [int]$ord[[Math]::Floor(($ord.Count - 1) / 2)]
        $tcpMin = ($tcp | Measure-Object -Minimum).Minimum
        $tcpMax = ($tcp | Measure-Object -Maximum).Maximum
    }
    $htMed = -1; $htMin = -1; $htMax = -1
    if ($http.Count -gt 0) {
        $ord = @($http | Sort-Object)
        $htMed = [int]$ord[[Math]::Floor(($ord.Count - 1) / 2)]
        $htMin = ($http | Measure-Object -Minimum).Minimum
        $htMax = ($http | Measure-Object -Maximum).Maximum
    }

    return [pscustomobject]@{
        Codigo = $codigo; Amostras = $http.Count
        TcpMin = $tcpMin; TcpMed = $tcpMed; TcpMax = $tcpMax
        HttpMin = $htMin; HttpMed = $htMed; HttpMax = $htMax
        UrlFinal = $urlFinal; CertificadoInvalido = $certRuim
        Processamento = $(if ($htMed -ge 0 -and $tcpMed -ge 0) { [Math]::Max(0, $htMed - $tcpMed) } else { -1 })
        Instavel = $(if ($htMax -ge 0 -and $htMin -ge 0) { (($htMax - $htMin) -gt 2000) } else { $false })
    }
}

# =====================================================================
# 4E. TASY - PRONTUARIO ELETRONICO
# =====================================================================
function Get-TasyProcessos {
    $nomes = @('tasy', 'TasyAgent', 'tasy-agentw', 'TasyAgentTray', 'javaw', 'java', 'jp2launcher', 'CentBrowser')
    $ps = @(Get-Process -Name $nomes -ErrorAction SilentlyContinue)
    [pscustomobject]@{
        Rodando = ($ps.Count -gt 0)
        Qtd     = $ps.Count
        Memoria = $(if ($ps.Count -gt 0) { ($ps | Measure-Object WorkingSet64 -Sum).Sum } else { 0 })
        Nomes   = @($ps | Select-Object -ExpandProperty ProcessName -Unique)
    }
}

function Get-TasyUrlsDescobertas {
    $achadas = @()

    # 1. atalhos .url e .lnk da area de trabalho e do menu iniciar
    $pastas = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('CommonDesktopDirectory'),
                [Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('CommonPrograms'))
    foreach ($pasta in $pastas) {
        if (-not (Test-Path -LiteralPath $pasta)) { continue }
        try {
            foreach ($f in (Get-ChildItem -LiteralPath $pasta -Filter '*.url' -Recurse -ErrorAction SilentlyContinue)) {
                $txt = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction SilentlyContinue
                if ($txt -match '(?im)^URL\s*=\s*(\S+)') {
                    $u = $Matches[1]
                    if ($u -match 'Tasy|Wheb|prontuario') { $achadas += $u }
                }
            }
        } catch { }
        try {
            $sh = New-Object -ComObject WScript.Shell
            foreach ($f in (Get-ChildItem -LiteralPath $pasta -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue)) {
                if ($f.Name -notmatch 'Tasy|Wheb|Prontu') { continue }
                $lnk = $sh.CreateShortcut($f.FullName)
                $alvo = "$($lnk.TargetPath) $($lnk.Arguments)"
                foreach ($m in [regex]::Matches($alvo, 'https?://[^\s"'']+')) { $achadas += $m.Value }
            }
        } catch { }
    }

    # 2. favoritos dos navegadores (so as entradas que citam Tasy)
    $bases = @(
        (Join-Path $env:LOCALAPPDATA 'CentBrowser\User Data'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data'),
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data')
    )
    foreach ($b in $bases) {
        if (-not (Test-Path -LiteralPath $b)) { continue }
        try {
            foreach ($fav in (Get-ChildItem -LiteralPath $b -Filter 'Bookmarks' -Recurse -ErrorAction SilentlyContinue)) {
                $txt = Get-Content -LiteralPath $fav.FullName -Raw -ErrorAction SilentlyContinue
                foreach ($m in [regex]::Matches("$txt", '"url":\s*"(https?://[^"]*[Tt]asy[^"]*)"')) {
                    $achadas += ($m.Groups[1].Value -replace '\\/', '/')
                }
            }
        } catch { }
    }

    # normaliza para a raiz do servidor
    $raizes = @()
    foreach ($u in $achadas) {
        try {
            $uri = [uri]$u
            $caminho = $uri.AbsolutePath
            if ($caminho -match '(?i)(/TasyAppServer)') { $caminho = $Matches[1] + '/' } else { $caminho = '/' }
            $raizes += ('{0}://{1}{2}{3}' -f $uri.Scheme, $uri.Host, $(if ($uri.IsDefaultPort) { '' } else { ':' + $uri.Port }), $caminho)
        } catch { }
    }
    return @($raizes | Select-Object -Unique)
}



function Get-TasyUrlsEfetivas {
    if ($script:TasyUrls -and $script:TasyUrls.Count -gt 0) { return @($script:TasyUrls) }
    if ($null -eq $script:TasyDescobertas) {
        $script:TasyDescobertas = @(Get-TasyUrlsDescobertas | ForEach-Object {
            $srv = try { ([uri]$_).Host } catch { $_ }
            [pscustomobject]@{ Nome = ('Descoberto: ' + $srv); Url = $_ }
        })
    }
    return @($script:TasyDescobertas)
}

function Get-CacheJava {
    $c = @(
        (Join-Path $env:LOCALAPPDATA 'Sun\Java\Deployment\cache'),
        (Join-Path $env:APPDATA 'Sun\Java\Deployment\cache'),
        (Join-Path $env:USERPROFILE 'AppData\LocalLow\Sun\Java\Deployment\cache')
    )
    return @($c | Where-Object { Test-Path -LiteralPath $_ })
}

function Invoke-DiagTasy {
    Write-Titulo 'Diagnostico do Tasy (prontuario eletronico)'

    # =============== 1. cliente na estacao ===============
    $prog = Get-ProgramasInstalados
    $tasyProg = @($prog | Where-Object { $_.Nome -match 'Tasy|Philips.*Tasy|Wheb' })
    foreach ($t in $tasyProg) { Write-Log ('Instalado: {0} {1}' -f $t.Nome, $t.Versao) 'DADO'; Add-ItemInv 'Tasy' 'Cliente' $t.Nome $t.Versao }
    if ($tasyProg.Count -eq 0) { Write-Log 'Sem cliente instalado: o Tasy roda no navegador (Wheb HTML5).' 'DADO' }

    # navegador do prontuario: e ele quem carrega a tela do Tasy
    $navProcs = @(Get-Process -Name 'CentBrowser', 'msedge', 'chrome' -ErrorAction SilentlyContinue)
    if ($navProcs.Count -gt 0) {
        $porNav = $navProcs | Group-Object ProcessName | ForEach-Object {
            [pscustomobject]@{ Nome = $_.Name; Qtd = $_.Count; Mem = ($_.Group | Measure-Object WorkingSet64 -Sum).Sum }
        } | Sort-Object Mem -Descending
        foreach ($n in $porNav) {
            Write-Log ('Navegador {0,-14} {1,10} em {2} processo(s)' -f $n.Nome, (Format-Bytes $n.Mem), $n.Qtd) 'DADO'
            Add-ItemInv 'Tasy' ('Navegador ' + $n.Nome) (Format-Bytes $n.Mem) ("$($n.Qtd) processos")
        }
        $pesado = $porNav | Select-Object -First 1
        if ($pesado -and $pesado.Mem -gt 2.5GB) {
            Add-Achado 'ALERTA' ('{0} consumindo {1} em {2} processos' -f $pesado.Nome, (Format-Bytes $pesado.Mem), $pesado.Qtd) 'O Tasy roda dentro do navegador: com essa carga a tela trava mesmo com o servidor respondendo rapido. Feche as abas que nao usa antes de abrir chamado.' 'Tasy' 'Alto'
        }
    }

    # cliente Java (Tasy antigo / relatorios)
    try {
        $javas = @(Get-CimInstance Win32_Process -Filter "Name='javaw.exe' OR Name='java.exe'" -ErrorAction SilentlyContinue)
        foreach ($j in $javas) {
            $xmx = ''
            if ("$($j.CommandLine)" -match '-Xmx(\d+)([mMgG])') { $xmx = ('limite de memoria ' + $Matches[1] + $Matches[2].ToUpper()) }
            $bits = $(if ("$($j.ExecutablePath)" -match 'Program Files \(x86\)') { '32 bits' } else { '64 bits' })
            Write-Log ('Java em execucao: {0} · {1} · {2}' -f (Format-Bytes $j.WorkingSetSize), $bits, $xmx) 'DADO'
            if ($bits -eq '32 bits') {
                Add-Achado 'ALERTA' 'Cliente Java do Tasy rodando em 32 bits' 'Java 32 bits nao passa de ~1,5 GB de memoria. Relatorio grande trava ou fecha sozinho. Peca ao TI o Java 64 bits.' 'Tasy' 'Medio'
            }
        }
    } catch { }

    # =============== 2. cada destino ===============
    $urls = Get-TasyUrlsEfetivas
    if ($urls.Count -eq 0) {
        Add-Achado 'ALERTA' 'Nenhum endereco do Tasy configurado' 'Preencha $script:TasyUrls na secao 1 do aplicativo.' 'Tasy' 'Medio'
        return
    }

    $resumo = @()
    foreach ($d in $urls) {
        if ($script:Cancelar) { return }
        $i = Get-InfoUrl -Url $d.Url
        if (-not $i) { continue }

        Write-Log '' 'DADO'
        Write-Log ('--- {0} ---' -f $d.Nome) 'DADO'
        Write-Log ('Endereco ..: {0}' -f $d.Url) 'DADO'
        Set-Status ('Medindo ' + $d.Nome + '...')
        Add-ItemInv 'Tasy' $d.Nome $d.Url

        $interno = Test-EnderecoInterno -Servidor $i.Servidor
        Write-Log ('Tipo ......: {0}' -f $(if ($interno) { 'servidor interno da rede' } else { 'servico externo, sai pela internet' })) 'DADO'

        # DNS
        $alvo = $i.Servidor
        if (-not $i.EhIp) {
            try {
                $rel = [System.Diagnostics.Stopwatch]::StartNew()
                $ips = [System.Net.Dns]::GetHostAddresses($i.Servidor) | Select-Object -ExpandProperty IPAddressToString
                $rel.Stop()
                Write-Log ('DNS .......: {0} ({1} ms)' -f ($ips -join ', '), $rel.ElapsedMilliseconds) 'DADO'
            } catch {
                Add-Achado 'CRITICO' ('{0}: o nome {1} nao resolve nesta rede' -f $d.Nome, $i.Servidor) 'A estacao nao esta enxergando o DNS. E problema de rede da estacao, nao do prontuario.' 'Tasy' 'Alto'
                $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = -1; Tcp = -1; Situacao = 'DNS nao resolve' }
                continue
            }
        }

        # portas
        $porta = $i.Porta
        $abertas = @()
        foreach ($p in @($porta, 80, 443, 8080, 28080 | Select-Object -Unique)) {
            $t = Test-PortaTcp -Alvo $alvo -Porta $p -TimeoutMs 1200
            if ($t.Ok) { $abertas += ('{0} ({1} ms)' -f $p, $t.Ms) }
        }
        Write-Log ('Portas ....: {0}' -f $(if ($abertas.Count -gt 0) { $abertas -join ' · ' } else { 'nenhuma respondeu' })) 'DADO'
        if ($abertas.Count -eq 0) {
            Add-Achado 'CRITICO' ('{0}: servidor {1} sem resposta' -f $d.Nome, $alvo) 'Se as outras estacoes tambem nao abrem, o servidor caiu: chamado imediato.' 'Tasy' 'Alto'
            $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = -1; Tcp = -1; Situacao = 'servidor sem resposta' }
            continue
        }

        # ---- medicao de desempenho: separa rede de servidor ----
        Write-Log 'Medindo o tempo de resposta (4 amostras)...' 'DADO'
        $m = Measure-RespostaHttp -Url $d.Url -Alvo $alvo -Porta $porta -Amostras 4

        if ($m.Amostras -eq 0) {
            Add-Achado 'CRITICO' ('{0}: portas abertas mas a pagina nao responde' -f $d.Nome) 'O servidor web esta no ar e a aplicacao nao. Chamado com o horario exato.' 'Tasy' 'Alto'
            $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = -1; Tcp = $m.TcpMed; Situacao = 'pagina nao responde' }
            continue
        }

        Write-Log ('Rede (TCP) ....: min {0} · tipico {1} · max {2} ms' -f $m.TcpMin, $m.TcpMed, $m.TcpMax) 'DADO'
        Write-Log ('Pagina (HTTP) .: min {0} · tipico {1} · max {2} ms   (HTTP {3})' -f $m.HttpMin, $m.HttpMed, $m.HttpMax, $m.Codigo) 'DADO'
        Write-Log ('Processamento .: ~{0} ms no servidor (HTTP menos rede)' -f $m.Processamento) 'DADO'
        Add-ItemInv 'Tasy' ($d.Nome + ' - resposta') ("$($m.HttpMed) ms") ("tcp=$($m.TcpMed)ms http=$($m.HttpMed)ms")
        if ($m.TcpMax -gt 0 -and $m.TcpMed -gt 0 -and $m.TcpMax -gt ($m.TcpMed * 4) -and $m.TcpMax -gt 300) {
            Write-Log ('Oscilacao: uma das amostras levou {0} ms contra {1} ms das outras. Rede instavel, nao lenta.' -f $m.TcpMax, $m.TcpMed) 'DADO'
        }

        $situacao = 'ok'
        if ($m.HttpMed -gt 3000) {
            $situacao = 'servidor lento'
            Add-Achado 'CRITICO' ('{0}: pagina levando {1} ms para responder' -f $d.Nome, $m.HttpMed) ('A rede responde em {0} ms, entao a demora e do servidor de aplicacao. Anexe este log no chamado: nao adianta mexer na estacao.' -f $m.TcpMed) 'Tasy' 'Alto'
        } elseif ($m.HttpMed -gt 1500) {
            $situacao = 'servidor pesado'
            Add-Achado 'ALERTA' ('{0}: pagina levando {1} ms para responder' -f $d.Nome, $m.HttpMed) ('Rede em {0} ms. A lentidao esta no servidor, nao na estacao.' -f $m.TcpMed) 'Tasy' 'Alto'
        } else {
            Add-Achado 'OK' ('{0}: respondendo em {1} ms' -f $d.Nome, $m.HttpMed) '' 'Tasy'
        }

        if ($m.TcpMed -gt 60 -and $interno) {
            Add-Achado 'ALERTA' ('{0}: latencia tipica de {1} ms para um servidor interno' -f $d.Nome, $m.TcpMed) 'Acima de 60 ms dentro da rede interna e caminho ruim. Confira a velocidade negociada da placa e a porta do switch antes de culpar o servidor.' 'Tasy' 'Medio'
        }
        if ($m.Instavel) {
            $situacao = 'instavel'
            Add-Achado 'ALERTA' ('{0}: resposta instavel (variou de {1} a {2} ms)' -f $d.Nome, $m.HttpMin, $m.HttpMax) 'Oscilacao desse tamanho e o que faz a tela do Tasy "travar" de vez em quando. Registre no chamado com o horario.' 'Tasy' 'Alto'
        }


        # HTTP x HTTPS: o atalho pode estar desatualizado
        if (-not $i.Https) {
            if ("$($m.UrlFinal)" -match '^https://') {
                Write-Log 'Seguranca .: o servidor redireciona sozinho para HTTPS. O atalho pode ficar como esta.' 'DADO'
            } else {
                $alt = Test-HttpsAlternativo -Url $d.Url
                if ($alt -and $alt.Disponivel) {
                    Add-Achado 'ALERTA' ('{0}: o atalho usa HTTP, mas o servidor tambem atende em HTTPS' -f $d.Nome) ('Troque o atalho para {0} . Em HTTP o navegador bloqueia parte do login unico e o trafego vai sem criptografia.' -f $alt.Url) 'Tasy' 'Medio'
                } else {
                    Write-Log 'Seguranca .: endereco so em HTTP. O trafego vai sem criptografia.' 'DADO'
                }
            }
        }
        if ($m.CertificadoInvalido) {
            Add-Achado 'ALERTA' ('{0}: certificado HTTPS nao confiavel nesta estacao' -f $d.Nome) 'O navegador mostra aviso de site nao seguro e pode bloquear download e impressao. Peca ao TI a instalacao do certificado da autoridade interna.' 'Tasy' 'Alto'
        }

        # O Tasy Wheb usa login por formulario proprio: a zona de Intranet do Windows
        # nao interfere. Fica so o registro, sem virar pendencia.
        if ($interno -and -not (Test-ZonaIntranet -Servidor $i.Servidor)) {
            Write-Log 'Zona ......: fora da Intranet do Windows - sem efeito aqui, o Tasy tem login proprio.' 'DADO'
        }

        $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = $m.HttpMed; Tcp = $m.TcpMed; Situacao = $situacao }
    }

    # =============== 3. comparativo ===============
    if ($resumo.Count -gt 1) {
        Write-Log '' 'DADO'
        Write-Log 'COMPARATIVO ENTRE OS AMBIENTES' 'TITULO'
        Write-Log ('{0,-34} {1,10} {2,10}  {3}' -f 'Ambiente', 'rede', 'pagina', 'situacao') 'DADO'
        foreach ($r in ($resumo | Sort-Object Http -Descending)) {
            $t = $(if ($r.Tcp -lt 0) { '-' } else { "$($r.Tcp) ms" })
            $h = $(if ($r.Http -lt 0) { '-' } else { "$($r.Http) ms" })
            Write-Log ('{0,-34} {1,10} {2,10}  {3}' -f $r.Nome, $t, $h, $r.Situacao) 'DADO'
        }
        $bons = @($resumo | Where-Object { $_.Http -ge 0 -and $_.Http -le 1500 })
        $ruins = @($resumo | Where-Object { $_.Http -gt 1500 })
        if ($bons.Count -gt 0 -and $ruins.Count -gt 0) {
            Write-Log '' 'DADO'
            Write-Log ('{0} responde rapido e {1} nao, da mesma estacao e no mesmo minuto.' -f $bons[0].Nome, $ruins[0].Nome) 'ACAO'
            Write-Log 'Isso descarta a estacao e a rede local: o problema esta no servidor lento.' 'ACAO'
        }
    }

    # =============== 4. caches e proxy ===============
    Write-Log '' 'DADO'
    $cj = Get-CacheJava
    $totJava = 0.0
    foreach ($c in $cj) { $b = Get-TamanhoPasta -Caminho $c -TimeoutSeg 45; if ($b -gt 0) { $totJava += $b } }
    if ($totJava -gt 0) {
        Write-Log ('Cache do Java Web Start: {0}' -f (Format-Bytes $totJava)) 'DADO'
        Add-ItemInv 'Tasy' 'Cache Java Web Start' (Format-Bytes $totJava)
        if ($totJava -gt 200MB) {
            Add-Achado 'ALERTA' ('Cache do Java Web Start com {0}' -f (Format-Bytes $totJava)) 'Cache antigo faz o Tasy abrir versao errada e falhar ao baixar arquivo. O Modulo 3 limpa.' 'Tasy' 'Medio'
        }
    }

    $centCache = Join-Path $env:LOCALAPPDATA 'CentBrowser\User Data\Default\Cache'
    if (Test-Path -LiteralPath $centCache) {
        $b = Get-TamanhoPasta -Caminho $centCache -TimeoutSeg 45
        Write-Log ('Cache do CentBrowser: {0}' -f (Format-Bytes $b)) 'DADO'
        Add-ItemInv 'Tasy' 'Cache do CentBrowser' (Format-Bytes $b)
        if ($b -gt 300MB) {
            Add-Achado 'ALERTA' ('Cache do CentBrowser com {0}' -f (Format-Bytes $b)) 'O Modulo 3 tem o item, mas vem DESMARCADO: e o navegador do prontuario e a decisao de limpar e sua.' 'Tasy' 'Medio'
        }
    }

    if ((Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyEnable') -eq 1) {
        $exc = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyOverride'
        $faltando = @()
        foreach ($d in $urls) {
            $i = Get-InfoUrl -Url $d.Url
            if (-not $i) { continue }
            if (-not (Test-EnderecoInterno -Servidor $i.Servidor)) { continue }
            if ("$exc" -notlike ('*' + $i.Servidor + '*')) { $faltando += $i.Servidor }
        }
        if ($faltando.Count -gt 0) {
            Add-Achado 'ALERTA' ('Proxy sem excecao para o Tasy interno: {0}' -f ($faltando -join ', ')) 'O trafego do prontuario passa pelo proxy sem necessidade e isso soma tempo em cada tela. Peca ao TI para incluir nas excecoes.' 'Tasy' 'Alto'
        }
    }

    Write-Log '' 'DADO'
    Write-Log 'COMO LER O RESULTADO' 'TITULO'
    Write-Log 'Rede baixa e pagina alta      -> servidor de aplicacao lento. Chamado, com estes numeros.' 'ACAO'
    Write-Log 'Rede alta e pagina alta       -> caminho de rede da estacao. Cabo em vez de Wi-Fi, e chamado de rede.' 'ACAO'
    Write-Log 'Tudo baixo e o Tasy travando  -> e a estacao: navegador com abas demais, cache velho ou memoria cheia.' 'ACAO'
    Write-Log 'Um ambiente rapido e outro lento -> o problema e daquele servidor, nao seu.' 'ACAO'
    Write-Log '' 'DADO'
    Write-Log 'Peca ao TI a exclusao das pastas de cache do navegador e do Java na varredura do antivirus: e ganho direto na navegacao do Tasy.' 'ACAO'
}

# =====================================================================
# 5. MODULO 2 - DIAGNOSTICO
# =====================================================================


function Invoke-DiagMemoria {
    Write-Titulo 'Memoria RAM'
    try {
        $so     = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $totalB = $so.TotalVisibleMemorySize * 1KB
        $livreB = $so.FreePhysicalMemory * 1KB
        $usoPct = [Math]::Round((($totalB - $livreB) / $totalB) * 100, 1)

        $txt = ('Memoria em uso: {0}% ({1} de {2}, livres {3})' -f $usoPct, (Format-Bytes ($totalB - $livreB)), (Format-Bytes $totalB), (Format-Bytes $livreB))
        if ($usoPct -ge $script:Lim.RamCritPct)       { Add-Achado 'CRITICO' $txt 'Feche programas e abas do navegador. O computador esta usando disco como memoria.' }
        elseif ($usoPct -ge $script:Lim.RamAlertaPct) { Add-Achado 'ALERTA'  $txt 'Feche o que nao estiver em uso, principalmente abas do navegador.' }
        else                                          { Add-Achado 'OK' $txt }

        if (($totalB / 1GB) -lt 8) {
            Add-Achado 'ALERTA' ('Memoria total baixa para o uso clinico: {0}' -f (Format-Bytes $totalB)) 'Abaixo de 8 GB, Excel + navegador + Teams travam o computador. Solicite ampliacao.'
        }
    } catch { Write-Log ('Falha ao medir memoria: {0}' -f $_.Exception.Message) 'ALERTA' }

    Write-Log 'Programas que mais consomem memoria:' 'DADO'
    try {
        $grupos = Get-Process -ErrorAction SilentlyContinue |
                  Group-Object ProcessName |
                  ForEach-Object {
                      [pscustomobject]@{
                          Nome       = $_.Name
                          Processos  = $_.Count
                          Memoria    = ($_.Group | Measure-Object WorkingSet64 -Sum).Sum
                      }
                  } | Sort-Object Memoria -Descending | Select-Object -First 10

        foreach ($g in $grupos) {
            Write-Log ('{0,-24} {1,10}   ({2} processo(s))' -f $g.Nome, (Format-Bytes $g.Memoria), $g.Processos) 'DADO'
        }

        $navegadores = $grupos | Where-Object { $_.Nome -match 'msedge|chrome|firefox' }
        foreach ($n in $navegadores) {
            if ($n.Memoria -gt 2GB) {
                Add-Achado 'ALERTA' ('{0} consumindo {1} em {2} processos' -f $n.Nome, (Format-Bytes $n.Memoria), $n.Processos) 'Feche as abas que nao estao em uso. Cada aba aberta e um processo com memoria propria.'
            }
        }
        $teams = $grupos | Where-Object { $_.Nome -match 'Teams|ms-teams' }
        if ($teams -and (($teams | Measure-Object Memoria -Sum).Sum -gt 1.5GB)) {
            Add-Achado 'ALERTA' 'Teams consumindo muita memoria' 'Feche e abra o Teams uma vez por dia, ou limpe o cache no Modulo 3.'
        }
    } catch { }
}






function Invoke-DiagCaches {
    Write-Titulo 'Espaco recuperavel (previa da limpeza)'
    $alvos = Get-AlvosLimpeza
    $total = 0.0
    foreach ($a in $alvos) {
        if ($script:Cancelar) { return }
        Set-Status ('Medindo ' + $a.Nome + '...')
        $b = Measure-Alvo -Alvo $a
        $a.Bytes = $b
        if ($b -gt 0) {
            Write-Log ('{0,-40} {1,10}' -f $a.Nome, (Format-Bytes $b)) 'DADO'
            if ($a.Padrao) { $total += $b }
        }
    }
    $script:AlvosAtuais = $alvos
    if ($total -gt 1GB) {
        Add-Achado 'ALERTA' ('{0} recuperaveis so com limpeza segura' -f (Format-Bytes $total)) 'Rode o Modulo 3 (Limpeza segura).'
    } else {
        Add-Achado 'OK' ('{0} recuperaveis com limpeza segura' -f (Format-Bytes $total))
    }
}



# ---- diagnostico acionavel: so o que o Modulo 3 resolve --------------
function Invoke-DiagSistemaAcionavel {
    Write-Titulo 'Estado da maquina'
    try {
        $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        Write-Log ('{0} · build {1} · {2} de memoria' -f $so.Caption, $so.BuildNumber, (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))) 'DADO'

        $horas = ((Get-Date) - $so.LastBootUpTime).TotalHours
        $txt = ('{0:N0} dias e {1:N0} horas ligado sem reiniciar' -f [Math]::Floor($horas / 24), ($horas % 24))
        if ($horas -ge $script:Lim.UptimeCritH) {
            Add-Achado 'CRITICO' $txt 'Reinicie o computador ao final da limpeza. Usar "Desligar" nao resolve: so "Reiniciar" limpa a memoria.' 'Reiniciar' 'Alto'
        } elseif ($horas -ge $script:Lim.UptimeAlertaH) {
            Add-Achado 'ALERTA' $txt 'Reinicie ao final da limpeza (botao no menu lateral).' 'Reiniciar' 'Medio'
        } else {
            Add-Achado 'OK' $txt '' 'Reiniciar'
        }

        if ((Get-ValorReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled') -eq 1) {
            Write-Log 'Inicializacao Rapida ligada: "Desligar" nao limpa a memoria, so "Reiniciar".' 'DADO'
        }
    } catch { }
}

function Invoke-DiagEspaco {
    Write-Titulo 'Espaco em disco'
    try {
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)) {
            if (-not $d.Size) { continue }
            $pct = [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
            $txt = ('Unidade {0}: {1} livres de {2} ({3}%)' -f $d.DeviceID, (Format-Bytes $d.FreeSpace), (Format-Bytes $d.Size), $pct)
            if ($pct -lt $script:Lim.DiscoCritPct) {
                Add-Achado 'CRITICO' $txt 'Disco quase cheio e a causa numero um de lentidao. O Modulo 3 libera espaco agora.' 'Disco' 'Alto'
            } elseif ($pct -lt $script:Lim.DiscoAlertaPct) {
                Add-Achado 'ALERTA' $txt 'Rode o Modulo 3 para liberar espaco.' 'Disco' 'Alto'
            } else {
                Add-Achado 'OK' $txt '' 'Disco'
            }
        }
    } catch { Write-Log 'Nao foi possivel ler as unidades.' 'ALERTA' }

    # C:\Windows\Temp nao e do usuario, mas pesa e serve de anexo no chamado
    try {
        $wt = Get-TamanhoPasta -Caminho (Join-Path $env:SystemRoot 'Temp') -TimeoutSeg 45
        if ($wt -gt 1GB) {
            Add-Achado 'ALERTA' ('C:\Windows\Temp com {0}' -f (Format-Bytes $wt)) 'Pasta do sistema: so o TI limpa, e a rotina de limpeza deles ja cobre. Registre o tamanho no chamado.' 'Disco' 'Medio'
        } elseif ($wt -gt 0) {
            Write-Log ('C:\Windows\Temp: {0} (do sistema, fora do alcance do usuario)' -f (Format-Bytes $wt)) 'DADO'
        }
    } catch { }
}

function Invoke-DiagLixeira {
    Write-Titulo 'Lixeira'
    $alvo = New-Alvo -Id 'LIXEIRA' -Nome 'Lixeira' -Modo 'Lixeira' -Caminhos @()
    $b = Measure-Alvo -Alvo $alvo
    if ($b -gt 500MB) {
        Add-Achado 'ALERTA' ('Lixeira com {0}' -f (Format-Bytes $b)) 'Esvaziar a Lixeira e rotina do Modulo 3. Confira antes se nao ha nada a recuperar.' 'Disco' 'Alto'
    } else {
        Add-Achado 'OK' ('Lixeira com {0}' -f (Format-Bytes $b)) '' 'Disco'
    }
}

function Invoke-DiagEncerraveis {
    Write-Titulo 'Programas que podem ser encerrados agora com seguranca'
    Write-Log 'Criterio: nao guardam documento aberto e voltam sozinhos quando voce abrir de novo.' 'DADO'

    $enc = Get-ProcessosEncerraveis
    if ($enc.Count -eq 0) {
        Add-Achado 'OK' 'Nenhum programa dispensavel rodando agora' '' 'Memoria'
        Write-Log 'Nada a encerrar: o que esta aberto e trabalho ou sistema.' 'DADO'
        return
    }

    $total = ($enc | Measure-Object Memoria -Sum).Sum
    foreach ($e in ($enc | Sort-Object Memoria -Descending)) {
        Write-Log ('{0,-34} {1,10}   ({2} processo(s))' -f $e.Rotulo, (Format-Bytes $e.Memoria), $e.Qtd) 'DADO'
        if ($e.Nota) { Write-Log ('   ' + $e.Nota) 'DADO' }
    }

    if ($total -gt 500MB) {
        Add-Achado 'ALERTA' ('{0} de memoria presos em {1} programa(s) dispensavel(is)' -f (Format-Bytes $total), $enc.Count) 'O Modulo 3 encerra todos de uma vez. Eles voltam quando voce abrir.' 'Memoria' 'Alto'
    } else {
        Add-Achado 'OK' ('{0} recuperaveis encerrando programas dispensaveis' -f (Format-Bytes $total)) '' 'Memoria'
    }

    Write-Log '' 'DADO'
    Write-Log 'Programas de trabalho, clinicos, de seguranca e navegadores nunca sao encerrados pelo kit.' 'DADO'
}

function Invoke-DiagInicioAcionavel {
    Write-Titulo 'Inicializacao do Windows'

    # quanto tempo levou do boot ate a area de trabalho
    try {
        $so  = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $exp = $null
        foreach ($e in @(Get-Process explorer -ErrorAction SilentlyContinue)) {
            try { $st = $e.StartTime } catch { continue }   # processo de outra sessao: acesso negado
            if (-not $exp -or $st -lt $exp.StartTime) { $exp = $e }
        }
        if ($exp -and $exp.StartTime -gt $so.LastBootUpTime) {
            $seg = [Math]::Round(($exp.StartTime - $so.LastBootUpTime).TotalSeconds)
            $script:LogonSegundos = $(if ($seg -le 1200) { $seg } else { 0 })
            if ($seg -gt 1200) {
                # a maquina ficou ligada antes do logon: a conta nao mede inicializacao
                Write-Log ('A maquina ficou {0:N1} h ligada antes deste logon.' -f ($seg / 3600)) 'DADO'
                Write-Log 'Por isso nao da para medir o tempo de inicializacao nesta sessao.' 'DADO'
                Write-Log 'Para medir: reinicie e rode o Modulo 2 logo depois de entrar.' 'ACAO'
            }
            elseif ($seg -gt 180) { Add-Achado 'CRITICO' ('Inicializacao levou {0} segundos' -f $seg) 'Veja abaixo o que pesa: itens de inicializacao, unidades de rede e impressoras de rede. O Modulo 3 resolve a parte do usuario.' 'Inicializacao' 'Alto' }
            elseif ($seg -gt 90)  { Add-Achado 'ALERTA'  ('Inicializacao levou {0} segundos' -f $seg) 'O Modulo 3 reduz isso desativando os itens seguros.' 'Inicializacao' 'Alto' }
            else                  { Add-Achado 'OK' ('Inicializacao em {0} segundos' -f $seg) '' 'Inicializacao' }
        }
    } catch { }

    $itens = Get-ItensInicializacao
    $seguros = @($itens | Where-Object { $_.Seguro })
    $outros  = @($itens | Where-Object { -not $_.Seguro })

    Write-Log ('Itens do seu usuario que abrem junto com o Windows: {0}' -f $itens.Count) 'DADO'
    Write-Log '' 'DADO'
    if ($seguros.Count -gt 0) {
        Write-Log 'Seguro NAO carregar de novo (o programa continua funcionando ao ser aberto):' 'DADO'
        foreach ($i in $seguros) { Write-Log ('[ desativar ] {0,-30} {1}' -f $i.Nome, $i.Rotulo) 'DADO' }
    }
    if ($outros.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log 'Mantidos como estao:' 'DADO'
        foreach ($i in $outros) { Write-Log ('[  manter  ] {0,-30} {1}' -f $i.Nome, $i.Rotulo) 'DADO' }
    }

    # logon longo com poucos itens de usuario: a conta esta em outro lugar
    if ($script:LogonSegundos -gt 90) {
        $unidades = @()
        try { $unidades = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=4' -ErrorAction SilentlyContinue) } catch { }
        $impRede = @()
        try { $impRede = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | Where-Object { $_.Network -or "$($_.PortName)" -like '\\*' -or "$($_.Name)" -like '\\*' }) } catch { }

        if ($unidades.Count -gt 0 -or $impRede.Count -gt 0) {
            Write-Log '' 'DADO'
            Write-Log 'O que mais pesa no logon desta estacao:' 'DADO'
            Write-Log ('{0} unidade(s) de rede e {1} impressora(s) de rede sao reconectadas a cada logon.' -f $unidades.Count, $impRede.Count) 'DADO'

            $lentos = @()
            foreach ($u in $unidades) {
                $srv = ("$($u.ProviderName)".Trim() -replace '^\\+', '' -split '\\')[0]
                if (-not $srv) { continue }
                $t = Test-PortaTcp -Alvo $srv -Porta 445 -TimeoutMs 700
                Write-Log ('{0} -> {1}   {2}' -f $u.DeviceID, $u.ProviderName, $(if ($t.Ok) { "$($t.Ms) ms" } else { 'SEM RESPOSTA' })) 'DADO'
                if (-not $t.Ok) { $lentos += ('{0} ({1})' -f $u.DeviceID, $srv) }
            }
            foreach ($i in ($impRede | Select-Object -First 6)) { Write-Log ('impressora: {0}' -f $i.Name) 'DADO' }

            if ($lentos.Count -gt 0) {
                Add-Achado 'CRITICO' ('{0} unidade(s) de rede sem resposta agora: {1}' -f $lentos.Count, ($lentos -join ', ')) 'Cada uma dessas o Windows tenta reconectar no logon e espera o tempo limite. E a explicacao mais provavel do logon longo. Peca ao TI para remover o mapeamento ou liberar o acesso.' 'Inicializacao' 'Alto'
            } elseif ($seguros.Count -lt 3) {
                Add-Achado 'ALERTA' ('Logon longo com apenas {0} item(ns) de inicializacao do usuario' -f $seguros.Count) ('O tempo nao vem dos programas do seu usuario. Sao {0} unidade(s) de rede e {1} impressora(s) de rede reconectadas a cada logon, mais os agentes corporativos. Leve estes numeros ao TI.' -f $unidades.Count, $impRede.Count) 'Inicializacao' 'Alto'
            }
        }
    }

    if ($seguros.Count -ge 3) {
        Add-Achado 'ALERTA' ('{0} programas abrem sozinhos e podem ser desativados com seguranca' -f $seguros.Count) 'Cada um deles disputa disco e CPU no logon. O Modulo 3 desativa todos de uma vez e o botao "Desfazer" reverte.' 'Inicializacao' 'Alto'
    } elseif ($seguros.Count -gt 0) {
        Add-Achado 'ALERTA' ('{0} programa(s) na inicializacao podem ser desativados' -f $seguros.Count) 'O Modulo 3 desativa.' 'Inicializacao' 'Medio'
    } else {
        Add-Achado 'OK' 'Inicializacao do usuario ja esta enxuta' '' 'Inicializacao'
    }
    if ($outros.Count -gt 0) {
        Write-Log 'Os itens da maquina (para todos os usuarios) e os de seguranca so o TI altera.' 'DADO'
    }
}

function Invoke-DiagAjustesPendentes {
    Write-Titulo 'Ajustes de desempenho pendentes'
    $aj = Get-AjustesPendentes
    if ($aj.Count -eq 0) {
        Add-Achado 'OK' 'Todos os ajustes de desempenho ja estao aplicados' '' 'Ajustes'
        return
    }
    foreach ($a in $aj) {
        Write-Log ('[ aplicar ] {0}' -f $a.Rotulo) 'DADO'
        Write-Log ('   ' + $a.Nota) 'DADO'
    }
    Add-Achado 'ALERTA' ('{0} ajuste(s) de desempenho ainda nao aplicado(s)' -f $aj.Count) 'O Modulo 3 aplica todos. Reversivel pelo botao "Desfazer otimizacoes".' 'Ajustes' 'Alto'
}

function Invoke-DiagMapeamentos {
    Write-Titulo 'Unidades de rede'
    $mortos = Get-MapeamentosMortos
    if ($mortos.Count -eq 0) {
        Add-Achado 'OK' 'Nenhuma unidade de rede desconectada' '' 'Rede'
        return
    }
    foreach ($m in $mortos) { Write-Log ('{0} -> {1}  [desconectada]' -f $m.Letra, $m.Destino) 'DADO' }
    Add-Achado 'CRITICO' ('{0} unidade(s) de rede desconectada(s)' -f $mortos.Count) 'Mapeamento morto congela o Explorer e o "Salvar como" do Office por 30 segundos ou mais. O Modulo 3 remove.' 'Rede' 'Alto'
}

# =====================================================================
# 6B. PERSISTENCIA - por que a estacao nao guarda ajuste no logoff
#
# Duas partes, de proposito. A primeira infere pelo estado do perfil e
# responde na hora. A segunda deixa marcador e so conclui depois de um
# logoff - e e ela que vale como prova, porque nao depende de palpite.
# =====================================================================
$script:MarcadorChaveReg = 'HKCU:\Software\KitSuporteRT'

function Get-SessaoAtual {
    # Identifica a sessao de logon de agora. Se o horario de inicio do
    # explorer mudou entre duas execucoes, houve logoff ou reinicio.
    $boot = ''; $logon = ''
    try { $boot = '{0:yyyy-MM-dd HH:mm:ss}' -f (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime } catch { }
    try {
        $meu = (Get-Process -Id $PID -ErrorAction Stop).SessionId
        $ex  = @(Get-Process explorer -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $meu } | Sort-Object StartTime)
        if ($ex.Count -gt 0) { $logon = '{0:yyyy-MM-dd HH:mm:ss}' -f $ex[0].StartTime }
    } catch { }
    return [pscustomobject]@{ Boot = $boot; Logon = $logon }
}

function Get-LocaisMarcador {
    # Cada local responde por uma pergunta diferente. O de disco so entra
    # se a pasta clinica existir - sem ela nao da para separar perfil
    # descartado de disco congelado.
    $lista = @(
        [pscustomobject]@{ Id='HKCU';    Rotulo='Registro do usuario (HKCU)'; Tipo='Reg'; Caminho=$script:MarcadorChaveReg; Prova='ajuste de registro do Modulo 3' }
        [pscustomobject]@{ Id='LOCAL';   Rotulo='AppData\Local do perfil';    Tipo='Arq'; Caminho=(Join-Path $env:LOCALAPPDATA 'KitSuporteRT\marcador.txt'); Prova='estado do Desfazer' }
        [pscustomobject]@{ Id='ROAMING'; Rotulo='AppData\Roaming do perfil';  Tipo='Arq'; Caminho=(Join-Path $env:APPDATA 'KitSuporteRT\marcador.txt');      Prova='autorrecuperacao do Excel' }
    )
    try {
        $util = Join-Path $script:PastaClinica 'UTILITARIOS'
        if (Test-Path -LiteralPath $util) {
            $lista += [pscustomobject]@{ Id='DISCO'; Rotulo='Disco fora do perfil'; Tipo='Arq'; Caminho=(Join-Path $util 'KitSuporteRT_marcador.txt'); Prova='portais, POPs e o proprio kit' }
        }
    } catch { }
    return $lista
}

function Read-Marcador {
    param($Local)
    try {
        if ($Local.Tipo -eq 'Reg') {
            return [string](Get-ValorReg $Local.Caminho 'Marcador')
        }
        if (Test-Path -LiteralPath $Local.Caminho) {
            return ([System.IO.File]::ReadAllText($Local.Caminho)).Trim()
        }
    } catch { }
    return ''
}

function Write-Marcador {
    param($Local, [string]$Conteudo)
    try {
        if ($Local.Tipo -eq 'Reg') {
            if (-not (Test-Path $Local.Caminho)) { New-Item -Path $Local.Caminho -Force -ErrorAction Stop | Out-Null }
            New-ItemProperty -Path $Local.Caminho -Name 'Marcador' -Value $Conteudo -PropertyType String -Force -ErrorAction Stop | Out-Null
            return $true
        }
        $pai = Split-Path -Parent $Local.Caminho
        if (-not (Test-Path -LiteralPath $pai)) { New-Item -ItemType Directory -Path $pai -Force -ErrorAction Stop | Out-Null }
        [System.IO.File]::WriteAllText($Local.Caminho, $Conteudo, (New-Object System.Text.UTF8Encoding($false)))
        return $true
    } catch { return $false }
}

function Get-DiagnosticoPerfil {
    # Tudo aqui e leitura. Nenhum item exige administrador.
    $r = [pscustomobject]@{
        Caminho = "$env:USERPROFILE"; Tipo = 'Local'; Obrigatorio = $false; Temporario = $false
        Movel = $false; CaminhoMovel = ''; Filtro = ''; ApagaCache = $false; Motivos = @()
    }
    $sid = ''
    try { $sid = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value } catch { }

    # Perfil obrigatorio: NTUSER.MAN no lugar do NTUSER.DAT. HKCU vira
    # somente leitura e tudo o que o Modulo 3 ajusta morre no logoff.
    try {
        if (Test-Path -LiteralPath (Join-Path $env:USERPROFILE 'NTUSER.MAN')) {
            $r.Obrigatorio = $true; $r.Tipo = 'Obrigatorio'
            $r.Motivos += 'NTUSER.MAN presente na raiz do perfil'
        }
    } catch { }

    # Perfil temporario: o Windows carrega um perfil descartavel quando
    # nao consegue abrir o do usuario. Sinal classico e a chave .bak.
    try {
        if ($sid) {
            $base = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\'
            $p = Get-ItemProperty -Path ($base + $sid) -ErrorAction SilentlyContinue
            if ($p) {
                $r.Caminho = "$($p.ProfileImagePath)"
                if (($p.State -band 4) -ne 0) { $r.Temporario = $true; $r.Motivos += ('State do perfil = {0} (bit de temporario ligado)' -f $p.State) }
            }
            if (Test-Path -Path ($base + $sid + '.bak')) {
                $r.Temporario = $true; $r.Motivos += 'existe chave .bak: o Windows trocou o perfil por um descartavel'
            }
        }
        if ("$env:USERPROFILE" -match '\\TEMP(\.|\\|$)') {
            $r.Temporario = $true; $r.Motivos += 'a pasta do perfil e TEMP'
        }
    } catch { }
    if ($r.Temporario) { $r.Tipo = 'Temporario' }

    # Perfil movel: o que vale e a copia do servidor. Se ela nao salvar,
    # ou se a politica apagar o cache local, tudo volta atras.
    try {
        $up = Get-CimInstance Win32_UserProfile -Filter "SID='$sid'" -ErrorAction SilentlyContinue
        if ($up) {
            if ($up.RoamingConfigured) {
                $r.Movel = $true; $r.CaminhoMovel = "$($up.RoamingPath)"
                if (-not $r.Obrigatorio -and -not $r.Temporario) { $r.Tipo = 'Movel' }
                $r.Motivos += ('perfil movel apontando para {0}' -f $up.RoamingPath)
            }
            if ($up.Temporary) { $r.Temporario = $true; $r.Tipo = 'Temporario'; $r.Motivos += 'Win32_UserProfile marca o perfil como temporario' }
        }
    } catch { }

    foreach ($k in @('HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon',
                     'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System')) {
        try {
            if ((Get-ValorReg $k 'DeleteRoamingCache') -eq 1) {
                $r.ApagaCache = $true
                $r.Motivos += 'politica DeleteRoamingCache ligada: a copia local e apagada no logoff'
            }
        } catch { }
    }

    # Filtro de escrita / congelador de disco: revertem no reinicio, nao
    # no logoff. Separar isso e o que evita tratar a causa errada.
    try {
        foreach ($s in @('UWFServicingSvc','DFServ','FrzState2k','ShadowDefender')) {
            $x = Get-Service -Name $s -ErrorAction SilentlyContinue
            if ($x) { $r.Filtro = ('{0} ({1})' -f $s, $x.Status) ; break }
        }
        if (-not $r.Filtro -and (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Services\uwfvol')) { $r.Filtro = 'driver uwfvol presente' }
        if (-not $r.Filtro) {
            $cong = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match '^DFServ$|^FrzState2k$|^ShadowDefender$' })
            if ($cong.Count -gt 0) { $r.Filtro = ($cong[0].ProcessName + ' em execucao') }
        }
    } catch { }
    return $r
}

function Invoke-DiagPersistencia {
    Clear-SnapshotsOperacao
    Write-Titulo 'Persistencia - o que sobrevive ao logoff'
    $d = Get-DiagnosticoPerfil

    Write-Log ('Perfil ....: {0}' -f $d.Caminho) 'DADO'
    Write-Log ('Tipo ......: {0}' -f $d.Tipo) 'DADO'
    Add-ItemInv 'Persistencia' 'Tipo de perfil' $d.Tipo $d.Caminho
    if ($d.Movel)  { Write-Log ('Copia movel: {0}' -f $d.CaminhoMovel) 'DADO'; Add-ItemInv 'Persistencia' 'Perfil movel' $d.CaminhoMovel }
    if ($d.Filtro) { Add-ItemInv 'Persistencia' 'Filtro de escrita' $d.Filtro }
    foreach ($m in $d.Motivos) { Write-Log ('   . ' + $m) 'DADO' }

    if ($d.Obrigatorio) {
        Add-Achado 'CRITICO' 'Perfil obrigatorio: o Windows descarta suas alteracoes em todo logoff' 'Nada que o Modulo 3 ajusta no registro sobrevive, e o Desfazer nunca acha o que desfazer. So o TI muda isso - leve este log ao chamado.' 'Persistencia' 'Alto'
    } elseif ($d.Temporario) {
        Add-Achado 'CRITICO' 'Perfil temporario: esta sessao inteira e descartada no logoff' 'O Windows nao conseguiu abrir o seu perfil e criou um descartavel. Nao adianta ajustar nada agora. Chamado para o TI recriar o perfil.' 'Persistencia' 'Alto'
    } elseif ($d.ApagaCache) {
        Add-Achado 'ALERTA' 'Politica apaga a copia local do perfil no logoff' 'O que nao subir para o servidor a tempo se perde. Vale conferir com o TI se o perfil movel esta salvando.' 'Persistencia' 'Alto'
    } elseif ($d.Movel) {
        Add-Achado 'ALERTA' 'Perfil movel: o que vale e a copia do servidor' 'Se o logoff nao terminar de sincronizar, o ajuste se perde. Feche os programas antes de sair.' 'Persistencia' 'Medio'
    } else {
        Add-Achado 'OK' 'Perfil local: o registro do usuario deve sobreviver ao logoff' '' 'Persistencia'
    }

    if ($d.Filtro) {
        Add-Achado 'CRITICO' ('Filtro de escrita no disco: {0}' -f $d.Filtro) 'O disco volta ao estado anterior a cada reinicio. Nada instalado ou ajustado permanece. So o TI desliga - leve este log ao chamado.' 'Persistencia' 'Alto'
    }

    # ---- a prova ----
    $sessao = Get-SessaoAtual
    $locais = Get-LocaisMarcador
    $anterior = ''
    foreach ($id in @('DISCO','ROAMING','LOCAL','HKCU')) {
        $l = $locais | Where-Object { $_.Id -eq $id }
        if ($l) { $v = Read-Marcador $l; if ($v) { $anterior = $v; break } }
    }

    if (-not $anterior) {
        Write-Log '' 'DADO'
        Write-Log 'Nenhum marcador anterior. Deixando um agora em cada lugar.' 'DADO'
        Write-Log 'Faca logoff (ou reinicie) e rode este modulo de novo: ele dira exatamente o que sobreviveu.' 'ACAO'
    } else {
        $partes = $anterior -split '\|'
        $logonAntes = $(if ($partes.Count -ge 3) { $partes[2] } else { '' })
        Write-Log '' 'DADO'
        if ($logonAntes -and $logonAntes -eq $sessao.Logon) {
            Write-Log ('Marcador de {0}, da mesma sessao de logon. Ainda nao houve logoff para comparar.' -f $partes[0]) 'DADO'
            Write-Log 'Faca logoff (ou reinicie) e rode de novo.' 'ACAO'
        } else {
            Write-Log ('Marcador anterior de {0}. Houve logoff desde entao - este e o resultado:' -f $partes[0]) 'DADO'
            $perdidos = @()
            foreach ($l in $locais) {
                $v = Read-Marcador $l
                if ($v -eq $anterior) {
                    Write-Log ('   SOBREVIVEU  {0,-30} guarda {1}' -f $l.Rotulo, $l.Prova) 'OK'
                } else {
                    Write-Log ('   PERDEU      {0,-30} perde {1}' -f $l.Rotulo, $l.Prova) 'ALERTA'
                    $perdidos += $l
                }
            }
            if ($perdidos.Count -eq 0) {
                Add-Achado 'OK' 'Tudo sobreviveu ao ultimo logoff' '' 'Persistencia'
            } else {
                $ids = @($perdidos | Select-Object -ExpandProperty Id)
                if ($ids -contains 'DISCO') {
                    Add-Achado 'CRITICO' 'Nem o disco fora do perfil sobreviveu ao logoff' 'Isso e filtro de escrita ou congelador de disco, nao problema de perfil. So o TI desliga.' 'Persistencia' 'Alto'
                } else {
                    Add-Achado 'CRITICO' ('O perfil nao guarda alteracao: {0} de {1} lugares perderam o marcador' -f $perdidos.Count, $locais.Count) ('Perdeu: ' + (($perdidos | Select-Object -ExpandProperty Rotulo) -join '; ') + '. O disco fora do perfil ficou. Isso e perfil descartado, e so o TI corrige.') 'Persistencia' 'Alto'
                }
            }
        }
    }

    $carimbo = '{0:yyyy-MM-dd HH:mm:ss}|{1}|{2}' -f (Get-Date), $sessao.Boot, $sessao.Logon
    $ok = 0
    foreach ($l in $locais) { if (Write-Marcador $l $carimbo) { $ok++ } }
    Write-Log ('Marcador renovado em {0} de {1} lugares.' -f $ok, $locais.Count) 'DADO'
    if ($ok -lt $locais.Count) { Write-Log 'Algum lugar nem aceitou gravar agora - o que ja e resposta.' 'ALERTA' }
}

# =====================================================================
# 4F. REGISTRO DO ECOSSISTEMA - so leitura, so relato
#
# <raiz>\RADIOTERAPIA_AI\_instalados.json e o registro compartilhado: cada
# instalador do ecossistema grava a propria entrada. O kit LE e RELATA.
# Nunca decide nada com base nele - nem o que proteger, nem o que limpar.
#
# Por que a fronteira importa: registro que mente e pior que registro vazio.
# Vazio faz a pessoa ir olhar; errado faz ela concluir. Nesta maquina ele
# declara AUTO_CONTORNO 2.1 em C:\RADIOTERAPIA_AI\AUTO_CONTORNO, pasta que nao
# existe - o app mora em LOCAL_SUITE\_montagem\. Quem protege a carga em
# execucao e a marca no caminho do processo (secao 2B das regras), que nao
# depende deste arquivo estar em dia.
# =====================================================================
function Get-ArquivoRegistroEcossistema {
    # Procura por MARCA em disco LOCAL, nunca por caminho fixo e nunca em
    # unidade de rede: "if exist X:\" em letra mapeada morta custa o tempo
    # limite do SMB, e isso ja travou o passo 3 da preparacao por 4 minutos.
    try {
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)) {
            $f = Join-Path ("$($d.DeviceID)\" + $script:MarcaEcossistema) '_instalados.json'
            if (Test-Path -LiteralPath $f) { return $f }
        }
    } catch { }
    return ''
}

function Test-CaminhoLocalExiste {
    # Caminho de rede nao e testado: servidor fora do ar devolveria o tempo
    # limite do SMB em vez de resposta. Sem resposta e "nao sei", nao "nao existe".
    param([string]$Caminho)
    if (-not $Caminho) { return $null }
    if ($Caminho -like '\\*') { return $null }
    try { return [bool](Test-Path -LiteralPath $Caminho) } catch { return $null }
}

function Get-AppsEcossistema {
    $r = [pscustomobject]@{ Arquivo = ''; Estado = 'ausente'; Apps = @(); SemRegistro = @() }
    $f = Get-ArquivoRegistroEcossistema
    if (-not $f) { return $r }
    $r.Arquivo = $f

    $txt = ''
    try { $txt = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8) } catch { $r.Estado = 'ilegivel'; return $r }
    if (-not "$txt".Trim()) { $r.Estado = 'vazio'; return $r }
    # No 5.1, JSON vazio devolve nulo em silencio e JSON truncado lanca. Os dois
    # tem de ser tratados, e nenhum pode virar caixa de erro na tela.
    $o = $null
    try { $o = $txt | ConvertFrom-Json -ErrorAction Stop } catch { $r.Estado = 'invalido'; return $r }
    if (-not $o) { $r.Estado = 'vazio'; return $r }

    $lista = @()
    foreach ($prop in $o.PSObject.Properties) {
        $cam = ''; $ver = ''; $dat = ''
        try { $cam = "$($prop.Value.caminho)" } catch { }
        try { $ver = "$($prop.Value.versao)" } catch { }
        try { $dat = "$($prop.Value.data)" } catch { }
        $lista += [pscustomobject]@{
            Projeto = $prop.Name; Versao = $ver; Data = $dat
            Caminho = $cam; Existe = (Test-CaminhoLocalExiste $cam)
        }
    }
    $r.Estado = 'ok'
    $r.Apps = @($lista | Sort-Object Projeto)

    # A mentira na direcao oposta: instalado no disco e ausente do registro.
    # Pasta com entrega.txt e instalacao; sem ele e pasta de dado ou de apoio,
    # e reportar essas geraria chamado para nao-problema.
    try {
        $raiz = Split-Path -Parent $f
        $nomes = @($r.Apps | Select-Object -ExpandProperty Projeto)
        foreach ($d in (Get-ChildItem -LiteralPath $raiz -Directory -ErrorAction Stop)) {
            if ($nomes -contains $d.Name) { continue }
            if (Test-Path -LiteralPath (Join-Path $d.FullName 'entrega.txt')) { $r.SemRegistro += $d.Name }
        }
    } catch { }
    return $r
}

function Invoke-InvEcossistema {
    Write-Titulo 'Aplicativos do radioterapia.ai'
    $r = Get-AppsEcossistema

    if ($r.Estado -eq 'ausente') {
        Write-Log 'Nenhum registro do ecossistema nesta estacao. E o estado normal de quem nao instalou nada.' 'DADO'
        Add-ItemInv 'Ecossistema' 'Registro' 'ausente'
        return
    }
    Write-Log ('Registro ..: {0}' -f $r.Arquivo) 'DADO'

    if ($r.Estado -ne 'ok') {
        $porque = switch ($r.Estado) {
            'ilegivel' { 'nao foi possivel ler o arquivo' }
            'vazio'    { 'o arquivo esta vazio' }
            default    { 'o conteudo nao e JSON valido' }
        }
        Add-Achado 'ALERTA' ('Registro do ecossistema ilegivel: {0}' -f $porque) 'Enquanto isso durar nenhum aplicativo consegue registrar versao, e o suporte fica sem saber o que esta instalado. Leve este log a quem cuida do instalador.' 'Ecossistema' 'Medio'
        Add-ItemInv 'Ecossistema' 'Registro' $r.Estado $r.Arquivo
        return
    }

    Write-Log ('{0} aplicativo(s) declarado(s):' -f $r.Apps.Count) 'DADO'
    $divergentes = @()
    foreach ($a in $r.Apps) {
        $situacao = if ($a.Existe -eq $true) { 'pasta ok' } elseif ($null -eq $a.Existe) { 'nao verificavel (caminho de rede)' } else { 'PASTA NAO EXISTE' }
        Write-Log ('   {0,-24} v{1,-10} {2,-12} {3}' -f $a.Projeto, $a.Versao, $a.Data, $situacao) 'DADO'
        Write-Log ('      {0}' -f $a.Caminho) 'DADO'
        Add-ItemInv 'Ecossistema' $a.Projeto ('v' + $a.Versao) ('{0} | {1} | {2}' -f $a.Data, $a.Caminho, $situacao)
        if ($a.Existe -eq $false) { $divergentes += $a }
    }

    if ($divergentes.Count -gt 0) {
        $det = ($divergentes | ForEach-Object { '{0} v{1} -> {2}' -f $_.Projeto, $_.Versao, $_.Caminho }) -join '; '
        Add-Achado 'ALERTA' ('{0} entrada(s) do registro apontam para pasta que nao existe' -f $divergentes.Count) ('O registro declara mas o disco nao tem: ' + $det + '. Registro que mente e pior que registro vazio: vazio faz procurar, errado faz concluir. Leve ao responsavel pelo instalador - o kit nao corrige registro de outro projeto.') 'Ecossistema' 'Medio'
    }
    if ($r.SemRegistro.Count -gt 0) {
        Add-Achado 'ALERTA' ('{0} aplicativo(s) instalado(s) e ausente(s) do registro' -f $r.SemRegistro.Count) ('Tem entrega.txt no disco mas nao aparece no registro: ' + ($r.SemRegistro -join ', ') + '. E a mentira na direcao oposta, e esconde do suporte a versao que esta em uso.') 'Ecossistema' 'Medio'
    }
    if ($divergentes.Count -eq 0 -and $r.SemRegistro.Count -eq 0) {
        Add-Achado 'OK' 'Registro do ecossistema confere com o disco' '' 'Ecossistema'
    }
}

function Invoke-InventarioDiagnostico {
    Clear-SnapshotsOperacao
    $script:Achados.Clear()
    $script:ItensInv.Clear()
    Write-Titulo 'Modulo 2 - Inventario e diagnostico'
    Write-Log ('Maquina {0} · usuario {1} · {2:dd/MM/yyyy HH:mm:ss}' -f $env:COMPUTERNAME, $env:USERNAME, (Get-Date)) 'DADO'
    Write-Log 'Parte 1: retrato da maquina. Parte 2: o que o Modulo 3 resolve sem administrador.' 'DADO'

    $etapas = @(
        @{ Nome = 'Inv · identificacao';   Fn = { Invoke-InvIdentificacao } }
        @{ Nome = 'Persistencia';          Fn = { Invoke-DiagPersistencia } }
        @{ Nome = 'Apps do ecossistema';   Fn = { Invoke-InvEcossistema } }
        @{ Nome = 'Inv · software';        Fn = { Invoke-InvSoftware } }
        @{ Nome = 'Inv · inicializacao';   Fn = { Invoke-InvInicializacaoCompleta } }
        @{ Nome = 'Inv · servicos';        Fn = { Invoke-InvServicos } }
        @{ Nome = 'Inv · Office e Excel';  Fn = { Invoke-InvOfficeExcel } }
        @{ Nome = 'Inv · pasta clinica';   Fn = { Invoke-InvPastaClinica } }
        @{ Nome = 'Inv · rede';            Fn = { Invoke-InvRedeImpressoras } }
        @{ Nome = 'Inv · migracao';        Fn = { Invoke-InvMigracao } }
        @{ Nome = 'Estado da maquina';     Fn = { Invoke-DiagSistemaAcionavel } }
        @{ Nome = 'Espaco em disco';       Fn = { Invoke-DiagEspaco } }
        @{ Nome = 'Memoria';               Fn = { Invoke-DiagMemoria } }
        @{ Nome = 'Encerraveis agora';     Fn = { Invoke-DiagEncerraveis } }
        @{ Nome = 'Sessao ativa';          Fn = { Invoke-DiagSessaoAtiva } }
        @{ Nome = 'Inicializacao';         Fn = { Invoke-DiagInicioAcionavel } }
        @{ Nome = 'OneDrive e Teams';      Fn = { Invoke-DiagNuvem } }
        @{ Nome = 'Caches recuperaveis';   Fn = { Invoke-DiagCaches } }
        @{ Nome = 'Lixeira';               Fn = { Invoke-DiagLixeira } }
        @{ Nome = 'Downloads';             Fn = { Invoke-DiagDownloads } }
        @{ Nome = 'Ajustes pendentes';     Fn = { Invoke-DiagAjustesPendentes } }
        @{ Nome = 'Unidades de rede';      Fn = { Invoke-DiagMapeamentos } }
        @{ Nome = 'Citrix';                Fn = { Invoke-DiagCitrix } }
        @{ Nome = 'Tasy';                  Fn = { Invoke-DiagTasy } }
    )
    $i = 0
    foreach ($e in $etapas) {
        if ($script:Cancelar) { Write-Log 'Diagnostico interrompido pelo usuario.' 'ALERTA'; break }
        $i++
        Set-Status ('Diagnostico {0}/{1}: {2}' -f $i, $etapas.Count, $e.Nome)
        try { & $e.Fn } catch { Write-Log ('Falha na etapa {0}: {1}' -f $e.Nome, $_.Exception.Message) 'ALERTA' }
    }

    # ---------------- mapa dos pontos criticos ----------------
    Write-Titulo 'Mapa dos pontos criticos'
    $crit = @($script:Achados | Where-Object { $_.Severidade -eq 'CRITICO' })
    $alt  = @($script:Achados | Where-Object { $_.Severidade -eq 'ALERTA' })

    if ($crit.Count -eq 0 -and $alt.Count -eq 0) {
        Write-Log 'Nada a corrigir por aqui: a maquina esta limpa do lado do usuario.' 'OK'
        Write-Log 'Se a lentidao continuar, o gargalo exige administrador (disco, antivirus, rede, hardware).' 'DADO'
        Write-Log 'Copie este log e abra chamado: ele ja mostra que o lado do usuario esta limpo.' 'ACAO'
    } else {
        Write-Log 'Onde esta o problema, por area:' 'DADO'
        $porArea = $script:Achados | Group-Object Categoria | ForEach-Object {
            [pscustomobject]@{ Area = $_.Name; Alto = @($_.Group | Where-Object { $_.Impacto -eq 'Alto' }).Count; Total = $_.Count }
        } | Sort-Object Alto, Total -Descending
        foreach ($a in $porArea) {
            $barra = ('#' * [Math]::Min(24, ($a.Alto * 4 + $a.Total)))
            Write-Log ('{0,-16} {1,-26} {2} achado(s), {3} de alto impacto' -f $a.Area, $barra, $a.Total, $a.Alto) 'DADO'
        }

        Write-Log ''
        Write-Log ('{0} problema(s) critico(s) e {1} alerta(s).' -f $crit.Count, $alt.Count)
        Write-Log ''
        $ordem = @()
        $ordem += @($crit | Where-Object { $_.Impacto -eq 'Alto' })
        $ordem += @($crit | Where-Object { $_.Impacto -ne 'Alto' })
        $ordem += @($alt  | Where-Object { $_.Impacto -eq 'Alto' })
        $ordem += @($alt  | Where-Object { $_.Impacto -ne 'Alto' })

        Write-Log 'RESOLVA NESTA ORDEM:' 'TITULO'
        $n = 0
        foreach ($a in $ordem) {
            $n++
            Write-Log ('{0}. [{1}/{2}] {3}' -f $n, $a.Categoria, $a.Impacto, $a.Titulo) $a.Severidade
            if ($a.Recomendacao) { Write-Log ('-> ' + $a.Recomendacao) 'ACAO' }
        }
    }

    Write-Log ''
    Write-Log 'Tudo o que esta acima e resolvido pelo Modulo 3 (Limpeza). Abra e clique em "Aplicar tudo".' 'ACAO'
    Write-Log 'O que precisa de administrador esta marcado como tal: use "Copiar log" no chamado.' 'DADO'
    Write-Log 'Este diagnostico ja incluiu os testes de Citrix (tres destinos) e de Tasy (tres ambientes).' 'DADO'
    Save-Inventario
}

# =====================================================================
# 6. MODULO 3 - LIMPEZA SEGURA
# =====================================================================
# Pastas de trabalho do ecossistema dentro do %TEMP%. O AUTO_CONTORNO cria o
# staging com tempfile.mkdtemp usando estes prefixos, copia a serie do paciente
# para la e o TotalSegmentator trabalha em cima - I/O pesado e continuo, por
# horas. Prefixos observados: radai_, radai_lote_, radai_prev_, radai_portal_ e
# ts_ (os tres do meio ja caem em radai_).
#
# O alvo TEMP do Modulo 3 nao tem filtro de idade e faz Remove-Item -Recurse
# -Force em tudo que encontra, inclusive pasta criada segundos antes. Apagar
# staging no meio de um lote nao da erro visivel: o cleaner anuncia os MB
# liberados e o lote falha depois, por um motivo que nao parece limpeza.
$script:PrefixosStagingEcossistema = @('radai_', 'ts_')

function Test-StagingEcossistema {
    param([string]$Nome)
    if (-not $Nome) { return $false }
    foreach ($pre in $script:PrefixosStagingEcossistema) {
        if ($Nome.StartsWith($pre, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function New-Alvo {
    param(
        [string]$Id, [string]$Nome, [string[]]$Caminhos,
        [ValidateSet('Conteudo','ArquivosRaiz','Lixeira')][string]$Modo = 'Conteudo',
        [int]$DiasMin = 0, [bool]$Padrao = $true, [string[]]$Fechar = @(), [string]$Obs = ''
    )
    [pscustomobject]@{
        Id = $Id; Nome = $Nome; Caminhos = $Caminhos; Modo = $Modo
        DiasMin = $DiasMin; Padrao = $Padrao; Fechar = $Fechar; Obs = $Obs; Bytes = -1
    }
}

function Get-AlvosLimpeza {
    $lad = $env:LOCALAPPDATA
    $ad  = $env:APPDATA
    $up  = $env:USERPROFILE
    $al  = @()
    $dias = $script:DiasCorte
    $txtPeriodo = $(if ($dias -gt 0) { 'mais de ' + $dias + ' dias' } else { 'tudo, sem limite de idade' })

    $al += New-Alvo -Id 'TEMP' -Nome 'Arquivos temporarios do usuario' -Caminhos @($env:TEMP) -Obs 'Arquivos em uso sao ignorados automaticamente.'
    $al += New-Alvo -Id 'EDGE' -Nome 'Cache do Microsoft Edge' -Fechar @('msedge') -Caminhos @(
        "$lad\Microsoft\Edge\User Data\*\Cache",
        "$lad\Microsoft\Edge\User Data\*\Code Cache",
        "$lad\Microsoft\Edge\User Data\*\GPUCache",
        "$lad\Microsoft\Edge\User Data\*\Service Worker\CacheStorage",
        "$lad\Microsoft\Edge\User Data\component_crx_cache",
        "$lad\Microsoft\Edge\User Data\GrShaderCache") -Obs 'Nao apaga favoritos, senhas nem historico.'
    $al += New-Alvo -Id 'CHROME' -Nome 'Cache do Google Chrome' -Fechar @('chrome') -Caminhos @(
        "$lad\Google\Chrome\User Data\*\Cache",
        "$lad\Google\Chrome\User Data\*\Code Cache",
        "$lad\Google\Chrome\User Data\*\GPUCache",
        "$lad\Google\Chrome\User Data\*\Service Worker\CacheStorage") -Obs 'Nao apaga favoritos, senhas nem historico.'
    $al += New-Alvo -Id 'FIREFOX' -Nome 'Cache do Firefox' -Fechar @('firefox') -Caminhos @(
        "$lad\Mozilla\Firefox\Profiles\*\cache2",
        "$lad\Mozilla\Firefox\Profiles\*\startupCache",
        "$lad\Mozilla\Firefox\Profiles\*\OfflineCache") -Obs 'Nao mexe no perfil com favoritos e senhas, que fica em Roaming.'
    $al += New-Alvo -Id 'TEAMS1' -Nome 'Cache do Teams (versao classica)' -Fechar @('Teams') -Caminhos @(
        "$ad\Microsoft\Teams\Cache", "$ad\Microsoft\Teams\blob_storage", "$ad\Microsoft\Teams\databases",
        "$ad\Microsoft\Teams\GPUCache", "$ad\Microsoft\Teams\IndexedDB", "$ad\Microsoft\Teams\Local Storage",
        "$ad\Microsoft\Teams\tmp", "$ad\Microsoft\Teams\Code Cache") -Obs 'Feche o Teams antes. As conversas ficam no servidor.'
    $al += New-Alvo -Id 'TEAMS2' -Nome 'Cache do Teams (versao nova)' -Fechar @('ms-teams') -Caminhos @(
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Cache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Code Cache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\GPUCache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Service Worker\CacheStorage",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Service Worker\ScriptCache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Cache Storage",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\PreviousVersions",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Temp")
    $al += New-Alvo -Id 'JAVAWS' -Nome 'Cache do Java Web Start (Tasy)' -Fechar @('javaw','java','jp2launcher') -Caminhos @(
        "$lad\Sun\Java\Deployment\cache", "$ad\Sun\Java\Deployment\cache",
        "$up\AppData\LocalLow\Sun\Java\Deployment\cache") -Obs 'Cache antigo do Java faz o Tasy falhar ao baixar arquivo. E baixado de novo no proximo acesso.'
    $al += New-Alvo -Id 'CENTCACHE' -Nome 'Cache do CentBrowser (navegador do prontuario)' -Padrao $false -Fechar @('CentBrowser') -Caminhos @(
        "$lad\CentBrowser\User Data\*\Cache", "$lad\CentBrowser\User Data\*\Code Cache",
        "$lad\CentBrowser\User Data\*\GPUCache") -Obs 'DESMARCADO de proposito: e o navegador do prontuario. Nao apaga favoritos nem senhas, so o cache. Marque se o Tasy estiver com tela em branco ou erro de arquivo.'
    $al += New-Alvo -Id 'RDPCACHE' -Nome 'Cache de imagens da Area de Trabalho Remota' -Caminhos @(
        "$lad\Microsoft\Terminal Server Client\Cache") -Obs 'Miniaturas de tela das sessoes remotas. Sao recriadas na proxima conexao.'
    $al += New-Alvo -Id 'CTEMP' -Nome ('Pasta C:\Temp (' + $txtPeriodo + ')') -DiasMin $dias -Caminhos @(
        'C:\Temp') -Obs 'Pasta de instalacao usada pelo TI. Arquivos em uso sao ignorados.'
    $al += New-Alvo -Id 'OPERAVIVALDI' -Nome 'Cache de Opera e Vivaldi' -Caminhos @(
        "$lad\Opera Software\Opera Next\Cache", "$lad\Opera Software\Opera Stable\Cache",
        "$lad\Vivaldi\User Data\*\Cache")
    $al += New-Alvo -Id 'INETCOOKIES' -Nome 'Cookies legados do Windows (INetCookies)' -Padrao $false -Caminhos @(
        "$lad\Microsoft\Windows\INetCookies") -Obs 'Pode pedir login de novo em sites internos antigos. Por isso vem desmarcado.'
    $al += New-Alvo -Id 'INETCACHE' -Nome 'Cache de internet do Windows' -Caminhos @("$lad\Microsoft\Windows\INetCache")
    $al += New-Alvo -Id 'WER' -Nome 'Relatorios de erro do Windows' -Caminhos @(
        "$lad\Microsoft\Windows\WER\ReportQueue", "$lad\Microsoft\Windows\WER\ReportArchive",
        "$ad\Microsoft\Windows\WER")
    $al += New-Alvo -Id 'DUMPS' -Nome 'Despejos de memoria de travamentos' -Caminhos @("$lad\CrashDumps")
    $al += New-Alvo -Id 'GPU' -Nome 'Cache de video (shaders)' -Caminhos @(
        "$lad\D3DSCache", "$lad\NVIDIA\DXCache", "$lad\NVIDIA\GLCache",
        "$lad\NVIDIA Corporation\NV_Cache", "$lad\AMD\DxCache", "$lad\Intel\ShaderCache")
    $al += New-Alvo -Id 'ONEDRIVE' -Nome 'Logs do OneDrive' -Caminhos @("$lad\Microsoft\OneDrive\logs")
    $al += New-Alvo -Id 'TEAMSANTIGO' -Nome 'Versoes antigas do Teams' -Caminhos @(
        "$lad\Microsoft\Teams\previous", "$lad\Microsoft\Teams\packages", "$lad\SquirrelTemp")
    $al += New-Alvo -Id 'EDGEUPDATE' -Nome 'Instaladores baixados do Edge/Chrome' -Caminhos @(
        "$lad\Microsoft\EdgeUpdate\Download", "$lad\Google\Update\Download")
    $al += New-Alvo -Id 'ADOBE' -Nome 'Cache do Adobe Acrobat/Reader' -Fechar @('Acrobat','AcroRd32') -Caminhos @(
        "$lad\Adobe\Acrobat\DC\Cache", "$lad\Adobe\Acrobat\DC\ConnectorIcons", "$lad\Adobe\Color\ACEC")
    $al += New-Alvo -Id 'OFFICELOG' -Nome 'Logs e diagnosticos do Office' -Caminhos @(
        "$lad\Microsoft\Office\16.0\Telemetry", "$lad\Temp\Diagnostics", "$lad\Microsoft\Office\Logs")
    # O cache do Citrix NUNCA e limpo por este kit (decisao do servico).
    $al += New-Alvo -Id 'CITRIXICA' -Nome ('Arquivos .ica soltos em Downloads (' + $txtPeriodo + ')') -Modo 'ArquivosRaiz' -DiasMin $dias -Caminhos @(
        "$up\Downloads") -Obs 'Apenas os lancadores baixados pelo navegador. O cache do Citrix nunca e tocado.'
    $al += New-Alvo -Id 'LIXEIRA' -Nome 'Lixeira' -Modo 'Lixeira' -Caminhos @()

    # itens que dependem de decisao do usuario
    $al += New-Alvo -Id 'OFFICECACHE' -Nome 'Cache de documentos do Office' -Padrao $false -Fechar @('excel','winword','powerpnt','outlook') -Caminhos @(
        "$lad\Microsoft\Office\16.0\OfficeFileCache") -Obs 'Feche todo o Office antes. Alteracoes ainda nao enviadas podem se perder.'
    $al += New-Alvo -Id 'MINIATURAS' -Nome 'Cache de miniaturas e icones' -Modo 'ArquivosRaiz' -Caminhos @(
        "$lad\Microsoft\Windows\Explorer") -Obs 'As miniaturas sao recriadas sozinhas.'
    $al += New-Alvo -Id 'EXCELRECUP' -Nome ('Autorrecuperacao do Excel (' + $txtPeriodo + ')') -Modo 'ArquivosRaiz' -DiasMin $dias -Caminhos @(
        "$ad\Microsoft\Excel") -Obs ('Nao mexe na pasta XLSTART. Com "tudo" selecionado, apaga tambem a recuperacao de hoje.')
    $al += New-Alvo -Id 'RECENTES' -Nome 'Lista de arquivos recentes' -Padrao $false -Caminhos @(
        "$ad\Microsoft\Windows\Recent") -Obs 'Some a lista de recentes do Office e do Explorer. Nenhum arquivo e apagado.'
    $al += New-Alvo -Id 'ORFAOS' -Nome ('Arquivos travados do Office na pasta clinica (' + $txtPeriodo + ')') -Modo 'ArquivosRaiz' -DiasMin $dias -Fechar @('excel','winword') -Caminhos @(
        $script:PastaClinica) -Obs 'Feche o Excel antes. Sao arquivos de bloqueio que sobraram de sessoes travadas.'

    return $al
}

function Resolve-CaminhosAlvo {
    param($Alvo)
    $lista = @()
    foreach ($c in $Alvo.Caminhos) {
        if ($c -match '\*') {
            try { $lista += (Get-Item -Path $c -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName) } catch { }
        } elseif (Test-Path -LiteralPath $c) {
            $lista += $c
        }
    }
    return ($lista | Select-Object -Unique)
}

function Measure-Alvo {
    param($Alvo)
    if ($Alvo.Modo -eq 'Lixeira') {
        try {
            $sh = New-Object -ComObject Shell.Application
            $itens = $sh.NameSpace(0x0A).Items()
            $t = 0.0
            foreach ($i in $itens) { try { $t += [double]$i.Size } catch { } }
            return $t
        } catch { return 0 }
    }
    $total = 0.0
    foreach ($c in (Resolve-CaminhosAlvo -Alvo $Alvo)) {
        if ($Alvo.Modo -eq 'ArquivosRaiz') {
            try {
                $itens = @(Get-ChildItem -LiteralPath $c -File -Force -ErrorAction SilentlyContinue)
                if ($Alvo.Id -eq 'MINIATURAS')  { $itens = @($itens | Where-Object { $_.Name -match '^(thumbcache|iconcache)' }) }
                if ($Alvo.Id -eq 'ORFAOS')      { $itens = @($itens | Where-Object { $_.Name -match '^~\$|\.tmp$' }) }
                if ($Alvo.Id -eq 'CITRIXICA')   { $itens = @($itens | Where-Object { $_.Extension -eq '.ica' }) }
                if ($Alvo.DiasMin -gt 0)        { $itens = @($itens | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Alvo.DiasMin) }) }
                foreach ($i in $itens) { $total += $i.Length }
            } catch { }
        } else {
            $b = Get-TamanhoPasta -Caminho $c -TimeoutSeg 45
            if ($b -gt 0) { $total += $b }
        }
    }
    return $total
}

function Clear-Alvo {
    param($Alvo)
    $liberado = 0.0
    $apagados = 0
    $bloqueados = 0
    $preservados = 0
    $corteFresco = (Get-Date).AddMinutes(-1 * $script:Lim.CacheFrescoMin)

    if ($Alvo.Modo -eq 'Lixeira') {
        $antes = Measure-Alvo -Alvo $Alvo
        if ($antes -le 0) { return [pscustomobject]@{ Bytes = 0; Itens = 0; Bloqueados = 0; Erro = $null } }
        try {
            Clear-RecycleBin -Force -ErrorAction Stop
            return [pscustomobject]@{ Bytes = $antes; Itens = 1; Bloqueados = 0; Erro = $null }
        } catch {
            # Lixeira ja vazia devolve um erro de token; nao e falha
            if ("$($_.Exception.Message)" -match 'token that does not exist|nao localizado|not exist|nao foi encontrado') {
                return [pscustomobject]@{ Bytes = 0; Itens = 0; Bloqueados = 0; Erro = $null }
            }
            return [pscustomobject]@{ Bytes = 0; Itens = 0; Bloqueados = 0; Erro = $_.Exception.Message }
        }
    }

    foreach ($c in (Resolve-CaminhosAlvo -Alvo $Alvo)) {
        try {
            if ($Alvo.Modo -eq 'ArquivosRaiz') {
                $itens = @(Get-ChildItem -LiteralPath $c -File -Force -ErrorAction SilentlyContinue)
                if ($Alvo.Id -eq 'MINIATURAS') { $itens = @($itens | Where-Object { $_.Name -match '^(thumbcache|iconcache)' }) }
                if ($Alvo.Id -eq 'ORFAOS')     { $itens = @($itens | Where-Object { $_.Name -match '^~\$|\.tmp$' }) }
                if ($Alvo.Id -eq 'CITRIXICA')  { $itens = @($itens | Where-Object { $_.Extension -eq '.ica' }) }
                if ($Alvo.DiasMin -gt 0)       { $itens = @($itens | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Alvo.DiasMin) }) }
            } else {
                $itens = @(Get-ChildItem -LiteralPath $c -Force -ErrorAction SilentlyContinue)
                if ($Alvo.DiasMin -gt 0)       { $itens = @($itens | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Alvo.DiasMin) }) }
            }

            # Staging do ecossistema: fora, sempre. E a pasta raiz do staging nem
            # muda de data enquanto o lote escreve no fundo dela, entao a guarda de
            # frescor abaixo nao bastaria - por isso as duas.
            $antesFiltro = $itens.Count
            $itens = @($itens | Where-Object { -not (Test-StagingEcossistema $_.Name) })
            # Mexido agora: nao se apaga o que esta sendo escrito neste instante.
            if ($Alvo.DiasMin -le 0) {
                $itens = @($itens | Where-Object { $_.LastWriteTime -lt $corteFresco })
            }
            $preservados += ($antesFiltro - $itens.Count)
            $n = 0
            foreach ($i in $itens) {
                if ($script:Cancelar) { break }
                $tam = 0.0
                try {
                    if ($i.PSIsContainer) { $tam = Get-TamanhoPastaNet -Caminho $i.FullName } else { $tam = $i.Length }
                } catch { }
                try {
                    Remove-Item -LiteralPath $i.FullName -Recurse -Force -ErrorAction Stop
                    $liberado += $tam
                    $apagados++
                } catch { $bloqueados++ }
                $n++
                if ($n % 25 -eq 0) { Pump }
            }
        } catch { }
    }
    return [pscustomobject]@{ Bytes = $liberado; Itens = $apagados; Bloqueados = $bloqueados; Preservados = $preservados; Erro = $null }
}

function Get-PlanoLimpeza {
    $plano = New-Object System.Collections.ArrayList

    # ---- 1. arquivos, caches e temporarios ----
    foreach ($a in (Get-AlvosLimpeza)) {
        if ($script:Cancelar) { break }
        if ($a.Modo -eq 'Lixeira') { continue }
        Set-Status ('Medindo ' + $a.Nome + '...')
        $b = Measure-Alvo -Alvo $a
        if ($b -le 0) { continue }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'LIMPAR'; Rotulo = $a.Nome; Valor = (Format-Bytes $b); Bytes = $b
            Marcar = [bool]$a.Padrao; Dados = $a; Nota = $a.Obs
        })
    }

    # ---- 2. lixeira (rotina) ----
    $lix = New-Alvo -Id 'LIXEIRA' -Nome 'Esvaziar a Lixeira' -Modo 'Lixeira' -Caminhos @()
    Set-Status 'Medindo a Lixeira...'
    $bl = Measure-Alvo -Alvo $lix
    [void]$plano.Add([pscustomobject]@{
        Grupo = 'LIXEIRA'; Rotulo = 'Esvaziar a Lixeira (rotina)'; Valor = (Format-Bytes $bl); Bytes = $bl
        Marcar = $true; Dados = $lix; Nota = 'Feito antes dos Downloads, para que o que sair de la ainda possa ser recuperado.'
    })

    # ---- 3. downloads ----
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (Test-Path -LiteralPath $dl) {
        $dias = $script:DiasCorte
        $itens = @(Get-ChildItem -LiteralPath $dl -Force -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -notlike 'Inventario_*' -and $_.Extension -ne '.ica' })
        if ($dias -gt 0) { $itens = @($itens | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$dias) }) }
        $bd = 0.0
        foreach ($it in $itens) {
            try { if ($it.PSIsContainer) { $bd += Get-TamanhoPastaNet -Caminho $it.FullName } else { $bd += $it.Length } } catch { }
        }
        $rot = if ($dias -gt 0) { ('Downloads com mais de {0} dias ({1} itens)' -f $dias, $itens.Count) } else { ('Downloads - tudo ({0} itens)' -f $itens.Count) }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'BAIXADOS'; Rotulo = $rot; Valor = (Format-Bytes $bd); Bytes = $bd
            Marcar = ($itens.Count -gt 0); Dados = $itens; Nota = 'Vai para a Lixeira, da para recuperar. Troque o periodo na caixa acima da lista.'
        })
    }

    # ---- 4. encerrar agora ----
    foreach ($e in (Get-ProcessosEncerraveis)) {
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'FECHAR'; Rotulo = ('{0} ({1} processo(s))' -f $e.Rotulo, $e.Qtd); Valor = (Format-Bytes $e.Memoria); Bytes = 0
            Marcar = $true; Dados = $e; Nota = $e.Nota
        })
    }

    # ---- 5. nao carregar na proxima inicializacao ----
    foreach ($i in (Get-ItensInicializacao)) {
        if ($i.Classe -eq 'Protegido') { continue }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'INICIAR'; Rotulo = ('{0} - {1}' -f $i.Nome, $i.Rotulo); Valor = ''; Bytes = 0
            Marcar = [bool]$i.Seguro; Dados = $i
            Nota = $(if ($i.Seguro) { 'Continua funcionando quando voce abrir pelo menu Iniciar.' } else { 'Desmarcado por seguranca: o kit nao reconheceu este item.' })
        })
    }

    # ---- 6. ajustes de desempenho ----
    foreach ($a in (Get-AjustesPendentes)) {
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'AJUSTE'; Rotulo = $a.Rotulo; Valor = ''; Bytes = 0
            Marcar = $true; Dados = $a; Nota = $a.Nota
        })
    }

    # ---- 7. mapeamentos mortos ----
    foreach ($m in (Get-MapeamentosMortos)) {
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'REDE'; Rotulo = ('Remover mapeamento morto {0} ({1})' -f $m.Letra, $m.Destino); Valor = ''; Bytes = 0
            Marcar = $true; Dados = $m; Nota = 'Rode o Modulo 1 (Preparar ambiente) depois para remapear.'
        })
    }

    # ---- 8. credenciais orfas deixadas pela migracao de dominio ----
    foreach ($c in (Get-CredenciaisOrfas)) {
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'CREDENC'; Rotulo = ('Remover credencial salva de {0}' -f $c.Servidor); Valor = ''; Bytes = 0
            Marcar = $false; Dados = $c
            Nota = 'DESMARCADO de proposito: e a unica acao do kit sem desfazer. Se o DNS estiver instavel, pode remover credencial valida.'
        })
    }

    return $plano
}


function Invoke-Modulo3Analise {
    Clear-SnapshotsOperacao
    Write-Titulo 'Modulo 3 - Limpeza (analise)'
    Write-Log 'Nada e alterado nesta etapa. Tudo o que e seguro ja vem marcado.' 'DADO'

    $plano = Get-PlanoLimpeza
    $script:ModoPlano = 'LIMPEZA'
    $script:PlanoAtual = $plano
    $script:ListaLimpeza.Items.Clear()

    $ordemGrupo = @{ 'LIMPAR' = 1; 'LIXEIRA' = 2; 'BAIXADOS' = 3; 'FECHAR' = 4; 'INICIAR' = 5; 'AJUSTE' = 6; 'REDE' = 7; 'CREDENC' = 8 }
    $plano = @($plano | Sort-Object @{ Expression = { $ordemGrupo[$_.Grupo] } }, @{ Expression = { -$_.Bytes } })
    $script:PlanoAtual = $plano

    foreach ($p in $plano) {
        $rotulo = '[{0,-8}] {1,-52} {2,10}' -f $p.Grupo, $p.Rotulo, $p.Valor
        [void]$script:ListaLimpeza.Items.Add($rotulo)
        $script:ListaLimpeza.SetItemChecked($script:ListaLimpeza.Items.Count - 1, $p.Marcar)
    }

    # resumo no log
    $bytes = ($plano | Where-Object { $_.Marcar -and $_.Grupo -match 'LIMPAR|LIXEIRA|BAIXADOS' } | Measure-Object Bytes -Sum).Sum
    $mem   = ($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'FECHAR' } | ForEach-Object { $_.Dados.Memoria } | Measure-Object -Sum).Sum
    $ini   = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'INICIAR' }).Count
    $aju   = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'AJUSTE' }).Count
    $red   = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'REDE' }).Count
    $naoRec = @($plano | Where-Object { $_.Grupo -eq 'INICIAR' -and -not $_.Marcar })

    Write-Log '' 'DADO'
    Write-Log 'O que vai acontecer ao clicar em "Aplicar tudo":' 'TITULO'
    Write-Log ('Espaco liberado em disco ...... {0}' -f (Format-Bytes $bytes)) 'ACAO'
    Write-Log ('Memoria devolvida agora ....... {0}' -f (Format-Bytes $mem)) 'ACAO'
    Write-Log ('Itens fora da inicializacao ... {0}' -f $ini) 'ACAO'
    Write-Log ('Ajustes de desempenho ......... {0}' -f $aju) 'ACAO'
    if ($red -gt 0) { Write-Log ('Mapeamentos mortos removidos .. {0}' -f $red) 'ACAO' }
    if ($naoRec.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ('{0} item(ns) de inicializacao ficaram desmarcados por nao serem reconhecidos:' -f $naoRec.Count) 'DADO'
        foreach ($n in $naoRec) { Write-Log ('   ' + $n.Rotulo) 'DADO' }
        Write-Log 'Marque na lista se souber que pode desativar.' 'DADO'
    }
    Write-Log '' 'DADO'
    Write-Log 'Desmarque o que nao quiser e clique em "Aplicar tudo".' 'ACAO'
    $script:PainelGrandes.Visible = $false
    $script:PainelSessao.Visible = $false
    $script:PainelLimpeza.Visible = $true
}

function Invoke-Modulo3Limpeza {
    if (-not $script:PlanoAtual -or $script:PlanoAtual.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Rode primeiro o Modulo 3 para analisar a maquina.', 'Limpeza', 'OK', 'Information') | Out-Null
        return
    }

    $marcados = @()
    for ($i = 0; $i -lt $script:ListaLimpeza.Items.Count; $i++) {
        if ($script:ListaLimpeza.GetItemChecked($i)) { $marcados += $script:PlanoAtual[$i] }
    }
    if ($marcados.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Nenhum item marcado.', 'Limpeza', 'OK', 'Information') | Out-Null
        return
    }

    $bytes = ($marcados | Where-Object { $_.Grupo -match 'LIMPAR|LIXEIRA|BAIXADOS' } | Measure-Object Bytes -Sum).Sum
    $nFechar = @($marcados | Where-Object { $_.Grupo -eq 'FECHAR' }).Count
    $nIni    = @($marcados | Where-Object { $_.Grupo -eq 'INICIAR' }).Count
    $nAju    = @($marcados | Where-Object { $_.Grupo -eq 'AJUSTE' }).Count

    if ($script:ModoPlano -eq 'SESSAO') {
        $nSes = @($marcados | Where-Object { $_.Grupo -eq 'SESSAO' }).Count
        $nTar = @($marcados | Where-Object { $_.Grupo -eq 'TAREFA' }).Count
        $memSes = ($marcados | Where-Object { $_.Grupo -eq 'SESSAO' } | Measure-Object Bytes -Sum).Sum
        $msg = "Otimizar a sessao de agora com {0} itens:`r`n`r`n" -f $marcados.Count
        $msg += " · Encerrar {0} programa(s), devolvendo cerca de {1}`r`n" -f $nSes, (Format-Bytes $memSes)
        $msg += " · Parar {0} tarefa(s) agendada(s) em execucao`r`n" -f $nTar
        $msg += " · Ajustar prioridade de CPU e compactar memoria`r`n`r`n"
        $msg += "Citrix, navegador do prontuario, navegador web e Office nao sao tocados.`r`n"
        $msg += "Nada e desinstalado: o proximo logon devolve tudo ao normal.`r`n`r`nContinuar?"
    } else {
        $msg = "Vao ser aplicados {0} itens:`r`n`r`n" -f $marcados.Count
        $msg += " · Liberar cerca de {0} em disco`r`n" -f (Format-Bytes $bytes)
        $msg += " · Encerrar {0} programa(s) dispensavel(is)`r`n" -f $nFechar
        $msg += " · Tirar {0} item(ns) da inicializacao`r`n" -f $nIni
        $msg += " · Aplicar {0} ajuste(s) de desempenho`r`n`r`n" -f $nAju
        $msg += "Downloads e Lixeira: o que sair vai para a Lixeira antes dela ser esvaziada, entao ainda da para recuperar.`r`n`r`nContinuar?"
    }

    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Aplicar limpeza', 'YesNo', 'Question') -ne 'Yes') {
        Write-Log 'Limpeza cancelada pelo usuario.' 'ALERTA'
        return
    }

    # Instantaneo DEPOIS do 'Sim', nunca antes. A caixa de confirmacao espera o
    # usuario por tempo indefinido, e e tempo em que ele pode alternar para o
    # Local Suite e disparar um lote - que e o motivo de otimizar a sessao. Um
    # retrato tirado antes do dialogo chegaria ao executor com a idade da
    # hesitacao do usuario, e nada aqui mede essa idade.
    Clear-SnapshotsOperacao
    [void](Get-EmUsoOperacao -Renovar)

    Set-Ocupado $true
    $script:Cancelar = $false
    $script:SessaoEncerrados = @()
    $script:ExplorerAtualizado = $false
$script:LogonSegundos      = 0
    Write-Titulo $(if ($script:ModoPlano -eq 'SESSAO') { 'Modulo 5 - Aplicando na sessao' } else { 'Modulo 3 - Aplicando' })

    $liberado = 0.0
    $memoria  = 0.0
    $desativados = @()
    $n = 0

    try {
        # ordem: fechar (libera arquivos travados) -> limpar -> lixeira -> downloads -> inicializacao -> ajustes -> rede
        $ordem = @('SESSAO', 'TAREFA', 'FECHAR', 'LIMPAR', 'LIXEIRA', 'BAIXADOS', 'INICIAR', 'AJUSTE', 'REDE', 'CREDENC', 'PRIORIDADE', 'MEMORIA')
        foreach ($g in $ordem) {
            $doGrupo = @($marcados | Where-Object { $_.Grupo -eq $g })
            if ($doGrupo.Count -eq 0) { continue }

            switch ($g) {
                'FECHAR'   { Write-Log '' ; Write-Log 'Encerrando programas dispensaveis...' 'TITULO' }
                'LIMPAR'   { Write-Log '' ; Write-Log 'Limpando caches e temporarios...' 'TITULO' }
                'LIXEIRA'  { Write-Log '' ; Write-Log 'Esvaziando a Lixeira...' 'TITULO' }
                'BAIXADOS' { Write-Log '' ; Write-Log 'Enviando Downloads antigos para a Lixeira...' 'TITULO' }
                'INICIAR'  { Write-Log '' ; Write-Log 'Tirando programas da inicializacao...' 'TITULO' }
                'AJUSTE'   { Write-Log '' ; Write-Log 'Aplicando ajustes de desempenho...' 'TITULO' }
                'REDE'     { Write-Log '' ; Write-Log 'Removendo mapeamentos mortos...' 'TITULO' }
                'CREDENC'  { Write-Log '' ; Write-Log 'Removendo credenciais de servidores inexistentes...' 'TITULO' }
                'SESSAO'     { Write-Log '' ; Write-Log 'Encerrando programas dispensaveis da sessao...' 'TITULO' }
                'TAREFA'     { Write-Log '' ; Write-Log 'Parando tarefas agendadas em execucao...' 'TITULO' }
                'PRIORIDADE' { Write-Log '' ; Write-Log 'Ajustando prioridade de CPU...' 'TITULO' }
                'MEMORIA'    { Write-Log '' ; Write-Log 'Compactando a memoria dos programas que ficam...' 'TITULO' }
            }

            foreach ($item in $doGrupo) {
                if ($script:Cancelar) { Write-Log 'Interrompido pelo usuario.' 'ALERTA'; break }
                $n++
                Set-Status ('Aplicando {0}/{1}: {2}' -f $n, $marcados.Count, $item.Rotulo)

                switch ($item.Grupo) {

                    'FECHAR' {
                        $antes = 0.0
                        $ok = 0
                        foreach ($id in $item.Dados.Ids) {
                            try {
                                $p = Get-Process -Id $id -ErrorAction Stop
                                if ($p.ProcessName -match $script:Protegidos) { continue }
                                $antes += $p.WorkingSet64
                                if ($p.MainWindowHandle -ne 0) { [void]$p.CloseMainWindow(); Start-Sleep -Milliseconds 500 }
                                if (-not $p.HasExited) { Stop-Process -Id $id -Force -ErrorAction Stop }
                                $ok++
                            } catch { }
                        }
                        $memoria += $antes
                        if ($ok -gt 0) {
                            $script:SessaoEncerrados += @($item.Dados.Nomes)
                            Write-Log ('{0}: {1} processo(s) encerrado(s), {2} devolvidos' -f $item.Dados.Rotulo, $ok, (Format-Bytes $antes)) 'OK'
                        }
                        else { Write-Log ('{0}: ja nao estava em execucao' -f $item.Dados.Rotulo) 'DADO' }
                    }

                    'LIMPAR' {
                        $r = Clear-Alvo -Alvo $item.Dados
                        $liberado += $r.Bytes
                        if ($r.Erro) { Write-Log ('{0}: {1}' -f $item.Rotulo, $r.Erro) 'ALERTA' }
                        elseif ($r.Bytes -le 0) { Write-Log ('{0}: nada a limpar' -f $item.Rotulo) 'DADO' }
                        else {
                            $partes = @()
                            if ($r.Bloqueados -gt 0)  { $partes += ('{0} em uso' -f $r.Bloqueados) }
                            if ($r.Preservados -gt 0) { $partes += ('{0} preservado(s): trabalho do ecossistema ou mexido agora' -f $r.Preservados) }
                            $extra = if ($partes.Count -gt 0) { (' (' + ($partes -join '; ') + ')') } else { '' }
                            Write-Log ('{0}: {1} liberados{2}' -f $item.Rotulo, (Format-Bytes $r.Bytes), $extra) 'OK'
                        }
                    }

                    'LIXEIRA' {
                        $r = Clear-Alvo -Alvo $item.Dados
                        $liberado += $r.Bytes
                        if ($r.Erro) { Write-Log ('Lixeira: {0}' -f $r.Erro) 'DADO' }
                        elseif ($r.Bytes -le 0) { Write-Log 'Lixeira: ja estava vazia' 'DADO' }
                        else { Write-Log ('Lixeira esvaziada: {0} liberados' -f (Format-Bytes $r.Bytes)) 'OK' }
                    }

                    'BAIXADOS' {
                        try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop } catch { }
                        $lib = 0.0; $qtd = 0; $sumidos = 0
                        $falhas = New-Object System.Collections.ArrayList
                        foreach ($f in $item.Dados) {
                            if ($script:Cancelar) { break }
                            if (-not (Test-Path -LiteralPath $f.FullName)) { $sumidos++; continue }
                            $tam = 0.0
                            try { if ($f.PSIsContainer) { $tam = Get-TamanhoPastaNet -Caminho $f.FullName } else { $tam = $f.Length } } catch { }
                            try {
                                if ($f.PSIsContainer) { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($f.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin') }
                                else                  { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin') }
                                $lib += $tam; $qtd++
                            } catch { [void]$falhas.Add($f.Name) }
                            if ($qtd % 10 -eq 0) { Pump }
                        }
                        $liberado += $lib
                        Write-Log ('Downloads: {0} itens para a Lixeira, {1} liberados' -f $qtd, (Format-Bytes $lib)) 'OK'
                        if ($sumidos -gt 0) { Write-Log ('{0} item(ns) ja tinham sido removidos por outra etapa.' -f $sumidos) 'DADO' }
                        if ($falhas.Count -gt 0) {
                            Write-Log ('{0} item(ns) em uso foram mantidos. Exemplos:' -f $falhas.Count) 'ALERTA'
                            foreach ($n in ($falhas | Select-Object -First 5)) { Write-Log ('   ' + $n) 'DADO' }
                            if ($falhas.Count -gt 5) { Write-Log ('   ... e outros {0}.' -f ($falhas.Count - 5)) 'DADO' }
                        }
                        Write-Log 'Confira a Lixeira antes de esvaziar de novo, caso queira recuperar algo.' 'ACAO'
                    }

                    'INICIAR' {
                        if ($item.Dados.Nome -match 'OneDrive' -and (Test-PastasNoOneDrive)) {
                            Write-Log 'REGISTRO: a Area de Trabalho ou Documentos deste usuario estao dentro do OneDrive.' 'ALERTA'
                            Write-Log 'Com o OneDrive fora da inicializacao, os arquivos so sobem para a nuvem quando ele for aberto.' 'ALERTA'
                        }
                        if (Disable-ItemInicializacao -Nome $item.Dados.Nome -Tipo $item.Dados.Tipo) {
                            Write-Log ('Nao vai mais abrir sozinho: {0}' -f $item.Dados.Nome) 'OK'
                            $desativados += @{ Nome = $item.Dados.Nome; Tipo = $item.Dados.Tipo }
                        } else {
                            Write-Log ('Nao foi possivel desativar {0}' -f $item.Dados.Nome) 'ALERTA'
                        }
                    }

                    'AJUSTE' {
                        switch ($item.Dados.Id) {
                            'EFEITOS'      { Set-AjusteEfeitos }
                            'SEGUNDOPLANO' { Set-AjusteSegundoPlano }
                            'STORAGE'      { Set-AjusteStorageSense }
                            'SUGESTOES'    { Set-AjusteSugestoes }
                            'CITRIXZONA'   { Set-AjusteZonasCitrix }
                            'BARRATAREFAS' { Set-AjusteBarraTarefas }
                            'ONEDRIVEICONE'{ Set-AjusteIconeOneDrive }
                        }
                        Write-Log ('Aplicado: {0}' -f $item.Rotulo) 'OK'
                    }

                    'REDE' {
                        # Programa externo nao lanca excecao: o catch nunca dispararia e a
                        # mensagem de sucesso sairia mesmo com o net use falhando. Confere o
                        # codigo de saida e depois se a letra sumiu da lista do Windows.
                        # A conferencia e por CIM, consulta local: nao toca no servidor morto.
                        $saidaNet = (& net use $item.Dados.Letra /delete /y 2>&1 | Out-String).Trim()
                        $codigoNet = $LASTEXITCODE
                        $aindaMapeada = $false
                        try {
                            $aindaMapeada = @(Get-CimInstance Win32_NetworkConnection -ErrorAction Stop |
                                Where-Object { "$($_.LocalName)" -eq "$($item.Dados.Letra)" }).Count -gt 0
                        } catch { }
                        if ($codigoNet -eq 0 -and -not $aindaMapeada) {
                            Write-Log ('Mapeamento {0} removido.' -f $item.Dados.Letra) 'OK'
                        } else {
                            Write-Log ('{0}: o Windows nao removeu o mapeamento (codigo {1}). A letra continua na lista.' -f $item.Dados.Letra, $codigoNet) 'ALERTA'
                            if ($saidaNet) { Write-Log ('   ' + (@($saidaNet -split "`r?`n")[0])) 'DADO' }
                        }
                    }

                    'SESSAO' {
                        Set-Status ('Encerrando {0}...' -f $item.Dados.Nome)
                        $liberou = 0.0
                        $ok = 0
                        foreach ($id in $item.Dados.Ids) {
                            try {
                                $pr = Get-Process -Id $id -ErrorAction Stop
                                # Caminho AO VIVO, de proposito: se este PID foi reciclado
                                # desde o instantaneo, o ocupante de agora e que decide.
                                # E a sinalizacao do Local Suite pela uniao das duas
                                # leituras: o arquivo relido agora, 9 ms, mais a
                                # descendencia, renovada a cada 15 s de lote.
                                if (-not (Test-PodeEncerrarSessao -Nome $pr.ProcessName -Caminho ([string]$pr.Path) -ProcId ([int]$pr.Id) -EmUso (Get-EmUsoOperacao))) { continue }
                                $liberou += $pr.WorkingSet64
                                if ($pr.MainWindowHandle -ne 0) { [void]$pr.CloseMainWindow(); Start-Sleep -Milliseconds 400 }
                                if (-not $pr.HasExited) { Stop-Process -Id $id -Force -ErrorAction Stop }
                                $ok++
                            } catch { }
                        }
                        $memoria += $liberou
                        if ($ok -gt 0) {
                            $script:SessaoEncerrados += $item.Dados.Nome
                            Write-Log ('{0}: {1} processo(s) encerrado(s), {2} devolvidos' -f $item.Dados.Nome, $ok, (Format-Bytes $liberou)) 'OK'
                        }
                        else { Write-Log ('{0}: nao foi possivel encerrar (protegido ou ja fechado)' -f $item.Dados.Nome) 'DADO' }
                    }

                    'TAREFA' {
                        try {
                            Stop-ScheduledTask -TaskName $item.Dados.Nome -TaskPath $item.Dados.Caminho -ErrorAction Stop
                            Write-Log ('Tarefa parada: {0}{1}' -f $item.Dados.Caminho, $item.Dados.Nome) 'OK'
                        } catch {
                            Write-Log ('{0}: sem permissao para parar (roda como SISTEMA - so o TI)' -f $item.Dados.Nome) 'DADO'
                        }
                    }

                    'PRIORIDADE' { Invoke-AcaoPrioridade }

                    'MEMORIA'    { Invoke-AcaoMemoria }

                    'CREDENC' {
                        # E a unica acao sem desfazer do kit, entao e a que mais precisa
                        # dizer a verdade. cmdkey e programa externo: nao lanca excecao,
                        # e o codigo de saida e o unico sinal confiavel de que removeu.
                        $saidaCred = (& cmdkey /delete:$($item.Dados.Alvo) 2>&1 | Out-String).Trim()
                        $codigoCred = $LASTEXITCODE
                        if ($codigoCred -eq 0) {
                            Write-Log ('Credencial removida (sem desfazer): {0}' -f $item.Dados.Alvo) 'OK'
                        } else {
                            Write-Log ('{0}: a credencial NAO foi removida (codigo {1}). Nada foi perdido.' -f $item.Dados.Alvo, $codigoCred) 'ALERTA'
                            if ($saidaCred) { Write-Log ('   ' + (@($saidaCred -split "`r?`n")[0])) 'DADO' }
                        }
                    }
                }
                Pump
            }
        }

        if ($desativados.Count -gt 0) {
            $e = Read-Estado
            $atual = @()
            if ($e.ContainsKey('inicializacao_desativada')) { $atual = @($e['inicializacao_desativada']) }
            Save-Estado 'inicializacao_desativada' ($atual + $desativados)
        }

        # ---------------- relatorio ----------------
        Write-Titulo 'Resultado'
        if ($script:ModoPlano -eq 'SESSAO') {
            Write-Log ('Memoria devolvida ............. {0}' -f (Format-Bytes $memoria)) 'OK'
            if ($script:UltimoGanhoMemoria -gt 0) {
                Write-Log ('Memoria compactada ............ {0}' -f (Format-Bytes $script:UltimoGanhoMemoria)) 'OK'
                Write-Log ('Total devolvido ............... {0}' -f (Format-Bytes ($memoria + $script:UltimoGanhoMemoria))) 'OK'
            }
            try {
                $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
                $pct = [Math]::Round(((($so.TotalVisibleMemorySize - $so.FreePhysicalMemory) / $so.TotalVisibleMemorySize) * 100), 1)
                Write-Log ('Memoria em uso agora .......... {0}% ({1} livres)' -f $pct, (Format-Bytes ($so.FreePhysicalMemory * 1KB))) 'DADO'
            } catch { }
            Write-Log ''
            if ($script:SessaoEncerrados.Count -gt 0) {
                $ant = @()
                $ee = Read-Estado
                if ($ee.ContainsKey('sessao_encerrados')) { $ant = @($ee['sessao_encerrados']) }
                Save-Estado 'sessao_encerrados' (@($ant + $script:SessaoEncerrados) | Select-Object -Unique)
            }
            $script:StatusSessao = [pscustomobject]@{
                Hora          = (Get-Date)
                Encerrados    = @($script:SessaoEncerrados)
                MemDevolvida  = $memoria
                MemCompactada = $script:UltimoGanhoMemoria
            }
            Write-Log ''
            Write-Log 'A sessao de agora esta mais leve. O painel a direita mostra o status.' 'DADO'
            Write-Log 'Para voltar agora, sem reiniciar: botao "REVERTER PARA O ORIGINAL".' 'ACAO'
            Show-StatusSessao -Estado 'OTIMIZADA'
            return
        }
        if ($script:SessaoEncerrados.Count -gt 0) {
            $eL = Read-Estado
            $antL = @()
            if ($eL.ContainsKey('sessao_encerrados')) { $antL = @($eL['sessao_encerrados']) }
            Save-Estado 'sessao_encerrados' (@($antL + $script:SessaoEncerrados) | Select-Object -Unique)
        }
        Write-Log ('Espaco liberado ............... {0}' -f (Format-Bytes $liberado)) 'OK'
        Write-Log ('Memoria devolvida ............. {0}' -f (Format-Bytes $memoria)) 'OK'
        Write-Log ('Itens fora da inicializacao ... {0}' -f $desativados.Count) 'OK'
        try {
            $c = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
            Write-Log ('Livre em C: agora ............. {0} ({1}%)' -f (Format-Bytes $c.FreeSpace), [Math]::Round(($c.FreeSpace / $c.Size) * 100, 1)) 'DADO'
        } catch { }
        Write-Log ''
        Write-Log 'Tudo o que foi desativado volta pelo botao "Desfazer otimizacoes".' 'DADO'

        $r = Show-Escolha -Titulo 'Reiniciar' `
            -Mensagem "Limpeza aplicada.`r`n`r`nO ganho na inicializacao e nos ajustes so aparece depois de reiniciar.`r`n`r`nQuer reiniciar agora? Salve seus arquivos antes." `
            -Opcoes @('Reiniciar agora', 'Reiniciar depois')
        if ($r -eq 0) {
            Write-Log 'Reinicio agendado para 15 segundos. Use shutdown /a no Executar para cancelar.' 'ALERTA'
            & shutdown /r /t 15 /c "Reinicio solicitado pelo Kit de Suporte" | Out-Null
        }
    } catch {
        Write-Log ('Erro durante a limpeza: {0}' -f $_.Exception.Message) 'CRITICO'
    } finally {
        Set-Ocupado $false
        Set-Status 'Limpeza concluida.'
        $script:Cancelar = $false
    }
}

# =====================================================================
# 7. ACOES RAPIDAS
# =====================================================================



function Restart-Computador {
    $msg = "O computador sera reiniciado agora.`r`n`r`nSalve os arquivos abertos antes de continuar.`r`n`r`nReiniciar (e nao Desligar) e o que realmente limpa a memoria.`r`n`r`nConfirmar?"
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Reiniciar o computador', 'YesNo', 'Warning') -eq 'Yes') {
        Write-Log 'Reiniciando o computador...' 'ALERTA'
        # shutdown e programa externo: se a politica da estacao negar o privilegio de
        # desligamento, ele falha em silencio e o aviso abaixo seria mentira.
        $saidaSd = (& shutdown /r /t 10 /c "Reinicio solicitado pelo Kit de Suporte" 2>&1 | Out-String).Trim()
        if ($LASTEXITCODE -eq 0) {
            Write-Log 'Reinicio agendado para 10 segundos. Use shutdown /a para cancelar.' 'DADO'
        } else {
            Write-Log ('O Windows nao aceitou o pedido de reinicio (codigo {0}). A maquina NAO vai reiniciar.' -f $LASTEXITCODE) 'ALERTA'
            if ($saidaSd) { Write-Log ('   ' + (@($saidaSd -split "`r?`n")[0])) 'DADO' }
            Write-Log 'Reinicie pelo menu Iniciar. Use "Reiniciar", nao "Desligar".' 'ACAO'
        }
    }
}

# =====================================================================
# 7B. ESTADO PERSISTENTE (para desfazer otimizacoes)
# =====================================================================
$script:PastaEstado = Join-Path $env:LOCALAPPDATA 'KitSuporteRT'

function Get-ArquivoEstado {
    if (-not (Test-Path -LiteralPath $script:PastaEstado)) {
        New-Item -ItemType Directory -Path $script:PastaEstado -Force | Out-Null
    }
    return (Join-Path $script:PastaEstado 'estado.json')
}

function Read-Estado {
    $f = Get-ArquivoEstado
    if (-not (Test-Path -LiteralPath $f)) { return @{} }
    try {
        $o = Get-Content -LiteralPath $f -Raw -Encoding UTF8 | ConvertFrom-Json
        $h = @{}
        foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    } catch { return @{} }
}

function Write-Estado {
    param([hashtable]$Estado)
    try { ($Estado | ConvertTo-Json -Depth 8) | Out-File -FilePath (Get-ArquivoEstado) -Encoding UTF8 -Force } catch { }
}

function Save-Estado {
    param([string]$Chave, $Valor)
    $e = Read-Estado
    $e[$Chave] = $Valor
    Write-Estado $e
}

# =====================================================================
# 7C. DIALOGOS AUXILIARES
# =====================================================================
function Show-Escolha {
    param([string]$Titulo, [string]$Mensagem, [string[]]$Opcoes)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Titulo; $f.FormBorderStyle = 'FixedDialog'; $f.StartPosition = 'CenterParent'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.BackColor = $script:Cor.Painel; $f.ForeColor = $script:Cor.Texto
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    # a altura da mensagem e medida, nao chutada: texto longo nao fica cortado
    $fonteMsg = New-Object System.Drawing.Font('Segoe UI', 9)
    $alturaMsg = 40
    try {
        $medida = [System.Windows.Forms.TextRenderer]::MeasureText(
            $Mensagem, $fonteMsg,
            (New-Object System.Drawing.Size(480, 2000)),
            ([System.Windows.Forms.TextFormatFlags]::WordBreak))
        $alturaMsg = [Math]::Max(40, $medida.Height + 8)
    } catch { $alturaMsg = 40 + (18 * [Math]::Ceiling($Mensagem.Length / 70)) }

    $f.ClientSize = New-Object System.Drawing.Size(520, (18 + $alturaMsg + 22 + ($Opcoes.Count * 44) + 14))

    $l = New-Object System.Windows.Forms.Label
    $l.AutoSize = $false
    $l.Text = $Mensagem; $l.SetBounds(20, 18, 480, $alturaMsg); $l.ForeColor = $script:Cor.Texto
    $l.Font = $fonteMsg
    $f.Controls.Add($l)

    $script:EscolhaResultado = -1
    $y = 18 + $alturaMsg + 22
    for ($i = 0; $i -lt $Opcoes.Count; $i++) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $Opcoes[$i]; $b.SetBounds(20, $y, 480, 36); $b.FlatStyle = 'Flat'
        $b.BackColor = $script:Cor.Botao; $b.ForeColor = $script:Cor.Texto
        $b.TextAlign = 'MiddleLeft'; $b.Padding = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
        $b.FlatAppearance.BorderColor = $script:Cor.Borda
        $idx = $i
        $b.Tag = $i
        $b.Add_Click({ $script:EscolhaResultado = $idx; $f.Close() }.GetNewClosure())
        $f.Controls.Add($b)
        $y += 44
    }
    [void]$f.ShowDialog()
    return $script:EscolhaResultado
}


# =====================================================================
# 7D. GRUPO A - ROTINAS CRITICAS DE IDENTIFICACAO
# =====================================================================













# C11 ----------------------------------------------------------------

# C12 ----------------------------------------------------------------
function Invoke-DiagDownloads {
    Write-Titulo 'Pasta Downloads'
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (-not (Test-Path -LiteralPath $dl)) { Write-Log 'Pasta Downloads nao encontrada.' 'DADO'; return }

    try {
        $arq = @(Get-ChildItem -LiteralPath $dl -File -Force -Recurse -ErrorAction SilentlyContinue)
        if ($arq.Count -eq 0) { Add-Achado 'OK' 'Pasta Downloads vazia' '' 'Disco'; return }

        $total  = ($arq | Measure-Object Length -Sum).Sum
        $velhos = @($arq | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) })
        $bVelho = ($velhos | Measure-Object Length -Sum).Sum

        Write-Log ('{0} arquivos · {1} no total' -f $arq.Count, (Format-Bytes $total)) 'DADO'
        Write-Log ('{0} arquivos com mais de 90 dias · {1}' -f $velhos.Count, (Format-Bytes $bVelho)) 'DADO'

        $porTipo = $arq | Group-Object Extension | ForEach-Object {
            [pscustomobject]@{ Tipo = $_.Name; Qtd = $_.Count; Bytes = ($_.Group | Measure-Object Length -Sum).Sum }
        } | Sort-Object Bytes -Descending | Select-Object -First 8
        foreach ($t in $porTipo) {
            Write-Log ('{0,-10} {1,10}  ({2} arquivo(s))' -f $(if ($t.Tipo) { $t.Tipo } else { 'sem ext' }), (Format-Bytes $t.Bytes), $t.Qtd) 'DADO'
        }

        Write-Log 'Os 10 maiores:' 'DADO'
        foreach ($f in ($arq | Sort-Object Length -Descending | Select-Object -First 10)) {
            Write-Log ('{0,10}  {1}  ({2:dd/MM/yyyy})' -f (Format-Bytes $f.Length), $f.Name, $f.LastWriteTime) 'DADO'
        }

        if ($total -gt 5GB -or $bVelho -gt 2GB) {
            Add-Achado 'ALERTA' ('Downloads ocupando {0}, sendo {1} com mais de 90 dias' -f (Format-Bytes $total), (Format-Bytes $bVelho)) 'Use o botao "Esvaziar Downloads" no menu lateral. Vai para a Lixeira, da para recuperar.' 'Disco' 'Alto'
        } else {
            Add-Achado 'OK' ('Downloads com {0}' -f (Format-Bytes $total)) '' 'Disco'
        }
    } catch { }
}


# C14 ----------------------------------------------------------------
function Invoke-DiagNuvem {
    Write-Titulo 'OneDrive e Teams'

    $od = @(Get-Process OneDrive -ErrorAction SilentlyContinue)
    if ($od.Count -gt 0) {
        $mem = ($od | Measure-Object WorkingSet64 -Sum).Sum
        Write-Log ('OneDrive rodando · {0}' -f (Format-Bytes $mem)) 'DADO'
        $pastaOD = Get-ValorReg 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' 'UserFolder'
        if (-not $pastaOD) { $pastaOD = $env:OneDrive }
        if ($pastaOD -and (Test-Path -LiteralPath $pastaOD)) {
            $t = Get-TamanhoPasta -Caminho $pastaOD -TimeoutSeg 90
            Write-Log ('Pasta do OneDrive: {0} · {1}' -f $pastaOD, (Format-Bytes $t)) 'DADO'
            if ($t -gt 20GB) {
                Add-Achado 'ALERTA' ('OneDrive sincronizando {0} no disco local' -f (Format-Bytes $t)) 'Ligue "Arquivos sob demanda" (clique no icone da nuvem > engrenagem > Configuracoes) para liberar espaco sem perder acesso.' 'Nuvem' 'Alto'
            }
        }
        if ($mem -gt 500MB) {
            Add-Achado 'ALERTA' ('OneDrive consumindo {0}' -f (Format-Bytes $mem)) 'Use o botao "Encerrar OneDrive agora" quando precisar de desempenho. Ele volta ao ser aberto de novo.' 'Nuvem' 'Medio'
        }
    } else {
        Write-Log 'OneDrive nao esta em execucao.' 'DADO'
    }

    $tm = @(Get-Process -Name 'Teams', 'ms-teams', 'msteams' -ErrorAction SilentlyContinue)
    if ($tm.Count -gt 0) {
        $mem = ($tm | Measure-Object WorkingSet64 -Sum).Sum
        Write-Log ('Teams rodando · {0} em {1} processo(s)' -f (Format-Bytes $mem), $tm.Count) 'DADO'
        if ($mem -gt 1GB) {
            Add-Achado 'ALERTA' ('Teams consumindo {0}' -f (Format-Bytes $mem)) 'Use "Encerrar Teams agora" durante o trabalho pesado nas planilhas. Ele volta ao ser aberto de novo.' 'Nuvem' 'Alto'
        }
    } else {
        Write-Log 'Teams nao esta em execucao.' 'DADO'
    }

    $ant = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\previous'
    if (Test-Path -LiteralPath $ant) {
        $t = Get-TamanhoPasta -Caminho $ant -TimeoutSeg 45
        if ($t -gt 200MB) {
            Add-Achado 'ALERTA' ('Versoes antigas do Teams ocupando {0}' -f (Format-Bytes $t)) 'O Modulo 3 remove com o item "Versoes antigas do Teams".' 'Nuvem' 'Medio'
        }
    }
}


# =====================================================================
# 7G. GRUPO D - ACOES QUE MUDAM A VELOCIDADE
# =====================================================================

# D16 ----------------------------------------------------------------

function Disable-ItemInicializacao {
    param([string]$Nome, [ValidateSet('Run','StartupFolder')][string]$Tipo = 'Run')
    $chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\' + $Tipo
    try {
        if (-not (Test-Path $chave)) { New-Item -Path $chave -Force | Out-Null }
        $desligado = [byte[]](3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        New-ItemProperty -Path $chave -Name $Nome -Value $desligado -PropertyType Binary -Force | Out-Null
        return $true
    } catch { return $false }
}

function Enable-ItemInicializacao {
    param([string]$Nome, [ValidateSet('Run','StartupFolder')][string]$Tipo = 'Run')
    $chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\' + $Tipo
    try {
        if (-not (Test-Path $chave)) { return $false }
        $ligado = [byte[]](2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        New-ItemProperty -Path $chave -Name $Nome -Value $ligado -PropertyType Binary -Force | Out-Null
        return $true
    } catch { return $false }
}





# Desfazer -----------------------------------------------------------
function Restore-Otimizacoes {
    Write-Titulo 'Desfazer otimizacoes'
    $e = Read-Estado
    if ($e.Count -eq 0) { Write-Log 'Nenhuma otimizacao registrada para desfazer.' 'OK'; return }

    $r = Show-Escolha -Titulo 'Desfazer' `
        -Mensagem "Isso devolve os efeitos visuais, os aplicativos em segundo plano e a inicializacao automatica ao estado anterior.`r`n`r`nContinuar?" `
        -Opcoes @('Desfazer tudo', 'Cancelar')
    if ($r -ne 0) { return }

    foreach ($grupo in @('efeitos_visuais', 'segundo_plano', 'storage_sense', 'barra_tarefas', 'onedrive_icone')) {
        if (-not $e.ContainsKey($grupo)) { continue }
        foreach ($a in $e[$grupo]) {
            try {
                if ($null -eq $a.Valor) {
                    Remove-ItemProperty -Path $a.Chave -Name $a.Nome -ErrorAction SilentlyContinue
                } else {
                    New-ItemProperty -Path $a.Chave -Name $a.Nome -Value $a.Valor -PropertyType $a.Tipo -Force | Out-Null
                }
                Write-Log ('Restaurado: {0}' -f $a.Nome) 'DADO'
            } catch { }
        }
    }

    if ($e.ContainsKey('inicializacao_desativada')) {
        foreach ($i in $e['inicializacao_desativada']) {
            if (Enable-ItemInicializacao -Nome $i.Nome -Tipo $i.Tipo) { Write-Log ('Reativado na inicializacao: {0}' -f $i.Nome) 'DADO' }
        }
    }
    foreach ($app in @('OneDrive', 'Teams')) {
        $k = 'inicializacao_' + $app
        if (-not $e.ContainsKey($k)) { continue }
        foreach ($nome in @('OneDrive', 'com.squirrel.Teams.Teams', 'Teams', 'TeamsMachineInstaller')) {
            [void](Enable-ItemInicializacao -Nome $nome -Tipo 'Run')
        }
        foreach ($lnk in @('OneDrive.lnk', 'Microsoft Teams.lnk', 'Teams.lnk')) {
            [void](Enable-ItemInicializacao -Nome $lnk -Tipo 'StartupFolder')
        }
        try {
            $kt = 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\MSTeams_8wekyb3d8bbwe\TeamsTfwStartupTask'
            if (Test-Path $kt) { Set-ItemProperty -Path $kt -Name 'State' -Value 2 -ErrorAction SilentlyContinue }
        } catch { }
        Write-Log ('Inicializacao do {0} restaurada.' -f $app) 'DADO'
    }

    if ($e.ContainsKey('zonas_citrix')) {
        foreach ($k in $e['zonas_citrix']) {
            try { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue; Write-Log ('Zona removida: {0}' -f $k) 'DADO' } catch { }
        }
    }

    Remove-Item -LiteralPath (Get-ArquivoEstado) -Force -ErrorAction SilentlyContinue
    Write-Log 'Tudo restaurado. Faca logoff para aplicar por completo.' 'OK'
}


function Set-AjusteRegistro {
    param([string]$Grupo, [array]$Ajustes)
    $backup = @()
    $bloqueados = 0
    foreach ($a in $Ajustes) {
        try {
            $atual = Get-ValorReg $a.Chave $a.Nome
            if (-not (Test-Path $a.Chave)) { New-Item -Path $a.Chave -Force -ErrorAction Stop | Out-Null }
            New-ItemProperty -Path $a.Chave -Name $a.Nome -Value $a.Valor -PropertyType $a.Tipo -Force -ErrorAction Stop | Out-Null
            $backup += @{ Chave = $a.Chave; Nome = $a.Nome; Valor = $atual; Tipo = $a.Tipo }
        } catch {
            $bloqueados++
            if ("$($_.Exception.Message)" -match 'autoriz|denied|negad') {
                Write-Log ('{0}: bloqueado por politica do TI (o Windows nao deixa o usuario mudar).' -f $a.Nome) 'ALERTA'
            } else {
                Write-Log ('{0}: nao foi possivel ajustar.' -f $a.Nome) 'ALERTA'
            }
        }
    }
    if ($bloqueados -gt 0) {
        Write-Log ('{0} de {1} ajustes deste grupo estao travados por politica. Os demais foram aplicados.' -f $bloqueados, $Ajustes.Count) 'DADO'
    }
    $e = Read-Estado
    $anterior = @()
    if ($e.ContainsKey($Grupo)) { $anterior = @($e[$Grupo]) }
    Save-Estado $Grupo ($anterior + $backup)
}

function Set-AjusteEfeitos {
    Set-AjusteRegistro -Grupo 'efeitos_visuais' -Ajustes @(
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects'; Nome = 'VisualFXSetting';    Valor = 2;   Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize';     Nome = 'EnableTransparency'; Valor = 0;   Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Control Panel\Desktop\WindowMetrics';                              Nome = 'MinAnimate';         Valor = '0'; Tipo = 'String' }
        @{ Chave = 'HKCU:\Control Panel\Desktop';                                            Nome = 'DragFullWindows';    Valor = '0'; Tipo = 'String' }
        @{ Chave = 'HKCU:\Control Panel\Desktop';                                            Nome = 'MenuShowDelay';      Valor = '0'; Tipo = 'String' }
    )
}

function Set-AjusteSegundoPlano {
    Set-AjusteRegistro -Grupo 'segundo_plano' -Ajustes @(
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications'; Nome = 'GlobalUserDisabled';       Valor = 1; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search';                       Nome = 'BackgroundAppGlobalToggle'; Valor = 0; Tipo = 'DWord' }
    )
}

function Set-AjusteSugestoes {
    Set-AjusteRegistro -Grupo 'segundo_plano' -Ajustes @(
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Nome = 'SilentInstalledAppsEnabled';    Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Nome = 'SystemPaneSuggestionsEnabled';  Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Nome = 'PreInstalledAppsEnabled';       Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Nome = 'OemPreInstalledAppsEnabled';    Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Nome = 'SubscribedContent-338389Enabled'; Valor = 0; Tipo = 'DWord' }
    )
}

function Update-Explorer {
    if ($script:ExplorerAtualizado) { return }
    $script:ExplorerAtualizado = $true
    try {
        Stop-Process -Name explorer -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) { Start-Process explorer.exe }
        Write-Log 'Explorer reiniciado: a barra de tarefas ja aparece limpa.' 'OK'
    } catch {
        Write-Log 'Nao foi possivel reiniciar o Explorer (bloqueio do antivirus).' 'DADO'
        Write-Log 'As mudancas da barra de tarefas aparecem no proximo logon.' 'DADO'
    }
}

function Set-AjusteBarraTarefas {
    Set-AjusteRegistro -Grupo 'barra_tarefas' -Ajustes @(
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search';                    Nome = 'SearchboxTaskbarMode';      Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced';         Nome = 'ShowTaskViewButton';        Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced';         Nome = 'TaskbarDa';                 Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced';         Nome = 'TaskbarMn';                 Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds';                     Nome = 'ShellFeedsTaskbarViewMode'; Valor = 2; Tipo = 'DWord' }
    )
    Update-Explorer
}

function Set-AjusteIconeOneDrive {
    Set-AjusteRegistro -Grupo 'onedrive_icone' -Ajustes @(
        @{ Chave = 'HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}';            Nome = 'System.IsPinnedToNameSpaceTree'; Valor = 0; Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'; Nome = 'System.IsPinnedToNameSpaceTree'; Valor = 0; Tipo = 'DWord' }
    )
    Update-Explorer
}

function Set-AjusteStorageSense {
    Set-AjusteRegistro -Grupo 'storage_sense' -Ajustes @(
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy'; Nome = '01';  Valor = 1;  Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy'; Nome = '04';  Valor = 1;  Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy'; Nome = '08';  Valor = 1;  Tipo = 'DWord' }
        @{ Chave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy'; Nome = '256'; Valor = 30; Tipo = 'DWord' }
    )
}

# =====================================================================
# 7H. MODULO 5 - ARQUIVOS E PASTAS GRANDES
# =====================================================================
function Get-VeredictoPasta {
    param([string]$Caminho)
    if (Test-DentroDaPasta -Caminho $Caminho -Pasta $script:PastaClinica)        { return @{ V = 'NAO APAGAR'; M = 'Pasta clinica. Nunca apague.' } }
    if ($Caminho -match $script:RaizesClinicas) { return @{ V = 'NAO APAGAR'; M = 'Dado de sistema clinico.' } }
    if ($Caminho -match 'SQL Server|Tomcat') { return @{ V = 'NAO APAGAR'; M = 'Aplicativo clinico em uso.' } }
    if ($Caminho -match 'CentBrowser')      { return @{ V = 'NAO APAGAR'; M = 'CentBrowser e o navegador do prontuario da clinica.' } }
    if ($Caminho -match 'Digitalcore|Onis') { return @{ V = 'NAO APAGAR'; M = 'Dados do visualizador DICOM Onis.' } }
    if ($Caminho -match '\\Citrix$|\\Citrix\\')  { return @{ V = 'NAO APAGAR'; M = 'Cache do Citrix. Este kit nunca o toca.' } }
    if ($Caminho -match '\\Windows$|\\Windows\\|\\Program Files|ProgramData')    { return @{ V = 'NAO APAGAR'; M = 'Pasta do sistema.' } }
    if ($Caminho -match '\\Downloads$')                                          { return @{ V = 'AVALIAR';    M = 'Downloads. O Modulo 3 remove o que passou do periodo escolhido.' } }
    if ($Caminho -match 'Cache|\\Temp$|\\CrashDumps|INetCache|\\Packages$')      { return @{ V = 'AVALIAR';    M = 'Cache. O Modulo 3 ja limpa boa parte.' } }
    if ($Caminho -match '\\Desktop$|\\Documents$|\\Documentos$|\\Pictures$|\\Videos$') { return @{ V = 'AVALIAR'; M = 'Seus arquivos. Mova para a rede o que for do servico.' } }
    return @{ V = 'AVALIAR'; M = 'Confira o conteudo antes de mover ou apagar.' }
}

function Get-CandidatosPastas {
    $lista = @()
    # somente o disco do sistema: e o unico cujo espaco livre afeta o desempenho
    $raizes = @($env:USERPROFILE, $env:LOCALAPPDATA, $env:APPDATA, $script:PastaClinica, 'C:\')
    foreach ($r in $raizes) {
        if (-not (Test-Path -LiteralPath $r)) { continue }
        try {
            $lista += @(Get-ChildItem -LiteralPath $r -Directory -Force -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -notmatch '^\$|^System Volume Information$|^Recovery$|^Windows$|^Program Files|^ProgramData$|^Users$|^PerfLogs$|^Intel$|^Config\.Msi$|^AppData$' } |
                        Select-Object -ExpandProperty FullName)
        } catch { }
    }
    return @($lista | Where-Object { $_ -like 'C:\*' } | Select-Object -Unique)
}

function Invoke-ArquivosGrandes {
    Clear-SnapshotsOperacao
    Write-Titulo 'Modulo 4 - Arquivos e pastas grandes'
    Write-Log 'Somente leitura. O kit nao apaga nada aqui: voce decide item por item.' 'DADO'
    Write-Log 'So o disco C:. E o unico cujo espaco livre afeta o desempenho do Windows.' 'DADO'
    Write-Log 'Pastas do Windows e Program Files ficam de fora: nao ha o que fazer nelas sem administrador.' 'DADO'

    # a paginacao pode morar em outro disco - se aquele disco encher, ai sim pesa
    try {
        foreach ($pf in (Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue)) {
            $letra = ($pf.Name.Substring(0, 2)).ToUpper()
            if ($letra -ne 'C:') {
                $d = Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='" + $letra + "'") -ErrorAction SilentlyContinue
                if ($d -and $d.Size) {
                    $pct = [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
                    Write-Log ('Paginacao do Windows fica em {0} ({1}% livre). Enquanto esse disco tiver espaco, ele nao pesa no desempenho.' -f $letra, $pct) 'DADO'
                    if ($pct -lt 10) {
                        Add-Achado 'ALERTA' ('Disco {0} da paginacao com apenas {1}% livre' -f $letra, $pct) 'A paginacao mora nesse disco: se ele encher, a maquina inteira trava. Libere espaco nele tambem.' 'Disco' 'Alto'
                    }
                }
            }
        }
    } catch { }

    $script:GrandesItens = @()
$script:ModoPlano    = 'LIMPEZA'
$script:UltimoGanhoMemoria = 0
$script:SessaoEncerrados   = @()
$script:UltimaPrioridade   = $null
$script:ExplorerAtualizado = $false
$script:LogonSegundos      = 0
$script:StatusSessao       = $null
    $script:ListaGrandes.Items.Clear()

    # ---------------- 10 maiores arquivos ----------------
    Set-Status 'Procurando arquivos grandes...'
    $raizesArq = @($env:USERPROFILE, $script:PastaClinica, 'C:\Temp', 'C:\Users\Public') |
                 Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique
    $arqs = @()
    foreach ($r in $raizesArq) {
        if ($script:Cancelar) { break }
        Set-Status ('Arquivos grandes em ' + $r)
        $arqs += Get-MaioresArquivos -Raiz $r -Top 30 -TimeoutSeg 90 -MinimoMB $script:Lim.ArquivoGrandeMB
    }
    $arqs = @($arqs | Sort-Object Length -Descending | Select-Object -First 10)

    Write-Log '' 'DADO'
    Write-Log 'OS 10 MAIORES ARQUIVOS' 'TITULO'
    if ($arqs.Count -eq 0) {
        Write-Log ('Nenhum arquivo acima de {0} MB.' -f $script:Lim.ArquivoGrandeMB) 'DADO'
    }
    $n = 0
    foreach ($f in $arqs) {
        $n++
        $v = Get-VeredictoArquivo -Arquivo $f
        switch ($v.V) {
            'MANTER' { $ver = 'NAO APAGAR';      $nivel = 'CRITICO' }
            'SEGURO' { $ver = 'PODE APAGAR';     $nivel = 'OK' }
            default  { $ver = 'ATENCAO-AVALIAR'; $nivel = 'ALERTA' }
        }
        Write-Log ('{0,2}. [{1,-15}] {2,10}  {3}' -f $n, $ver, (Format-Bytes $f.Length), $f.FullName) $nivel
        Write-Log ('    ' + $v.M) 'DADO'
        $script:GrandesItens += [pscustomobject]@{ Tipo = 'Arquivo'; Caminho = $f.FullName; Bytes = $f.Length; Veredicto = $ver; Motivo = $v.M }
        [void]$script:ListaGrandes.Items.Add(('[{0,-15}] {1,10}  {2}' -f $ver, (Format-Bytes $f.Length), $f.FullName))
        Add-ItemInv 'Arquivo grande' $f.Name (Format-Bytes $f.Length) ($ver + ' - ' + $f.FullName)
    }

    # ---------------- 10 maiores pastas ----------------
    Write-Log '' 'DADO'
    Write-Log 'OS 10 MAIORES DIRETORIOS' 'TITULO'
    $cands = Get-CandidatosPastas
    $medidas = @()
    $i = 0
    foreach ($c in $cands) {
        if ($script:Cancelar) { break }
        $i++
        Set-Status ('Medindo pasta {0}/{1}: {2}' -f $i, $cands.Count, (Split-Path $c -Leaf))
        $b = Get-TamanhoPasta -Caminho $c -TimeoutSeg 45
        if ($b -gt 0) { $medidas += [pscustomobject]@{ Caminho = $c; Bytes = $b } }
    }
    $medidas = @($medidas | Sort-Object Bytes -Descending | Select-Object -First 10)

    $n = 0
    foreach ($m in $medidas) {
        $n++
        $v = Get-VeredictoPasta -Caminho $m.Caminho
        $ver = $(if ($v.V -eq 'NAO APAGAR') { 'NAO APAGAR' } else { 'ATENCAO-AVALIAR' })
        $nivel = $(if ($ver -eq 'NAO APAGAR') { 'CRITICO' } else { 'ALERTA' })
        Write-Log ('{0,2}. [{1,-15}] {2,10}  {3}' -f $n, $ver, (Format-Bytes $m.Bytes), $m.Caminho) $nivel
        Write-Log ('    ' + $v.M) 'DADO'
        $script:GrandesItens += [pscustomobject]@{ Tipo = 'Pasta'; Caminho = $m.Caminho; Bytes = $m.Bytes; Veredicto = $ver; Motivo = $v.M }
        [void]$script:ListaGrandes.Items.Add(('[{0,-15}] {1,10}  {2}' -f $ver, (Format-Bytes $m.Bytes), $m.Caminho))
        Add-ItemInv 'Pasta grande' (Split-Path $m.Caminho -Leaf) (Format-Bytes $m.Bytes) ($v.V + ' - ' + $m.Caminho)
    }

    $pode  = @($script:GrandesItens | Where-Object { $_.Veredicto -eq 'PODE APAGAR' })
    $aten  = @($script:GrandesItens | Where-Object { $_.Veredicto -eq 'ATENCAO-AVALIAR' })
    Write-Log '' 'DADO'
    Write-Log ('PODE APAGAR ..... {0} em {1} item(ns) - descartavel, o Windows recria ou nao usa mais' -f (Format-Bytes (($pode | Measure-Object Bytes -Sum).Sum)), $pode.Count) 'DADO'
    Write-Log ('ATENCAO-AVALIAR . {0} em {1} item(ns) - confira o conteudo antes de mover ou apagar' -f (Format-Bytes (($aten | Measure-Object Bytes -Sum).Sum)), $aten.Count) 'DADO'
    Write-Log 'NAO APAGAR ...... dado clinico, banco, navegador do prontuario ou arquivo de sistema' 'DADO'
    Write-Log 'Nenhum diretorio recebe "PODE APAGAR": pasta inteira sempre exige conferencia.' 'DADO'
    Write-Log 'Selecione qualquer linha no painel a direita e clique em "Abrir no Explorer".' 'ACAO'
    $script:PainelLimpeza.Visible = $false
    $script:PainelSessao.Visible = $false
    $script:PainelGrandes.Visible = $true
}

function Open-ItemGrande {
    $idx = $script:ListaGrandes.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:GrandesItens.Count) {
        [System.Windows.Forms.MessageBox]::Show('Selecione uma linha da lista.', 'Abrir no Explorer', 'OK', 'Information') | Out-Null
        return
    }
    $item = $script:GrandesItens[$idx]
    if (-not (Test-Path -LiteralPath $item.Caminho)) {
        Write-Log ('Nao existe mais: {0}' -f $item.Caminho) 'ALERTA'
        return
    }
    try {
        if ($item.Tipo -eq 'Pasta') {
            Start-Process explorer.exe -ArgumentList ('"{0}"' -f $item.Caminho)
        } else {
            Start-Process explorer.exe -ArgumentList ('/select,"{0}"' -f $item.Caminho)
        }
        Write-Log ('Aberto no Explorer: {0}' -f $item.Caminho) 'DADO'
    } catch {
        Write-Log ('Nao foi possivel abrir: {0}' -f $_.Exception.Message) 'ALERTA'
    }
}



# =====================================================================
# 7J. MODULO 7 - OTIMIZAR A SESSAO ATUAL
# ---------------------------------------------------------------------
# Tudo aqui vale so para a sessao de agora. O proximo logon devolve o
# estado normal da maquina: nada e desinstalado nem desativado.
# Navegador do prontuario, navegador web e Citrix sao sempre preservados.
# =====================================================================

# Processos que sustentam a area clinica: nunca sao encerrados nem rebaixados.
# Audio, microfone e utilitarios de periferico: encerrados na sessao mesmo estando
# na lista geral de protegidos. O som continua funcionando - o driver e o servico de
# audio do Windows rodam como SISTEMA. Sai apenas a camada de realce e os icones.
$script:SessaoAudio = 'RAVBg|RAVCpl|RtkNGUI|RtkAudUService|RtHDVBg|RtHDVCpl|RtlUpd|RealtekAudio' +
    '|WavesSvc|WavesSysSvc|WavesAudio|MaxxAudio' +
    '|Nahimic|A-Volute|AudioCenter' +
    '|DolbyDAX|DAX3API|DolbyAudio|DTSAPO|DTSAudio|SonicStudio|SonicSuite|SmartAudio|CxAudMsg|CxUIU' +
    '|LogiOptions|LogiLDA|Logitech|iCUE|SteelSeries|CorsairService' +
    '|WebcamService|CameraHelper|YourPhoneCamera' +
    '|fsquirt|BTTray|BTStackServer|BluetoothUserService|IntelBluetooth|btplayerctrl|BluetoothHeadset|BtwRSupportService'

# Enfeites do Windows: widget, barra de pesquisa, area de trabalho virtual,
# hub de comentarios e avisos de atualizacao. Todos voltam sozinhos ou no logon.
$script:SessaoEnfeites = 'Widgets|WidgetService|SearchHost|SearchApp|SearchUI' +
    '|StartMenuExperienceHost|ShellExperienceHost|ShellHost|TaskViewHost|MultitaskingViewHost|TextInputHost' +
    '|FeedbackHub|PilotshubApp|GetHelp|Windows\.Feedback' +
    '|jusched|jucheck|JavaUpdate|JavaCheck|OneDriveStandaloneUpdater|GoogleUpdate|MicrosoftEdgeUpdate|EdgeUpdate' +
    '|SCNotification|UserOOBEBroker|CompPkgSrv|WindowsInternal|LockApp|SystemSettingsBroker'

# Acesso remoto: e por aqui que o TI entra na maquina. Nunca encerrado.
$script:SessaoRemoto = 'TeamViewer|tv_w32|tv_x64|uvnc|winvnc|vncserver|tvnserver|ScreenConnect|AnyDesk|RustDesk' +
    '|CmRcService|CmRcViewer|RcAgent|DameWare|BeyondTrust|bomgar|LogMeIn|LMIGuardian|Splashtop|ZohoAssist' +
    '|quickassist|^msra$|^mstsc$|RdpClip|rdpinit'

$script:SessaoPreservar = 'wfica32|wfcrun32|CDViewer|SelfService|Receiver|concentr|CtxWebHelper|AuthManSvr|redirector|HdxRtcEngine|CtxCFRUI|Citrix' +
    '|CentBrowser|msedge|chrome|firefox|iexplore|opera|vivaldi|brave' +
    '|ssonsvr|Microsoft\.AAD|AADBroker|TokenBroker' +
    '|WindowsTerminal|OpenConsole|FortiTray|FortiClient|FortiSSLVPN|Forti' +
    '|EXCEL|WINWORD|POWERPNT|OUTLOOK|MSACCESS|^olk$|onenote|StickyNot' +
    '|javaw|^java$|jp2launcher|Wheb|tasy' +
    # Prontuario eletronico com presenca mundial. Tasy e o da Philips usado no
    # Brasil e estava sozinho aqui; num parque fora dele o prontuario tem outro
    # nome, e prontuario encerrado no meio de um atendimento e o mesmo estrago
    # com qualquer marca.
    #
    # 'Epic' era o caso grave: o token 'Epic' da lista de CONHECIDOS existe para
    # o lancador de jogo, e casava Epic Systems. Medido, com janela aberta e
    # 800 MB: vinha PRE-MARCADO, com a nota 'programa conhecido e nao guarda
    # documento' - falsa nas duas afirmacoes para um prontuario.
    '|^Epic$|EpicSystems|Hyperspace|PowerChart|Cerner|MEDITECH|Soarian|Sectra' +
    '|Vitrea|Vsp|^VI\.|Mirada|Medis|4DM|Corridor|NeuroQ|Olea|TomTec|ARIA|Eclipse|Varian|MOSAIQ|Monaco|MIM|RayStation|Velocity|Onis|Digitalcore|dicom|PACS' +
    '|sqlservr|Tomcat|catalina|w3wp|inetinfo|MSMQ|postgres|mysqld|oracle|firebird' +
    '|^claude$|^claude-code$' +
    '|' + $script:SessaoRemoto

# Programas que ficam abertos mas nao devem ter a memoria compactada:
# sessao publicada do Citrix, aplicativo clinico de imagem, banco local e acesso remoto.
$script:SessaoNaoCompactar = '^claude$|^claude-code$|^python$|^pythonw$|wfica32|wfcrun32|CDViewer|concentr|Vitrea|Vsp|^VI\.|Mirada|Medis|4DM|Corridor|NeuroQ|Olea|TomTec|ARIA|Eclipse|Varian|MOSAIQ|Monaco|MIM|RayStation|Onis|Digitalcore|sqlservr|Tomcat|w3wp|' + $script:SessaoRemoto

# ---------------------------------------------------------------------
# INSTANTANEO DE PROCESSOS  -  uma consulta por operacao, nao por uso
#
# Medido nesta maquina, 318 processos, PowerShell 5.1:
#     Get-CimInstance Win32_Process (todos os campos) ... 641 ms
#     Get-CimInstance Win32_Process (os 4 que uso) ..... 367 ms
#     Get-Process ....................................... 28 ms
#
# O Modulo 5 repetia essa consulta em sete lugares, e o executor a fazia UMA
# VEZ POR PID atraves de Get-AppEmUso.
#
# CRONOMETRADO de ponta a ponta, 60 PIDs, e o numero depende de o Local Suite
# estar aberto - porque Get-AppEmUso sai no Test-Path quando nao ha em_uso.json,
# e so paga o CIM quando ha:
#
#                                        ANTES      DEPOIS
#     com em_uso.json (Local Suite)      16,7 s      ~2 s     <- o caso que importa
#     sem em_uso.json                     1,6 s      ~1,3 s
#
#   Por PID: Get-AppEmUso completo sem instantaneo 278 ms; com o instantaneo
#   pronto 25 ms; no modo -Rapido 9 ms.
#
# A primeira versao deste comentario dizia "45,6 s -> 0,37 s, 124x". Era
# projecao a partir do custo unitario do CIM, nao medicao: supunha que todo
# Get-AppEmUso pagasse a consulta, o que nao acontece com o arquivo ausente.
# Fica registrado porque o erro e instrutivo - extrapolar custo unitario para
# custo total inventou uma ordem de grandeza.
#
# POR QUE O INSTANTANEO TEM PRAZO. Decidir encerrar processo numa estacao
# clinica com dado velho e inaceitavel: um PID reciclado passaria a ser visto
# com o caminho do ocupante ANTERIOR, e um processo clinico nascido depois do
# instantaneo perderia a protecao por caminho. Por isso:
#   - o instantaneo vale dentro de UMA operacao e e descartado no inicio da
#     seguinte, nunca reaproveitado entre cliques;
#   - o EXECUTOR nao confia nele para decidir: relê o caminho ao vivo do
#     processo que esta a ponto de encerrar (1,7 ms cada, 100 ms para 60);
#   - a sinalizacao em_uso.json e relida do disco antes de cada encerramento,
#     que custa 9 ms medidos, e so a caminhada de descendencia usa o
#     instantaneo.
# ---------------------------------------------------------------------
$script:SnapProc = $null

function Get-SnapshotSvc {
    # Medido nesta maquina: Win32_Service completo 1.465 ms, com 3 campos 938 ms,
    # Get-Service 166 ms. Para casar NOME e ESTADO - que e tudo o que o papel de
    # servidor e o Bluetooth precisam - Get-Service basta e e 9x mais rapido.
    # Projetado para o mesmo formato, com State em vez de Status, para que quem
    # consome nao precise saber de onde veio.
    #
    # O inventario do Modulo 2 segue usando Win32_Service: ele precisa de
    # StartMode e StartName, que Get-Service nao da no PowerShell 5.1.
    if ($script:SnapSvc -and $script:SnapSvc.Count -gt 0) { return $script:SnapSvc }

    # SilentlyContinue, nao Stop. Get-Service tropeca em servico individual que
    # o usuario nao pode consultar; com -ErrorAction Stop esse tropeco virava
    # terminante, o catch engolia e a lista saia VAZIA. E lista vazia aqui nao
    # parece erro: parece "esta maquina nao hospeda servico nenhum", e o aviso
    # CRITICO de papel de servidor desaparecia calado. Com SilentlyContinue vem
    # o que deu para ler, que e o comportamento certo para um diagnostico.
    $script:SnapSvc = @(Get-Service -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; DisplayName = $_.DisplayName; State = "$($_.Status)" }
    })

    if ($script:SnapSvc.Count -eq 0) {
        # segunda tentativa por outro caminho, porque zero servico nao existe
        try {
            $script:SnapSvc = @(Get-CimInstance Win32_Service -Property Name,DisplayName,State -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ Name = $_.Name; DisplayName = $_.DisplayName; State = "$($_.State)" }
            })
        } catch { }
    }
    if ($script:SnapSvc.Count -eq 0) {
        # Nao cacheia o fracasso e NAO o deixa passar por resposta. Uma maquina
        # sem nenhum servico nao existe; se a lista veio vazia, foi a leitura
        # que falhou, e quem le o log precisa saber que aquela verificacao nao
        # aconteceu - em vez de concluir que deu tudo certo.
        Write-Log 'Nao foi possivel ler a lista de servicos desta maquina: a verificacao de papel de servidor NAO foi feita.' 'ALERTA'
    }
    return $script:SnapSvc
}

function Clear-SnapshotsOperacao {
    # Chamado no inicio de cada modulo. O instantaneo vale para UMA operacao.
    $script:SnapProc         = $null
    $script:SnapSvc          = $null
    $script:EmUsoOperacao    = $null
    $script:SnapProcEm       = $null
    $script:EmUsoOperacaoEm  = $null
    $script:SnapSessao       = @{}
    $script:SnapGrupos       = $null
}

function Clear-SnapshotProc { $script:SnapProc = $null }

function Get-EmUsoOperacao {
    # O conjunto de PIDs que o Local Suite protege AGORA. Junta duas leituras que
    # sozinhas deixam buraco:
    #
    #   -Rapido  le so o arquivo, 9 ms. Traz PID que o Local Suite acabou de
    #            listar, e NAO traz descendencia.
    #   completo caminha a descendencia, 278 ms. Traz o worker que nasceu de um
    #            PID listado, e nao sabe do que entrou na lista depois.
    #
    # O codigo anterior escolhia um OU outro por "if Protegidos.Count -eq 0",
    # o que trocava o rapido pelo completo exatamente quando o arquivo estava
    # vazio - isto e, quando nao havia nada a proteger. No caso que importa, com
    # o arquivo cheio, a descendencia era simplesmente descartada.
    #
    # Agora e uniao, e a descendencia tem prazo em vez de ser calculada uma vez
    # no inicio: o TotalSegmentator cria worker o tempo todo, e um Aplicar leva
    # minutos. Renovar custa 278 ms a cada 15 s de lote.
    param([switch]$Renovar)
    $agora = Get-Date
    $venceu = $Renovar -or (-not $script:EmUsoOperacao) -or (-not $script:EmUsoOperacaoEm) -or
              (($agora - $script:EmUsoOperacaoEm).TotalSeconds -gt $script:SnapMaxSeg)
    if ($venceu) {
        Clear-SnapshotProc
        $script:EmUsoOperacao   = Get-AppEmUso
        $script:EmUsoOperacaoEm = $agora
    }
    $r = Get-AppEmUso -Rapido
    $set = @{}
    foreach ($id in $r.Protegidos) { $set[[int]$id] = $true }
    if ($script:EmUsoOperacao -and $script:EmUsoOperacao.Fresco) {
        foreach ($id in $script:EmUsoOperacao.Protegidos) { $set[[int]$id] = $true }
        # Fresco da uniao: se qualquer das duas leituras esta fresca, a protecao
        # vale. Test-PodeEncerrarSessao so olha Protegidos quando Fresco e certo,
        # entao perder este sinalizador anularia a uniao inteira em silencio.
        $r.Fresco = $true
    }
    $r.Protegidos = @($set.Keys)
    return $r
}

function Get-CaminhoAoVivo {
    # O caminho de UM processo, para quem vai decidir algo sobre ele.
    #
    # Ao vivo primeiro, instantaneo como reforco - nunca o contrario. Processo
    # nascido depois do retrato nao esta no mapa, e e exatamente o caso que
    # importa: a inferencia que o lote do Local Suite acabou de abrir. Buscando
    # so no mapa, ela vinha com caminho vazio, Test-CaminhoEcossistema
    # respondia 'nao sei' e a protecao por marca simplesmente nao acontecia.
    #
    # O mapa continua servindo para o contrario: processo cujo .Path nao abre,
    # onde o CIM deu o caminho e o .NET nao da.
    param($Processo, $Mapa)
    $cam = try { [string]$Processo.Path } catch { '' }
    if (-not $cam -and $Mapa) { $cam = [string]$Mapa[[int]$Processo.Id] }
    return [string]$cam
}

function Get-SnapshotProc {
    # Devolve @{ Caminhos = @{pid->caminho}; Pais = @{pid->pidPai}; Vivos = @{pid} }
    #
    # O cache TEM PRAZO. A guarda anterior era só "if ($script:SnapProc)", e um
    # pscustomobject vazio e verdadeiro em PowerShell: consulta que falhasse
    # ficava guardada para a operacao inteira, com Vivos e Caminhos vazios, o que
    # desligava em silencio a caminhada de descendencia. O cache existe para nao
    # repetir trabalho BOM, nao para congelar trabalho que nao aconteceu.
    # A guarda confere o CONTEUDO, nao so a existencia: instantaneo sem nenhum
    # processo vivo nao e resposta, e a leitura nao deve ter de confiar em que
    # todo escritor respeitou a regra. Os dois lados guardam a mesma invariante.
    if ($script:SnapProc -and $script:SnapProcEm -and $script:SnapProc.Vivos.Count -gt 0 -and
        ((Get-Date) - $script:SnapProcEm).TotalSeconds -le $script:SnapMaxSeg) {
        return $script:SnapProc
    }
    $snap = [pscustomobject]@{ Caminhos = @{}; Pais = @{}; Vivos = @{} }
    try {
        foreach ($proc in (Get-CimInstance Win32_Process -Property ProcessId,ParentProcessId,Name,ExecutablePath -ErrorAction Stop)) {
            $id = [int]$proc.ProcessId
            $snap.Vivos[$id] = $true
            $snap.Pais[$id]  = [int]$proc.ParentProcessId
            if ($proc.ExecutablePath) { $snap.Caminhos[$id] = "$($proc.ExecutablePath)" }
        }
    } catch { }
    # So guarda o que veio de fato. Vazio significa que a consulta nao respondeu,
    # e a proxima chamada tem de tentar de novo em vez de herdar o vazio.
    if ($snap.Vivos.Count -gt 0) {
        $script:SnapProc   = $snap
        $script:SnapProcEm = Get-Date
    }
    return $snap
}

function Get-CaminhosProcesso {
    # Mantida pelo nome, agora servida pelo instantaneo. Processo sem caminho e
    # nucleo do Windows, ja protegido por nome: ausencia nao abre brecha, e
    # Test-CaminhoEcossistema trata ausencia como 'nao sei'.
    return (Get-SnapshotProc).Caminhos
}

function Get-SufixosDns {
    # Os sufixos de DNS que valem NESTA maquina, descobertos em tempo de execucao,
    # mais o que estiver configurado. Nunca um dominio de terceiro chumbado: numa
    # rede que nao e aquela a consulta sempre falha, e o nome sai no ar.
    $l = @()
    try { $l += [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName } catch { }
    if ($env:USERDNSDOMAIN) { $l += $env:USERDNSDOMAIN }
    try {
        $r = Get-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Windows NT\DNSClient' -ErrorAction SilentlyContinue
        if ($r -and $r.SearchList) { $l += ("$($r.SearchList)" -split ',') }
    } catch { }
    if ($script:SufixosDnsExtra) { $l += $script:SufixosDnsExtra }
    return @($l | Where-Object { $_ } | ForEach-Object { "$_".Trim().TrimStart('.').ToLower() } |
             Where-Object { $_ } | Select-Object -Unique)
}

function Test-DentroDaPasta {
    # $Caminho esta DENTRO de $Pasta (ou e ela mesma)?
    #
    # Duas regras, e as duas ja custaram defeito neste projeto:
    #
    # 1. PASTA VAZIA RESPONDE FALSO. Nao ha pasta configurada, entao nao ha o que
    #    dizer. A versao anterior montava o curinga com a variavel dentro - com
    #    ela vazia, '-like (''*'' + '''' + ''*'')' vira '-like ''**''' e casa com
    #    TODO caminho. Medido: todo arquivo grande saia com veredicto MANTER e
    #    toda pasta com NAO APAGAR, com o motivo 'Esta na pasta clinica' - que e
    #    falso - e sem uma linha de erro.
    #
    # 2. COMPARA POR COMPONENTE, nao por prefixo de texto. 'C:\PASTA' como
    #    prefixo tambem casa 'C:\PASTA_ANTIGA'. Mesma familia do nome de regex
    #    sem ancora que casa demais, e mesma solucao de Test-CaminhoEcossistema.
    param([string]$Caminho, [string]$Pasta)
    if (-not "$Pasta".Trim() -or -not "$Caminho".Trim()) { return $false }
    $sep = [char]92
    $a = "$Caminho".TrimEnd($sep, [char]47)
    $b = "$Pasta".TrimEnd($sep, [char]47)
    if ($a.Length -lt $b.Length) { return $false }
    if (-not $a.StartsWith($b, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    if ($a.Length -eq $b.Length) { return $true }
    $prox = $a[$b.Length]
    return ($prox -eq $sep -or $prox -eq [char]47)
}

function Test-CaminhoEcossistema {
    # Carga do proprio ecossistema: inferencia do AUTO_CONTORNO, vigia, CADS.
    # Compara COMPONENTE DE PASTA, nao prefixo: acha C:\RADIOTERAPIA_AI\... e
    # D:\RADIOTERAPIA_AI\LOCAL_SUITE\... sem precisar saber a raiz de antemao.
    param([string]$Caminho)
    if (-not $Caminho) { return $false }   # sem caminho = nao sei, nao 'nao e nosso'
    # Split por CODIGO de caractere, nao por classe de regex. Escapar barra
    # invertida atravessa varias camadas de ferramenta e uma delas come o
    # escape: a primeira versao desta linha virou '[\/]' e passou a casar so
    # barra normal, entao nenhum caminho do Windows era reconhecido. O teste
    # pegou; sem ele teria virado protecao que nao protege.
    return ($Caminho.Split([char]92, [char]47) -contains $script:MarcaEcossistema)
}

# =====================================================================
# 4G. SINALIZACAO DE APP EM USO  -  acordo com o Radioterapia.AI Local Suite
#
# O Local Suite grava, enquanto o vigia ou um lote estiver vivo:
#     %LOCALAPPDATA%\RADIOTERAPIA_AI\em_uso.json
#     {"versao":1,"app":"...","atualizado":"<ISO-8601>","pids":[...],
#      "porta":8777,"segmentando":true}
# renovado a cada ~60 s e apagado na saida limpa.
#
# POR QUE ISSO EXISTE, e a marca de caminho nao resolve: o cmd.exe que abre o
# vigia mora em C:\Windows\System32 e o navegador tambem. Nenhuma marca de
# arvore instalada alcanca os dois, e fechar qualquer um deles derruba o app no
# meio de uma segmentacao.
#
# INVARIANTE: este arquivo so sabe ACRESCENTAR protecao. Nenhum caminho aqui
# devolve "pode encerrar". Arquivo malformado, velho ou adulterado deixa o kit
# conservador demais - nunca permissivo. Ele vive em %LOCALAPPDATA%, onde o
# usuario ja pode tudo, entao nao abre superficie nova.
#
# CARIMBO ILEGIVEL conta como FRESCO, de proposito. Ignorar arriscaria matar um
# lote; honrar arriscaria protecao eterna por um arquivo esquecido. A saida nao e
# escolher um dos dois: e VISIBILIDADE. O Modulo 5 sempre imprime quantos
# processos o arquivo protegeu e de quando ele e, entao sobreprotecao aparece na
# tela em vez de virar "o kit nao libera mais memoria e nao diz por que".
#
# PID RECICLADO e limitacao conhecida: em 5 minutos o Windows pode reusar um PID
# e eu protegeria um processo alheio. A consequencia e sobreprotecao, que e a
# direcao segura, e o aviso na tela a torna visivel.
# =====================================================================
$script:ArquivoEmUso = Join-Path $env:LOCALAPPDATA (Join-Path $script:MarcaEcossistema 'em_uso.json')
$script:EmUsoMaxMin  = 5

function Get-AppEmUso {
    # Devolve sempre um objeto; nunca lanca. Apps = vazio significa "sem
    # protecao extra", e e o estado normal de quem nao tem o Local Suite aberto.
    #
    # -Rapido le SO o arquivo, sem a caminhada de descendencia: 9 ms medidos,
    # contra 278 ms do modo completo sem instantaneo e 25 ms com ele. O
    # executor usa o modo rapido antes de cada encerramento, para que
    # um PID recem-listado pelo Local Suite seja respeitado mesmo no meio de um
    # Aplicar longo. A descendencia vem do conjunto completo, calculado uma vez.
    param([switch]$Rapido)
    $r = [pscustomobject]@{
        Existe = $false; Fresco = $false; Idade = $null; CarimboLegivel = $true
        App = ''; Porta = 0; Segmentando = $false; Pids = @(); Protegidos = @()
    }
    try {
        if (-not (Test-Path -LiteralPath $script:ArquivoEmUso)) { return $r }
        $r.Existe = $true
        $txt = [System.IO.File]::ReadAllText($script:ArquivoEmUso, [System.Text.Encoding]::UTF8)
        if (-not "$txt".Trim()) { $r.CarimboLegivel = $false; $r.Fresco = $true; return $r }
        $o = $null
        try { $o = $txt | ConvertFrom-Json -ErrorAction Stop } catch { }
        if (-not $o) { $r.CarimboLegivel = $false; $r.Fresco = $true; return $r }

        $r.App         = "$($o.app)"
        $r.Porta       = try { [int]$o.porta } catch { 0 }
        $r.Segmentando = ($o.segmentando -eq $true)
        $r.Pids        = @(@($o.pids) | ForEach-Object { try { [int]$_ } catch { } } | Where-Object { $_ -gt 0 })

        # carimbo: ilegivel conta como fresco, e fica registrado que era ilegivel
        $quando = $null
        try { $quando = [datetime]::Parse("$($o.atualizado)", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind) } catch { }
        if ($null -eq $quando) { try { $quando = [datetime]"$($o.atualizado)" } catch { } }
        if ($null -eq $quando) {
            $r.CarimboLegivel = $false
            $r.Fresco = $true
        } else {
            $r.Idade  = ((Get-Date) - $quando)
            $r.Fresco = ($r.Idade.TotalMinutes -le $script:EmUsoMaxMin -and $r.Idade.TotalMinutes -ge -$script:EmUsoMaxMin)
        }
        if (-not $r.Fresco) { return $r }
        if ($Rapido) { $r.Protegidos = @($r.Pids); return $r }

        # PIDs listados que estao vivos, mais toda a descendencia deles: o
        # TotalSegmentator cria worker o tempo todo, e worker nasce entre duas
        # renovacoes do arquivo.
        $snap = Get-SnapshotProc
        $pai = $snap.Pais
        $vivos = $snap.Vivos
        $set = @{}
        # Todo PID listado entra, esteja ou nao no instantaneo. O filtro por
        # vivos que havia aqui errava a direcao: proteger PID que ja morreu nao
        # custa nada - nao ha o que encerrar -, mas descartar PID vivo porque o
        # retrato envelheceu tira a protecao de carga clinica. Numa lista de
        # protecao, a duvida se resolve protegendo.
        foreach ($id in $r.Pids) { $set[$id] = $true }
        # sobe a cadeia de pais de cada processo vivo: se encontrar um PID
        # protegido, o descendente tambem entra. Corta em 40 niveis por seguranca.
        foreach ($id in $vivos.Keys) {
            if ($set.ContainsKey($id)) { continue }
            $atual = $id; $n = 0
            while ($n -lt 40) {
                if (-not $pai.ContainsKey($atual)) { break }
                $atual = $pai[$atual]
                if ($atual -le 0) { break }
                if ($set.ContainsKey($atual)) { $set[$id] = $true; break }
                $n++
            }
        }
        $r.Protegidos = @($set.Keys)
    } catch { }
    return $r
}

function Write-LogEmUso {
    # Chamado pelo Modulo 5 SEMPRE que ha protecao por arquivo, para que
    # sobreprotecao nunca seja silenciosa.
    param($EmUso)
    if (-not $EmUso -or -not $EmUso.Existe) { return }
    if (-not $EmUso.Fresco) {
        Write-Log ('Sinalizacao de app em uso encontrada mas VELHA ({0:N0} min). Ignorada.' -f $EmUso.Idade.TotalMinutes) 'DADO'
        return
    }
    $quando = if ($EmUso.CarimboLegivel) { ('atualizada ha {0:N0} min' -f $EmUso.Idade.TotalMinutes) } else { 'com data ILEGIVEL, tratada como recente' }
    $oque   = if ($EmUso.Segmentando) { 'lote em andamento' } else { 'aplicativo aberto' }
    Write-Log ('{0}: {1} protege {2} processo(s), {3}.' -f $EmUso.App, $oque, $EmUso.Protegidos.Count, $quando) 'ALERTA'
    Write-Log ('   Nenhum deles sera encerrado, rebaixado ou compactado. Porta {0}.' -f $EmUso.Porta) 'DADO'
    if (-not $EmUso.CarimboLegivel) {
        Write-Log '   A data do arquivo nao pode ser lida. Se o aplicativo nao estiver aberto, apague:' 'ACAO'
        Write-Log ('   ' + $script:ArquivoEmUso) 'DADO'
    }
}

function Test-PodeEncerrarSessao {
    param([string]$Nome, [string]$Caminho = '', [int]$ProcId = 0, $EmUso = $null)
    # Sinalizacao do Local Suite vem ANTES de tudo, e so sabe proteger.
    if ($ProcId -gt 0 -and $EmUso -and $EmUso.Fresco -and ($EmUso.Protegidos -contains $ProcId)) { return $false }
    # A carga do ecossistema vem ANTES de tudo: o processo e python.exe, nome que
    # nao distingue uma inferencia de 7 GB de um script descartavel. Quem
    # distingue e o caminho, e errar aqui interrompe um lote de contorno.
    if (Test-CaminhoEcossistema $Caminho) { return $false }
    if ($Nome -match $script:SessaoAudio) { return $true }          # audio e periferico: liberado
    if ($Nome -match $script:SessaoEnfeites) { return $true }       # enfeite do Windows: liberado
    if ($Nome -match $script:SessaoPreservar) { return $false }     # clinico, navegador, remoto
    if ($Nome -match $script:Protegidos -and $Nome -notmatch 'msedge|chrome|firefox|CentBrowser') { return $false }
    return $true
}


function Get-ProcessosSessao {
    # Por padrao, SO a sessao deste usuario - que e a unica onde o kit age.
    #
    # -TodasAsSessoes existe EXCLUSIVAMENTE PARA RELATORIO, e nenhum caminho de
    # acao chama com ele. Serve para um caso concreto: se o operador roda o kit
    # numa sessao diferente da que hospeda o trabalho, a classe Ecossistema sai
    # vazia e o relatorio afirma que nao ha nada em andamento. Correto para
    # aquela sessao, enganoso como retrato da maquina.
    #
    # Medido sem admin: WorkingSet64 e MainWindowHandle legiveis em 40 de 40
    # processos de outra sessao, zero falhas.
    #
    # O RESULTADO FICA GUARDADO POR OPERACAO, e isso e seguro por um motivo
    # especifico que nao se deve perder de vista: esta lista existe para MOSTRAR
    # e para montar o plano, nunca para decidir um encerramento. Quem decide e
    # Test-PodeEncerrarSessao, chamada pelo executor com o caminho lido ao vivo
    # e a sinalizacao relida do disco, PID por PID. Se um dia alguem usar esta
    # lista para decidir, o prazo dela passa a importar e este comentario esta
    # errado.
    #
    # A analise chamava esta funcao quatro vezes - duas por Get-GruposSessao, uma
    # pelo relatorio, uma com -TodasAsSessoes -, e cada uma refazia o percurso
    # inteiro de ~350 processos. Medido: 1,5 a 2,5 s por chamada.
    param([switch]$TodasAsSessoes)
    $chave = if ($TodasAsSessoes) { 'todas' } else { 'minha' }
    if ($script:SnapSessao -and $script:SnapSessao.ContainsKey($chave)) { return $script:SnapSessao[$chave] }

    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $caminhos = Get-CaminhosProcesso
    $emUso = Get-AppEmUso
    # -contains sobre array e busca linear, e aqui seria uma por processo. Com o
    # Local Suite segmentando, Protegidos tem a descendencia inteira.
    $protSet = @{}
    foreach ($id in $emUso.Protegidos) { $protSet[[int]$id] = $true }
    $saida = New-Object System.Collections.ArrayList
    try {
        foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
            if (-not $TodasAsSessoes -and $p.SessionId -ne $eu) { continue }   # servico ou outro usuario
            if ($p.Id -eq $PID) { continue }
            $nome = $p.ProcessName
            # Ao vivo, com o instantaneo de reforco: processo nascido depois do
            # retrato tem de aparecer como Ecossistema no relatorio tambem, senao
            # o kit mostra ao usuario uma classificacao que o executor nao segue.
            $cam  = Get-CaminhoAoVivo -Processo $p -Mapa $caminhos
            if ($emUso.Fresco -and $protSet.ContainsKey([int]$p.Id)) {
                $classe = 'Ecossistema'
            } elseif (Test-CaminhoEcossistema $cam) {
                $classe = 'Ecossistema'
            } elseif ($nome -match $script:SessaoAudio -or $nome -match $script:SessaoEnfeites) {
                $classe = 'Dispensavel'
            } elseif ($nome -match $script:SessaoPreservar) {
                $classe = 'Preservar'
            } elseif ($nome -match $script:Protegidos -and $nome -notmatch 'msedge|chrome|firefox|CentBrowser') {
                $classe = 'Essencial'
            } else {
                $classe = 'Dispensavel'
            }
            [void]$saida.Add([pscustomobject]@{
                Nome = $nome; Id = $p.Id; Memoria = $p.WorkingSet64
                TemJanela = ($p.MainWindowHandle -ne 0); Classe = $classe; Caminho = $cam
                Sessao = $p.SessionId; Minha = ($p.SessionId -eq $eu)
            })
        }
    } catch { }
    $r = @($saida)
    if ($r.Count -gt 0) { $script:SnapSessao[$chave] = $r }
    return $r
}

function Get-GruposSessao {
    # Mesmo acordo de Get-ProcessosSessao: agrupamento para MOSTRAR e para montar
    # o plano, nunca para decidir. Guardado por operacao porque a analise chama
    # duas vezes - o relatorio e Get-PlanoSessao.
    if ($script:SnapGrupos) { return $script:SnapGrupos }
    # Agrupa os dispensaveis por nome, separando quem tem janela aberta.
    $procs = @(Get-ProcessosSessao | Where-Object { $_.Classe -eq 'Dispensavel' })
    $conhecidos = 'OneDrive|OneDriveStandaloneUpdater|Teams|ms-teams|msteams|Update|Updater|GoogleUpdate|EdgeUpdate|MicrosoftEdgeUpdate|AdobeARM|AdobeGCClient|armsvc|jusched|jucheck|jp2launcher|Squirrel|SCNotification|Notification|Toast|UserOOBEBroker|CompPkgSrv|WindowsInternal|CCXProcess|Creative|CCLibrary|AdobeIPC|AdobeNotification|GameBar|Xbox|Gaming|NVIDIA Share|NVIDIA Web|YourPhone|PhoneExperience|Widget|SearchApp|Cortana|Copilot|SupportAssist|DellSupport|HPSupport|HPPrint|Vantage|ImController|McUICnt|Spotify|Dropbox|Box|GoogleDriveFS|Zoom|Slack|Discord|Skype|Steam|EpicGamesLauncher|EpicWebHelper|iTunesHelper|AppleMobileDevice|iPodService|StartMenuExperienceHost|ShellExperienceHost|TextInputHost|WindowsInternal'
    $conhecidos = $conhecidos + '|' + $script:SessaoAudio + '|' + $script:SessaoEnfeites
    $grupos = $procs | Group-Object Nome | ForEach-Object {
        $g = $_.Group
        $janela = @($g | Where-Object { $_.TemJanela }).Count -gt 0
        [pscustomobject]@{
            Nome      = $_.Name
            Qtd       = $_.Count
            Memoria   = ($g | Measure-Object Memoria -Sum).Sum
            Ids       = @($g | Select-Object -ExpandProperty Id)
            TemJanela = $janela
            Conhecido = ($_.Name -match $conhecidos)
        }
    }
    $r = @($grupos | Sort-Object Memoria -Descending)
    if ($r.Count -gt 0) { $script:SnapGrupos = $r }
    return $r
}

function Get-TarefasEmExecucao {
    # Get-ScheduledTask custa de 2,5 a 4 s medidos, e e o passo mais lento da
    # analise do Modulo 5. Medi as alternativas e nenhuma serve: schtasks /query
    # custa 10,6 s, com /v custa 60 s, enumerar pasta por pasta custa 9,9 s, e
    # filtrar so a raiz economiza 0,8 s mas PERDE 14 tarefas - inclusive as de
    # subpasta como \McAfee\wps\. O custo e inerente; o que da para melhorar e
    # dizer ao usuario que o kit esta nisso, em vez de deixar a janela parada.
    Set-Status 'Lendo tarefas agendadas (alguns segundos)...'
    Pump
    $lista = @()
    try {
        foreach ($t in (Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.State -eq 'Running' })) {
            if ("$($t.TaskPath)" -like '\Microsoft\Windows\*') { continue }   # tarefa de sistema: so o TI para

            # Acao apontando para a arvore do ecossistema: e o lote de contorno
            # preparando material para o dia seguinte. Parar isso as 2h da manha
            # nao tem ninguem para notar, e de manha ha um caso a menos sem
            # explicacao. Nem entra na lista - nao e questao de vir desmarcado.
            $exes = @()
            try { $exes = @($t.Actions | ForEach-Object { "$($_.Execute)" } | Where-Object { $_ }) } catch { }
            $doEcossistema = $false
            foreach ($x in $exes) {
                # o Execute vem com aspas, e as vezes com %VARIAVEL% por expandir
                $limpo = [Environment]::ExpandEnvironmentVariables(($x -replace '"', ''))
                if (Test-CaminhoEcossistema $limpo) { $doEcossistema = $true; break }
            }
            if ($doEcossistema) { continue }

            $prox = $null
            try { $prox = (Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop).NextRunTime } catch { }
            $lista += [pscustomobject]@{ Nome = $t.TaskName; Caminho = $t.TaskPath; Proxima = $prox; Executavel = ($exes -join ' | ') }
        }
    } catch { }
    return $lista
}

function Invoke-DiagSessaoAtiva {
    Write-Titulo 'Sessao atual: o que esta ativo agora'

    $todos = Get-ProcessosSessao
    if ($todos.Count -eq 0) { Write-Log 'Nao foi possivel mapear os processos da sessao.' 'ALERTA'; return }

    # Censo antes dos numeros, para a propria linha carregar a ressalva em vez de
    # depender de quem le lembrar dela. As classes abaixo valem para ESTA sessao,
    # que e a unica em que o kit age.
    $euId = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $maquina = @(Get-ProcessosSessao -TodasAsSessoes)
    $qtdSessoes = @($maquina | Select-Object -ExpandProperty Sessao -Unique).Count
    Write-Log ('Medido na sessao {0}, de {1} sessao(oes) ativa(s) nesta maquina.' -f $euId, $qtdSessoes) 'DADO'
    Add-ItemInv 'Sessao' 'Sessao medida' ("$euId") ("de $qtdSessoes sessao(oes) ativa(s)")

    foreach ($c in @('Essencial', 'Preservar', 'Ecossistema', 'Dispensavel')) {
        $g = @($todos | Where-Object { $_.Classe -eq $c })
        $mem = ($g | Measure-Object Memoria -Sum).Sum
        $rot = switch ($c) {
            'Essencial'   { 'Sistema e seguranca (intocavel)' }
            'Preservar'   { 'Clinico: Citrix, prontuario, navegador e Office' }
            'Ecossistema' { 'Carga do radioterapia.ai (contorno, vigia)' }
            default       { 'Dispensavel nesta sessao' }
        }
        Write-Log ('{0,-46} {1,4} processo(s)  {2,10}' -f $rot, $g.Count, (Format-Bytes $mem)) 'DADO'
        Add-ItemInv 'Sessao' $c ("$($g.Count) processos") ((Format-Bytes $mem) + " | sessao $euId")
    }

    # Carga do ecossistema FORA desta sessao. Reporto so isto das outras sessoes,
    # de proposito: classificar processo de sessao 0 como "dispensavel" daria um
    # numero que o kit nunca pode realizar - ali quem age e o TI, com elevacao.
    # Numero de maquina errada apresentado como numero da maquina certa e pior
    # que numero nenhum.
    $ecoFora = @($maquina | Where-Object { $_.Classe -eq 'Ecossistema' -and -not $_.Minha })
    if ($ecoFora.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ('ATENCAO: ha carga do radioterapia.ai em OUTRA sessao ({0} processo(s)):' -f $ecoFora.Count) 'ALERTA'
        foreach ($e in ($ecoFora | Sort-Object Memoria -Descending | Select-Object -First 6)) {
            Write-Log ('   sessao {0,-4} {1,-16} PID {2,-7} {3,10}' -f $e.Sessao, $e.Nome, $e.Id, (Format-Bytes $e.Memoria)) 'DADO'
        }
        Add-ItemInv 'Sessao' 'Carga do ecossistema fora desta sessao' ("$($ecoFora.Count) processos") (Format-Bytes (($ecoFora | Measure-Object Memoria -Sum).Sum))
    }

    # Carga do ecossistema merece linha propria: se ha inferencia em curso, o
    # usuario precisa saber ANTES de pedir otimizacao agressiva.
    # O achado olha a MAQUINA, nao a sessao: trabalho em andamento noutra sessao
    # continua sendo trabalho em andamento, e quem vai clicar em otimizar precisa
    # saber. So a lista detalhada fica na propria sessao.
    Write-LogEmUso (Get-AppEmUso)
    $eco = @($todos | Where-Object { $_.Classe -eq 'Ecossistema' })
    $ecoMaquina = @($maquina | Where-Object { $_.Classe -eq 'Ecossistema' })
    if ($eco.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log 'Carga do ecossistema em execucao (protegida dos Modulos 3 e 5):' 'DADO'
        foreach ($e in ($eco | Sort-Object Memoria -Descending | Select-Object -First 6)) {
            Write-Log ('   {0,-16} PID {1,-7} {2,10}   {3}' -f $e.Nome, $e.Id, (Format-Bytes $e.Memoria), $e.Caminho) 'DADO'
        }
    }
    $pesado = @($ecoMaquina | Where-Object { $_.Memoria -gt 1GB } | Sort-Object Memoria -Descending)
    if ($pesado.Count -gt 0) {
        $onde = if ($pesado[0].Minha) { 'nesta sessao' } else { ('na sessao ' + $pesado[0].Sessao) }
        Add-Achado 'ALERTA' ('Trabalho do radioterapia.ai em andamento {0}: {1} usando {2}' -f $onde, $pesado[0].Nome, (Format-Bytes $pesado[0].Memoria)) 'Nao encerre nem compacte agora. O kit nao toca nesses processos, mas fechar o programa por fora interrompe o lote - e num lote noturno nao ha ninguem para notar.' 'Sessao' 'Alto'
    }

    $grupos = Get-GruposSessao
    $memDisp = ($grupos | Measure-Object Memoria -Sum).Sum
    if ($grupos.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log 'Os maiores dispensaveis (j = com janela aberta):' 'DADO'
        foreach ($g in ($grupos | Select-Object -First 12)) {
            Write-Log ('[{0}] {1,-30} {2,10}  ({3} processo(s))' -f $(if ($g.TemJanela) { 'j' } else { ' ' }), $g.Nome, (Format-Bytes $g.Memoria), $g.Qtd) 'DADO'
        }
    }

    # estacao cabeada nao precisa de Bluetooth ligado
    try {
        $bt = @(Get-SnapshotSvc | Where-Object { $_.Name -eq 'bthserv' -and $_.State -eq 'Running' })
        $cabo = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -notmatch 'Wi-?Fi|Wireless|802\.11' })
        if ($bt.Count -gt 0 -and $cabo.Count -gt 0) {
            Write-Log '' 'DADO'
            Add-Achado 'ALERTA' 'Suporte a Bluetooth ligado numa estacao cabeada' 'O servico roda como SISTEMA e so o TI desliga. O Modulo 5 encerra os utilitarios de Bluetooth da sua sessao, mas o servico continua. Vale pedir a desativacao em massa nas estacoes de consultorio.' 'Sessao' 'Medio'
        }
    } catch { }

    $tarefas = Get-TarefasEmExecucao
    if ($tarefas.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ('Tarefas agendadas em execucao agora: {0}' -f $tarefas.Count) 'DADO'
        foreach ($t in $tarefas) { Write-Log ('- {0}{1}' -f $t.Caminho, $t.Nome) 'DADO' }
    }

    if ($memDisp -gt 400MB -or $grupos.Count -ge 6) {
        Add-Achado 'ALERTA' ('{0} em {1} programa(s) dispensavel(is) na sessao de agora' -f (Format-Bytes $memDisp), $grupos.Count) 'O Modulo 7 encerra tudo isso de uma vez, preserva Citrix e navegadores e devolve a memoria para o trabalho clinico. Volta ao normal no proximo logon.' 'Sessao' 'Alto'
    } else {
        Add-Achado 'OK' ('Sessao enxuta: {0} em dispensaveis' -f (Format-Bytes $memDisp)) '' 'Sessao'
    }
}

function Get-PlanoSessao {
    $plano = New-Object System.Collections.ArrayList

    # ---- 1. encerrar dispensaveis ----
    foreach ($g in (Get-GruposSessao)) {
        # Desconhecido COM janela ja vinha desmarcado - alguem esta olhando para
        # ele. Mas desconhecido SEM janela vinha MARCADO, e carga de fundo pesada
        # e exatamente isso: um lote de contorno rodando por um interpretador fora
        # da arvore instalada, um worker que o kit nao catalogou. A janela protege
        # quem esta na frente da tela; nao protege quem trabalha de madrugada.
        $pesadoDesconhecido = ((-not $g.Conhecido) -and $g.Memoria -gt ($script:Lim.SessaoDesconhecidoMB * 1MB))
        $marcar = ((-not $g.TemJanela) -or $g.Conhecido) -and (-not $pesadoDesconhecido)
        $nota = if ($pesadoDesconhecido) {
            ('Nao reconhecido e ocupando {0}. Esse tamanho em segundo plano parece trabalho em andamento, nao lixo: desmarcado por seguranca.' -f (Format-Bytes $g.Memoria))
        } elseif ($g.TemJanela -and -not $g.Conhecido) {
            'Tem janela aberta e o kit nao reconheceu: desmarcado por seguranca.'
        } elseif ($g.TemJanela) {
            'Tem janela aberta, mas e programa conhecido e nao guarda documento.'
        } else {
            'Roda em segundo plano. Volta sozinho quando for necessario.'
        }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'SESSAO'; Rotulo = ('Encerrar {0} ({1} proc.)' -f $g.Nome, $g.Qtd)
            Valor = (Format-Bytes $g.Memoria); Bytes = $g.Memoria
            Marcar = $marcar; Dados = $g; Nota = $nota
        })
    }

    # ---- 2. parar tarefas agendadas em execucao ----
    foreach ($t in (Get-TarefasEmExecucao)) {
        # "Volta no proximo agendamento" sem dizer quando e promessa vazia: para um
        # lote noturno o proximo agendamento e a noite seguinte. Diz a data, e se o
        # retorno esta longe vem DESMARCADO - quem marca e o usuario, mesma regra do
        # programa nao reconhecido com janela aberta.
        $marcarT = $true
        $notaT   = 'Volta a rodar no proximo agendamento.'
        if ($t.Proxima) {
            $horas = try { ([datetime]$t.Proxima - (Get-Date)).TotalHours } catch { 0 }
            if ($horas -gt 4) {
                $marcarT = $false
                $notaT = ('So volta a rodar em {0:dd/MM HH:mm}, daqui a {1:N0} h. Desmarcado por isso.' -f [datetime]$t.Proxima, $horas)
            } else {
                $notaT = ('Volta a rodar em {0:dd/MM HH:mm}.' -f [datetime]$t.Proxima)
            }
        } else {
            $marcarT = $false
            $notaT = 'Nao da para saber quando voltaria a rodar. Desmarcado por isso.'
        }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'TAREFA'; Rotulo = ('Parar tarefa {0}' -f $t.Nome); Valor = ''; Bytes = 0
            Marcar = $marcarT; Dados = $t; Nota = $notaT
        })
    }

    # ---- 3. prioridade de CPU para o que e clinico ----
    [void]$plano.Add([pscustomobject]@{
        Grupo = 'PRIORIDADE'; Rotulo = 'Dar prioridade de CPU ao Citrix, prontuario e Office'; Valor = ''; Bytes = 0
        Marcar = $true; Dados = 'prioridade'
        Nota = 'Sobe o clinico para acima do normal e rebaixa o resto. Zera no proximo logon.'
    })

    # ---- 4. devolver memoria dos programas que ficam ----
    [void]$plano.Add([pscustomobject]@{
        Grupo = 'MEMORIA'; Rotulo = 'Compactar a memoria dos programas que ficam abertos'; Valor = ''; Bytes = 0
        Marcar = $true; Dados = 'trim'
        Nota = 'Devolve ao Windows a memoria que o navegador e o Office reservaram sem usar. A sessao publicada do Citrix nao e mexida, para nao engasgar.'
    })

    return $plano
}

function Invoke-Modulo7Sessao {
    Clear-SnapshotsOperacao
    Write-Titulo 'Modulo 5 - Otimizar a sessao atual'
    Write-Log 'Nada e desinstalado nem desativado: o proximo logon devolve tudo ao normal.' 'DADO'
    Write-Log 'Citrix, navegador do prontuario, navegador web, Office e acesso remoto ficam sempre de fora.' 'DADO'
    Write-Log 'Audio e microfone sao encerrados: o som continua funcionando, sai so a camada de realce.' 'DADO'

    # em maquina que hospeda servico, otimizar a sessao pede combinado previo
    try {
        $procs = @(Get-Process -ErrorAction SilentlyContinue)
        $svcs = Get-SnapshotSvc
        $papeis = @()
        foreach ($c in (Get-CatalogoServidor)) {
            if (@($procs | Where-Object { $_.ProcessName -match $c.Padrao }).Count -gt 0 -or
                @($svcs | Where-Object { $_.Name -match $c.Padrao -or $_.DisplayName -match $c.Padrao }).Count -gt 0) { $papeis += $c.Chave }
        }
        if ($papeis.Count -gt 0) {
            Write-Log '' 'DADO'
            Write-Log ('ATENCAO: esta maquina hospeda servico ({0}).' -f ($papeis -join ', ')) 'CRITICO'
            Write-Log 'Os servicos rodam como SISTEMA e nao entram na lista, mas confirme com quem depende deles' 'ALERTA'
            Write-Log 'antes de otimizar uma estacao que serve outras pessoas.' 'ALERTA'
        }
    } catch { }

    Set-Status 'Lendo processos da sessao...'
    $plano = Get-PlanoSessao
    Set-Status 'Montando o plano...'
    $script:ModoPlano = 'SESSAO'
    $script:ListaLimpeza.Items.Clear()

    $ordemGrupo = @{ 'SESSAO' = 1; 'TAREFA' = 2; 'PRIORIDADE' = 3; 'MEMORIA' = 4 }
    $plano = @($plano | Sort-Object @{ Expression = { $ordemGrupo[$_.Grupo] } }, @{ Expression = { -$_.Bytes } })
    $script:PlanoAtual = $plano

    foreach ($p in $plano) {
        [void]$script:ListaLimpeza.Items.Add(('[{0,-10}] {1,-52} {2,10}' -f $p.Grupo, $p.Rotulo, $p.Valor))
        $script:ListaLimpeza.SetItemChecked($script:ListaLimpeza.Items.Count - 1, $p.Marcar)
    }

    $mem = ($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'SESSAO' } | Measure-Object Bytes -Sum).Sum
    $nProc = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'SESSAO' }).Count
    $nTar = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'TAREFA' }).Count
    $naoRec = @($plano | Where-Object { $_.Grupo -eq 'SESSAO' -and -not $_.Marcar })

    Write-Log '' 'DADO'
    Write-Log 'O que vai acontecer ao clicar em "Aplicar tudo":' 'TITULO'
    Write-Log ('Programas encerrados .......... {0}' -f $nProc) 'ACAO'
    Write-Log ('Memoria devolvida ............. {0}' -f (Format-Bytes $mem)) 'ACAO'
    Write-Log ('Tarefas agendadas paradas ..... {0}' -f $nTar) 'ACAO'
    Write-Log 'Prioridade de CPU ............. clinico acima do normal, resto abaixo' 'ACAO'
    Write-Log 'Memoria compactada ............ navegador e Office (Citrix intocado)' 'ACAO'
    if ($naoRec.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ('{0} programa(s) com janela aberta ficaram desmarcados por nao serem reconhecidos:' -f $naoRec.Count) 'DADO'
        foreach ($n in $naoRec) { Write-Log ('   ' + $n.Rotulo) 'DADO' }
        Write-Log 'Marque na lista se souber que pode fechar.' 'DADO'
    }
    Write-Log '' 'DADO'
    Write-Log 'Desmarque o que nao quiser e clique em "Aplicar tudo".' 'ACAO'
    $script:PainelGrandes.Visible = $false
    $script:PainelLimpeza.Visible = $true
}

# ---------------------------------------------------------------------
# Painel de status da sessao otimizada
# ---------------------------------------------------------------------
function Get-ResumoPreservados {
    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $procs = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $eu })
    $grupos = @(
        [pscustomobject]@{ Nome = 'Citrix';        Padrao = 'wfica32|wfcrun32|CDViewer|SelfService|Receiver|concentr|redirector|AuthManSvr|CtxWebHelper' }
        [pscustomobject]@{ Nome = 'Tasy';          Padrao = 'tasy|Wheb' }
        [pscustomobject]@{ Nome = 'Java (Tasy)';   Padrao = '^javaw?$|jp2launcher' }
        [pscustomobject]@{ Nome = 'Navegador';     Padrao = 'CentBrowser|msedge|chrome|firefox' }
        [pscustomobject]@{ Nome = 'Office';        Padrao = 'EXCEL|WINWORD|POWERPNT|OUTLOOK|^olk$' }
        [pscustomobject]@{ Nome = 'Clinico';       Padrao = 'Vitrea|Vsp|Mirada|Medis|4DM|Corridor|NeuroQ|Olea|TomTec|ARIA|Eclipse|Varian|MOSAIQ|Monaco|MIM|RayStation|Onis' }
        [pscustomobject]@{ Nome = 'Seguranca';     Padrao = 'ntrtscan|tmlisten|TMBM|PccNTMon|ShowMsg|DSASvc|CSFalcon|stAgent|Cortex|cyserver|CcmExec|QualysAgent|SecurityHealth' }
        [pscustomobject]@{ Nome = 'Acesso remoto'; Padrao = $script:SessaoRemoto }
    )
    $saida = @()
    foreach ($g in $grupos) {
        $p = @($procs | Where-Object { $_.ProcessName -match $g.Padrao })
        if ($p.Count -gt 0) {
            $saida += [pscustomobject]@{ Nome = $g.Nome; Qtd = $p.Count; Memoria = ($p | Measure-Object WorkingSet64 -Sum).Sum }
        }
    }
    return $saida
}

function Add-LinhaStatus {
    param([string]$Texto = '', $Cor = $null)
    if (-not $script:TxtSessao) { return }
    $script:TxtSessao.SelectionStart  = $script:TxtSessao.TextLength
    $script:TxtSessao.SelectionLength = 0
    $script:TxtSessao.SelectionColor  = $(if ($Cor) { $Cor } else { $script:Cor.Texto })
    $script:TxtSessao.AppendText($Texto + "`r`n")
}

function Show-StatusSessao {
    param([ValidateSet('OTIMIZADA','ORIGINAL')][string]$Estado = 'OTIMIZADA')
    if (-not $script:TxtSessao) { return }
    $script:TxtSessao.Clear()

    if ($Estado -eq 'OTIMIZADA') {
        Add-LinhaStatus 'SESSAO OTIMIZADA PARA RADIOTERAPIA' $script:Cor.Ok
        if ($script:StatusSessao) {
            Add-LinhaStatus ('Otimizada as {0:HH:mm} de {0:dd/MM/yyyy}' -f $script:StatusSessao.Hora) $script:Cor.Fraco
        }
    } else {
        Add-LinhaStatus 'SESSAO NO ESTADO ORIGINAL' $script:Cor.Alerta
        Add-LinhaStatus ('Revertida as {0:HH:mm}' -f (Get-Date)) $script:Cor.Fraco
    }
    Add-LinhaStatus ''

    # ---- memoria ----
    Add-LinhaStatus 'MEMORIA' $script:Cor.Titulo
    if ($Estado -eq 'OTIMIZADA' -and $script:StatusSessao) {
        $tot = $script:StatusSessao.MemDevolvida + $script:StatusSessao.MemCompactada
        Add-LinhaStatus ('  Devolvida ao encerrar ....... {0}' -f (Format-Bytes $script:StatusSessao.MemDevolvida))
        Add-LinhaStatus ('  Devolvida ao compactar ...... {0}' -f (Format-Bytes $script:StatusSessao.MemCompactada))
        Add-LinhaStatus ('  Total ....................... {0}' -f (Format-Bytes $tot)) $script:Cor.Ok
    }
    try {
        $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $usoPct = [Math]::Round(((($so.TotalVisibleMemorySize - $so.FreePhysicalMemory) / $so.TotalVisibleMemorySize) * 100), 1)
        $cor = $(if ($usoPct -ge 90) { $script:Cor.Critico } elseif ($usoPct -ge 80) { $script:Cor.Alerta } else { $script:Cor.Ok })
        Add-LinhaStatus ('  Em uso agora ................ {0}%  ({1} livres)' -f $usoPct, (Format-Bytes ($so.FreePhysicalMemory * 1KB))) $cor
    } catch { }
    Add-LinhaStatus ''

    # ---- encerrados ----
    if ($Estado -eq 'OTIMIZADA' -and $script:StatusSessao -and $script:StatusSessao.Encerrados.Count -gt 0) {
        Add-LinhaStatus ('ENCERRADOS NESTA SESSAO ({0})' -f $script:StatusSessao.Encerrados.Count) $script:Cor.Titulo
        foreach ($n in ($script:StatusSessao.Encerrados | Select-Object -Unique)) {
            Add-LinhaStatus ('  ' + $n) $script:Cor.Fraco
        }
        Add-LinhaStatus ''
    }

    # ---- prioridade ----
    if ($Estado -eq 'OTIMIZADA' -and $script:UltimaPrioridade) {
        Add-LinhaStatus 'PRIORIDADE DE CPU' $script:Cor.Titulo
        Add-LinhaStatus ('  {0} clinico(s) acima do normal, {1} abaixo' -f $script:UltimaPrioridade.Acima, $script:UltimaPrioridade.Abaixo)
        Add-LinhaStatus ''
    }

    # ---- preservado e rodando agora ----
    Add-LinhaStatus 'PRESERVADO E RODANDO AGORA' $script:Cor.Titulo
    $pres = Get-ResumoPreservados
    if ($pres.Count -eq 0) {
        Add-LinhaStatus '  nada detectado' $script:Cor.Fraco
    } else {
        foreach ($g in $pres) {
            Add-LinhaStatus ('  {0,-16} {1,3} proc.  {2,10}' -f $g.Nome, $g.Qtd, (Format-Bytes $g.Memoria))
        }
    }
    Add-LinhaStatus ''

    if ($Estado -eq 'OTIMIZADA') {
        Add-LinhaStatus 'Reverter devolve as prioridades e reabre o que foi fechado.' $script:Cor.Acao
        Add-LinhaStatus 'O proximo logon devolve tudo ao normal de qualquer forma.' $script:Cor.Fraco
    } else {
        Add-LinhaStatus 'Prioridades normalizadas e programas reabertos.' $script:Cor.Acao
        Add-LinhaStatus 'O que saiu da inicializacao volta em "Desfazer otimizacoes".' $script:Cor.Fraco
    }

    $script:TxtSessao.SelectionStart = 0
    $script:TxtSessao.ScrollToCaret()

    # a lista do que seria aplicado sai de cena
    $script:ListaLimpeza.Items.Clear()
    $script:PlanoAtual = @()
    $script:PainelLimpeza.Visible = $false
    $script:PainelGrandes.Visible = $false
    $script:PainelSessao.Visible = $true
    Pump
}

function Invoke-RestaurarSessao {
    Write-Titulo 'Restaurar a sessao'
    Write-Log 'Devolve as prioridades ao normal e reabre o que foi encerrado no Modulo 5.' 'DADO'

    # ---- 1. prioridades de volta ao normal ----
    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $n = 0
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.SessionId -ne $eu -or $p.Id -eq $PID) { continue }
        try {
            if ($p.PriorityClass -ne [System.Diagnostics.ProcessPriorityClass]::Normal) {
                $p.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal
                $n++
            }
        } catch { }
    }
    Write-Log ('Prioridade normalizada em {0} programa(s).' -f $n) 'OK'

    # ---- 2. reabrir o que se sabe reabrir ----
    $e = Read-Estado
    $fechados = @()
    if ($e.ContainsKey('sessao_encerrados')) { $fechados = @($e['sessao_encerrados']) }

    if ($fechados.Count -eq 0) {
        Write-Log 'Nenhum registro de encerramento nesta sessao. Nada a reabrir.' 'DADO'
    } else {
        Write-Log ('Registro do Modulo 5: {0} programa(s) encerrado(s).' -f $fechados.Count) 'DADO'
    }

    # O Google Drive instala versoes lado a lado, entao o caminho nao pode ser
    # fixo. Medido nesta maquina: 130.0.2.0 e 131.0.2.0 instaladas e a que rodava
    # era a 130 - ou seja, a mais nova instalada NAO e necessariamente a que esta
    # em execucao. A lista vai da mais nova para a mais antiga e o laco de
    # restauracao tenta uma por uma ate alguma abrir, entao errar a primeira
    # apenas cai para a seguinte.
    #
    # Sem isso o kit fechava o Drive e nao sabia reabrir: a sincronizacao de
    # C:\AI_PROJETOS e a entrega em C:\AI_DEPLOY ficavam paradas sem ninguem notar,
    # ate a publicacao seguinte falhar em silencio.
    $driveFs = @()
    try {
        foreach ($base in @((Join-Path $env:ProgramFiles 'Google\Drive File Stream'),
                            (Join-Path ${env:ProgramFiles(x86)} 'Google\Drive File Stream'))) {
            if (-not $base -or -not (Test-Path -LiteralPath $base)) { continue }
            $driveFs += @(Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
                Sort-Object { try { [version]$_.Name } catch { [version]'0.0' } } -Descending |
                ForEach-Object { Join-Path $_.FullName 'GoogleDriveFS.exe' })
        }
    } catch { }

    $mapa = @(
        [pscustomobject]@{ Padrao = '^GoogleDriveFS$'; Rotulo = 'Google Drive'; Caminhos = $driveFs; Args = '' }
        [pscustomobject]@{ Padrao = 'OneDrive'; Rotulo = 'OneDrive'; Caminhos = @(
            (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
            (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')); Args = '/background' }
        [pscustomobject]@{ Padrao = 'ms-teams|msteams|^Teams$'; Rotulo = 'Microsoft Teams'; Caminhos = @(
            (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\MSTeams_8wekyb3d8bbwe\ms-teams.exe')); Args = '' }
        [pscustomobject]@{ Padrao = 'soffice'; Rotulo = 'LibreOffice'; Caminhos = @(
            (Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.exe')); Args = '' }
    )

    $reabertos = 0
    foreach ($m in $mapa) {
        $foiFechado = @($fechados | Where-Object { "$_" -match $m.Padrao }).Count -gt 0
        if (-not $foiFechado) { continue }
        if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match $m.Padrao }).Count -gt 0) {
            Write-Log ('{0}: ja esta aberto.' -f $m.Rotulo) 'DADO'
            continue
        }
        $abriu = $false
        foreach ($c in $m.Caminhos) {
            if (-not $c -or -not (Test-Path -LiteralPath $c)) { continue }
            try {
                if ($m.Args) { Start-Process -FilePath $c -ArgumentList $m.Args | Out-Null }
                else { Start-Process -FilePath $c | Out-Null }
                Write-Log ('{0}: reaberto.' -f $m.Rotulo) 'OK'
                $abriu = $true; $reabertos++
                break
            } catch { }
        }
        if (-not $abriu) { Write-Log ('{0}: nao encontrei o executavel para reabrir. Abra pelo menu Iniciar.' -f $m.Rotulo) 'ALERTA' }
    }

    # ---- 3. o que volta sozinho ----
    $sozinhos = @($fechados | Where-Object { "$_" -match 'SearchApp|ShellExperience|StartMenu|TextInput|Notification|Update|jusched|jucheck|Broker|CompPkgSrv|Widget|Copilot|Audio|Waves|Nahimic|Realtek|RAVBg|Rtk|Dolby|DTS|Logi|iCUE' })
    if ($sozinhos.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ('{0} item(ns) voltam sozinhos quando o Windows precisar deles:' -f $sozinhos.Count) 'DADO'
        foreach ($x in ($sozinhos | Select-Object -First 12 -Unique)) { Write-Log ('   ' + $x) 'DADO' }
        Write-Log 'A camada de audio e os icones de bandeja voltam no proximo logon.' 'DADO'
    }

    # limpa o registro: a sessao foi restaurada
    try {
        $e2 = Read-Estado
        if ($e2.ContainsKey('sessao_encerrados')) {
            $e2.Remove('sessao_encerrados')
            Write-Estado $e2
        }
    } catch { }

    Write-Log '' 'DADO'
    Write-Log ('Sessao restaurada: {0} programa(s) reabertos, prioridades normalizadas.' -f $reabertos) 'OK'
    Write-Log 'O que foi tirado da inicializacao pelo Modulo 3 volta pelo botao "Desfazer otimizacoes".' 'DADO'
}

function Invoke-AcaoPrioridade {
    # PRIORIDADE e MEMORIA sao os dois ultimos grupos da ordem de execucao, logo
    # rodam na maior distancia possivel do instantaneo. Por isso renovam: o
    # retrato do inicio do lote nao descreve mais a maquina.
    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $emUso = Get-EmUsoOperacao -Renovar
    $caminhos = Get-CaminhosProcesso
    $subiu = 0
    $baixou = 0
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.SessionId -ne $eu -or $p.Id -eq $PID) { continue }
        try {
            if ($p.ProcessName -match $script:SessaoPreservar -and $p.ProcessName -notmatch $script:SessaoAudio -and $p.ProcessName -notmatch $script:SessaoEnfeites) {
                $p.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal
                $subiu++
            } elseif (Test-PodeEncerrarSessao -Nome $p.ProcessName -Caminho (Get-CaminhoAoVivo -Processo $p -Mapa $caminhos) -ProcId ([int]$p.Id) -EmUso $emUso) {
                $p.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::BelowNormal
                $baixou++
            }
        } catch { }
    }
    $script:UltimaPrioridade = [pscustomobject]@{ Acima = $subiu; Abaixo = $baixou }
    Write-Log ('Prioridade ajustada: {0} programa(s) clinico(s) acima do normal, {1} abaixo do normal.' -f $subiu, $baixou) 'OK'
    Write-Log 'O Windows zera as prioridades no proximo logon.' 'DADO'
}

function Invoke-AcaoMemoria {
    # Usa so as propriedades do .NET. A versao anterior compilava uma funcao do
    # Windows em tempo de execucao - desnecessario, e o metodo abaixo funciona igual.
    #
    # Renova o instantaneo pelo mesmo motivo de Invoke-AcaoPrioridade: e o ultimo
    # grupo do lote. Pagar a consulta duas vezes no fim de um Aplicar que ja
    # levou segundos e barato ao lado de compactar uma inferencia de 7 GB.
    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $emUso = Get-EmUsoOperacao -Renovar
    $caminhos = Get-CaminhosProcesso
    $antes = 0.0
    $depois = 0.0
    $n = 0
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.SessionId -ne $eu -or $p.Id -eq $PID) { continue }
        # Citrix, imagem clinica, banco local e acesso remoto ficam de fora:
        # compactar ali causa engasgo visivel na tela.
        if ($p.ProcessName -match $script:SessaoNaoCompactar) { continue }
        # Carga do ecossistema: compactar working set de uma inferencia de 7 GB no
        # meio do lote forca o Windows a paginar tudo de volta. Dano garantido.
        if (Test-CaminhoEcossistema (Get-CaminhoAoVivo -Processo $p -Mapa $caminhos)) { continue }
        if ($emUso.Fresco -and ($emUso.Protegidos -contains [int]$p.Id)) { continue }
        try {
            $ws = $p.WorkingSet64
            # baixar o teto de memoria forca o Windows a devolver as paginas;
            # em seguida o teto original volta, sem limitar o programa
            $minAnt = $p.MinWorkingSet
            $maxAnt = $p.MaxWorkingSet
            $p.MinWorkingSet = [IntPtr]204800
            $p.MaxWorkingSet = [IntPtr]1048576
            $p.MaxWorkingSet = $maxAnt
            $p.MinWorkingSet = $minAnt
            $antes += $ws
            $p.Refresh()
            $depois += $p.WorkingSet64
            $n++
        } catch { }
        if ($n % 25 -eq 0) { Pump }
    }
    $ganho = $antes - $depois
    if ($ganho -lt 0) { $ganho = 0 }
    $script:UltimoGanhoMemoria = $ganho
    if ($n -eq 0) { Write-Log 'Nenhum programa aceitou a compactacao nesta maquina.' 'ALERTA'; return }
    Write-Log ('Memoria compactada em {0} programa(s): {1} devolvidos ao Windows.' -f $n, (Format-Bytes $ganho)) 'OK'
    Write-Log 'A primeira acao dentro de cada programa pode demorar um instante: ele recarrega o que precisa.' 'DADO'
}

function New-Botao {
    param([string]$Texto, [int]$Y, [int]$Altura = 42, $Cor = $null, [switch]$Secundario)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Texto
    $b.SetBounds(14, $Y, 244, $Altura)
    $b.FlatStyle = 'Flat'
    $b.BackColor = if ($Secundario) { $script:Cor.Painel } else { $script:Cor.Botao }
    $b.ForeColor = if ($Cor) { $Cor } else { $script:Cor.Texto }
    $b.Font      = New-Object System.Drawing.Font('Segoe UI', $(if ($Secundario) { 8.5 } else { 10 }), $(if ($Secundario) { [System.Drawing.FontStyle]::Regular } else { [System.Drawing.FontStyle]::Bold }))
    $b.TextAlign = 'MiddleLeft'
    $b.Padding   = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
    $b.FlatAppearance.BorderColor = $script:Cor.Borda
    $b.FlatAppearance.BorderSize  = 1
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text          = 'Kit de Suporte - Estacoes de Radioterapia'
$script:Form.Size          = New-Object System.Drawing.Size(1240, 820)
$script:Form.MinimumSize   = New-Object System.Drawing.Size(980, 620)
$script:Form.StartPosition = 'CenterScreen'
$script:Form.BackColor     = $script:Cor.Fundo
$script:Form.ForeColor     = $script:Cor.Texto
$script:Form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

# --- cabecalho ---
$cab = New-Object System.Windows.Forms.Panel
$cab.Dock = 'Top'; $cab.Height = 58; $cab.BackColor = $script:Cor.Painel
$lblTit = New-Object System.Windows.Forms.Label
$lblTit.Text = 'Kit de Suporte  ·  Estacoes de Radioterapia'
$lblTit.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
$lblTit.ForeColor = $script:Cor.Texto
$lblTit.SetBounds(18, 9, 620, 24)
$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = ('Somente acoes sem administrador  ·  versao {0}  ·  {1}\{2}' -f $script:Versao, $env:COMPUTERNAME, $env:USERNAME)
$lblSub.ForeColor = $script:Cor.Fraco
$lblSub.SetBounds(20, 33, 700, 18)
$cab.Controls.AddRange(@($lblTit, $lblSub))

# --- barra inferior ---
$rodape = New-Object System.Windows.Forms.Panel
$rodape.Dock = 'Bottom'; $rodape.Height = 46; $rodape.BackColor = $script:Cor.Painel

$script:LblStatus = New-Object System.Windows.Forms.Label
$script:LblStatus.Text = 'Pronto.'
$script:LblStatus.ForeColor = $script:Cor.Fraco
$script:LblStatus.SetBounds(18, 15, 430, 18)

$script:Barra = New-Object System.Windows.Forms.ProgressBar
$script:Barra.SetBounds(452, 14, 150, 16)
$script:Barra.Style = 'Blocks'

$btnCopiar = New-Object System.Windows.Forms.Button
$btnCopiar.Text = 'Copiar log'; $btnCopiar.SetBounds(620, 9, 100, 28); $btnCopiar.FlatStyle = 'Flat'
$btnCopiar.BackColor = $script:Cor.Botao; $btnCopiar.ForeColor = $script:Cor.Texto
$btnCopiar.FlatAppearance.BorderColor = $script:Cor.Borda

$btnSalvar = New-Object System.Windows.Forms.Button
$btnSalvar.Text = 'Salvar log'; $btnSalvar.SetBounds(728, 9, 100, 28); $btnSalvar.FlatStyle = 'Flat'
$btnSalvar.BackColor = $script:Cor.Botao; $btnSalvar.ForeColor = $script:Cor.Texto
$btnSalvar.FlatAppearance.BorderColor = $script:Cor.Borda

$btnLimparLog = New-Object System.Windows.Forms.Button
$btnLimparLog.Text = 'Limpar tela'; $btnLimparLog.SetBounds(836, 9, 100, 28); $btnLimparLog.FlatStyle = 'Flat'
$btnLimparLog.BackColor = $script:Cor.Botao; $btnLimparLog.ForeColor = $script:Cor.Texto
$btnLimparLog.FlatAppearance.BorderColor = $script:Cor.Borda

$script:BtnCancelar = New-Object System.Windows.Forms.Button
$script:BtnCancelar.Text = 'Parar'; $script:BtnCancelar.SetBounds(944, 9, 90, 28); $script:BtnCancelar.FlatStyle = 'Flat'
$script:BtnCancelar.BackColor = $script:Cor.Botao; $script:BtnCancelar.ForeColor = $script:Cor.Alerta
$script:BtnCancelar.FlatAppearance.BorderColor = $script:Cor.Borda
$script:BtnCancelar.Enabled = $false

$rodape.Controls.AddRange(@($script:LblStatus, $script:Barra, $btnCopiar, $btnSalvar, $btnLimparLog, $script:BtnCancelar))

# --- menu lateral ---
function New-Secao {
    param([string]$Texto, [int]$Y)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Texto; $l.SetBounds(18, $Y, 244, 18); $l.ForeColor = $script:Cor.Fraco
    $l.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
    return $l
}

$lateral = New-Object System.Windows.Forms.Panel
$lateral.Dock = 'Left'; $lateral.Width = 292; $lateral.BackColor = $script:Cor.Painel
$lateral.Padding = New-Object System.Windows.Forms.Padding(0,10,0,0)
$lateral.AutoScroll = $true

$script:BtnMod1 = New-Botao -Texto '1 · Preparar ambiente' -Y 16
$script:BtnMod2 = New-Botao -Texto '2 · Inventario e diagnostico' -Y 66
$script:BtnMod3 = New-Botao -Texto '3 · Limpeza segura' -Y 116

$script:BtnGrandes = New-Botao -Texto '4 · Arquivos e pastas grandes' -Y 166
$script:BtnSessao  = New-Botao -Texto '5 · Otimizar sessao atual'     -Y 216
$script:BtnPersist = New-Botao -Texto '6 · O que sobrevive ao logoff' -Y 266

$sec3 = New-Secao -Texto 'REVERTER E FECHAR O CICLO' -Y 330
$script:BtnRestaurar = New-Botao -Texto 'Restaurar sessao (reabrir)' -Y 352 -Altura 32 -Secundario -Cor $script:Cor.Titulo
$script:BtnDesfazer  = New-Botao -Texto 'Desfazer otimizacoes'    -Y 390 -Altura 32 -Secundario -Cor $script:Cor.Acao
$script:BtnReiniciar = New-Botao -Texto 'Reiniciar o computador'  -Y 428 -Altura 32 -Secundario -Cor $script:Cor.Alerta

$lblRodape = New-Object System.Windows.Forms.Label
$lblRodape.Text = "Ferramenta de apoio tecnico.`r`nNao validada para uso clinico.`r`nNenhuma acao exige administrador.`r`nTudo o que e desativado volta pelo Desfazer."
$lblRodape.SetBounds(18, 480, 244, 68); $lblRodape.ForeColor = $script:Cor.Fraco
$lblRodape.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$lateral.Controls.AddRange(@($script:BtnMod1, $script:BtnMod2, $script:BtnMod3,
                             $script:BtnGrandes, $script:BtnSessao, $script:BtnPersist,
                             $sec3, $script:BtnRestaurar, $script:BtnDesfazer, $script:BtnReiniciar, $lblRodape))

# --- painel de limpeza (direita) ---
function New-BotaoPainel {
    param([string]$Texto, [int]$X, [int]$Y, [int]$L, [int]$A, $Cor = $null, [switch]$Fraco)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Texto; $b.SetBounds($X, $Y, $L, $A); $b.FlatStyle = 'Flat'
    $b.BackColor = $(if ($Fraco) { $script:Cor.Painel } else { $script:Cor.Botao })
    $b.ForeColor = $(if ($Cor) { $Cor } else { $script:Cor.Texto })
    $b.FlatAppearance.BorderColor = $script:Cor.Borda
    return $b
}

# ---------------- painel: limpeza ----------------
$script:PainelLimpeza = New-Object System.Windows.Forms.Panel
$script:PainelLimpeza.Dock = 'Right'; $script:PainelLimpeza.Width = 520
$script:PainelLimpeza.BackColor = $script:Cor.Painel
$script:PainelLimpeza.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)
$script:PainelLimpeza.Visible = $false

$topoLimp = New-Object System.Windows.Forms.Panel
$topoLimp.Dock = 'Top'; $topoLimp.Height = 146; $topoLimp.BackColor = $script:Cor.Painel

$lblLimp = New-Object System.Windows.Forms.Label
$lblLimp.Text = 'O que vai ser aplicado'
$lblLimp.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblLimp.SetBounds(0, 12, 300, 20); $lblLimp.ForeColor = $script:Cor.Texto

$lblLimp2 = New-Object System.Windows.Forms.Label
$lblLimp2.Text = 'Tudo o que e seguro ja vem marcado. Desmarque o que nao quiser e clique em APLICAR TUDO.'
$lblLimp2.SetBounds(0, 34, 480, 32); $lblLimp2.ForeColor = $script:Cor.Fraco
$lblLimp2.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$lblDown = New-Object System.Windows.Forms.Label
$lblDown.Text = 'Remover o que tem mais de:'
$lblDown.SetBounds(0, 76, 180, 20); $lblDown.ForeColor = $script:Cor.Texto

$script:ComboPeriodo = New-Object System.Windows.Forms.ComboBox
$script:ComboPeriodo.SetBounds(184, 72, 290, 24)
$script:ComboPeriodo.DropDownStyle = 'DropDownList'
$script:ComboPeriodo.BackColor = $script:Cor.Fundo
$script:ComboPeriodo.ForeColor = $script:Cor.Texto
$script:ComboPeriodo.FlatStyle = 'Flat'
foreach ($op in @('3 dias', '7 dias', '15 dias', '30 dias', 'tudo, sem limite de idade')) { [void]$script:ComboPeriodo.Items.Add($op) }
$script:ComboPeriodo.SelectedIndex = 1

$btnMarcar    = New-BotaoPainel -Texto 'Marcar tudo'    -X 0   -Y 106 -L 150 -A 28
$btnDesmarcar = New-BotaoPainel -Texto 'Desmarcar tudo' -X 158 -Y 106 -L 150 -A 28
$topoLimp.Controls.AddRange(@($lblLimp, $lblLimp2, $lblDown, $script:ComboPeriodo, $btnMarcar, $btnDesmarcar))

$rodapeLimp = New-Object System.Windows.Forms.Panel
$rodapeLimp.Dock = 'Bottom'; $rodapeLimp.Height = 66; $rodapeLimp.BackColor = $script:Cor.Painel
$script:BtnLimparSel = New-BotaoPainel -Texto 'APLICAR TUDO' -X 0 -Y 10 -L 300 -A 44 -Cor $script:Cor.Ok
$script:BtnLimparSel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnFecharPainel = New-BotaoPainel -Texto 'Fechar' -X 312 -Y 10 -L 162 -A 44 -Fraco
$rodapeLimp.Controls.AddRange(@($script:BtnLimparSel, $btnFecharPainel))

$script:ListaLimpeza = New-Object System.Windows.Forms.CheckedListBox
$script:ListaLimpeza.Dock = 'Fill'
$script:ListaLimpeza.BackColor = $script:Cor.Fundo
$script:ListaLimpeza.ForeColor = $script:Cor.Texto
$script:ListaLimpeza.BorderStyle = 'FixedSingle'
$script:ListaLimpeza.CheckOnClick = $true
$script:ListaLimpeza.HorizontalScrollbar = $true
$script:ListaLimpeza.Font = New-Object System.Drawing.Font('Consolas', 9)

$script:PainelLimpeza.Controls.AddRange(@($script:ListaLimpeza, $rodapeLimp, $topoLimp))

# ---------------- painel: arquivos e pastas grandes ----------------
$script:PainelGrandes = New-Object System.Windows.Forms.Panel
$script:PainelGrandes.Dock = 'Right'; $script:PainelGrandes.Width = 520
$script:PainelGrandes.BackColor = $script:Cor.Painel
$script:PainelGrandes.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)
$script:PainelGrandes.Visible = $false

$topoGr = New-Object System.Windows.Forms.Panel
$topoGr.Dock = 'Top'; $topoGr.Height = 78; $topoGr.BackColor = $script:Cor.Painel
$lblGr = New-Object System.Windows.Forms.Label
$lblGr.Text = 'Arquivos e pastas grandes'
$lblGr.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblGr.SetBounds(0, 12, 320, 20); $lblGr.ForeColor = $script:Cor.Texto
$lblGr2 = New-Object System.Windows.Forms.Label
$lblGr2.Text = 'AVALIAR = pode render espaco · NAO APAGAR = deixe como esta. Duplo clique abre no Explorer.'
$lblGr2.SetBounds(0, 34, 480, 34); $lblGr2.ForeColor = $script:Cor.Fraco
$lblGr2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$topoGr.Controls.AddRange(@($lblGr, $lblGr2))

$rodapeGr = New-Object System.Windows.Forms.Panel
$rodapeGr.Dock = 'Bottom'; $rodapeGr.Height = 66; $rodapeGr.BackColor = $script:Cor.Painel
$btnAbrir = New-BotaoPainel -Texto 'ABRIR NO EXPLORER' -X 0 -Y 10 -L 300 -A 44 -Cor $script:Cor.Titulo
$btnAbrir.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnFecharGr = New-BotaoPainel -Texto 'Fechar' -X 312 -Y 10 -L 162 -A 44 -Fraco
$rodapeGr.Controls.AddRange(@($btnAbrir, $btnFecharGr))

$script:ListaGrandes = New-Object System.Windows.Forms.ListBox
$script:ListaGrandes.Dock = 'Fill'
$script:ListaGrandes.BackColor = $script:Cor.Fundo
$script:ListaGrandes.ForeColor = $script:Cor.Texto
$script:ListaGrandes.BorderStyle = 'FixedSingle'
$script:ListaGrandes.HorizontalScrollbar = $true
$script:ListaGrandes.Font = New-Object System.Drawing.Font('Consolas', 9)

$script:PainelGrandes.Controls.AddRange(@($script:ListaGrandes, $rodapeGr, $topoGr))

# ---------------- painel: status da sessao ----------------
$script:PainelSessao = New-Object System.Windows.Forms.Panel
$script:PainelSessao.Dock = 'Right'; $script:PainelSessao.Width = 520
$script:PainelSessao.BackColor = $script:Cor.Painel
$script:PainelSessao.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)
$script:PainelSessao.Visible = $false

$topoSes = New-Object System.Windows.Forms.Panel
$topoSes.Dock = 'Top'; $topoSes.Height = 72; $topoSes.BackColor = $script:Cor.Painel
$lblSes = New-Object System.Windows.Forms.Label
$lblSes.Text = 'Status da sessao'
$lblSes.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblSes.SetBounds(0, 12, 320, 20); $lblSes.ForeColor = $script:Cor.Texto
$lblSes2 = New-Object System.Windows.Forms.Label
$lblSes2.Text = 'Vale so para a sessao de agora. O proximo logon devolve a maquina ao estado normal.'
$lblSes2.SetBounds(0, 34, 480, 32); $lblSes2.ForeColor = $script:Cor.Fraco
$lblSes2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$topoSes.Controls.AddRange(@($lblSes, $lblSes2))

$rodapeSes = New-Object System.Windows.Forms.Panel
$rodapeSes.Dock = 'Bottom'; $rodapeSes.Height = 66; $rodapeSes.BackColor = $script:Cor.Painel
$script:BtnReverterSes = New-BotaoPainel -Texto 'REVERTER PARA O ORIGINAL' -X 0 -Y 10 -L 240 -A 44 -Cor $script:Cor.Alerta
$script:BtnReverterSes.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$btnAtualizarSes = New-BotaoPainel -Texto 'Atualizar' -X 252 -Y 10 -L 110 -A 44
$btnFecharSes = New-BotaoPainel -Texto 'Fechar' -X 374 -Y 10 -L 118 -A 44 -Fraco
$rodapeSes.Controls.AddRange(@($script:BtnReverterSes, $btnAtualizarSes, $btnFecharSes))

$script:TxtSessao = New-Object System.Windows.Forms.RichTextBox
$script:TxtSessao.Dock = 'Fill'
$script:TxtSessao.BackColor = $script:Cor.Fundo
$script:TxtSessao.ForeColor = $script:Cor.Texto
$script:TxtSessao.BorderStyle = 'FixedSingle'
$script:TxtSessao.ReadOnly = $true
$script:TxtSessao.WordWrap = $false
$script:TxtSessao.ScrollBars = 'Both'
$script:TxtSessao.DetectUrls = $false
$script:TxtSessao.Font = New-Object System.Drawing.Font('Consolas', 9.5)

$script:PainelSessao.Controls.AddRange(@($script:TxtSessao, $rodapeSes, $topoSes))

# --- console de log ---
$script:Log = New-Object System.Windows.Forms.RichTextBox
$script:Log.Dock = 'Fill'
$script:Log.BackColor = $script:Cor.Fundo
$script:Log.ForeColor = $script:Cor.Texto
$script:Log.Font = New-Object System.Drawing.Font('Consolas', 9.5)
$script:Log.BorderStyle = 'None'
$script:Log.ReadOnly = $true
$script:Log.WordWrap = $false
$script:Log.ScrollBars = 'Both'
$script:Log.DetectUrls = $false

$script:Form.Controls.AddRange(@($script:Log, $script:PainelSessao, $script:PainelGrandes, $script:PainelLimpeza, $lateral, $rodape, $cab))

# =====================================================================
# 9. EVENTOS
# =====================================================================
function Invoke-ComProtecao {
    param([scriptblock]$Bloco, [string]$Nome)
    if ($script:Ocupado) { return }
    $script:Cancelar = $false
    Set-Ocupado $true
    try { & $Bloco }
    catch { Write-Log ('Erro inesperado em {0}: {1}' -f $Nome, $_.Exception.Message) 'CRITICO' }
    finally {
        Set-Ocupado $false
        Set-Status 'Pronto.'
        $script:Cancelar = $false
    }
}

$script:BtnMod1.Add_Click({ Invoke-ComProtecao { Invoke-PrepararAmbiente } 'Modulo 1' })
$script:BtnMod2.Add_Click({ Invoke-ComProtecao { Invoke-InventarioDiagnostico } 'Modulo 2' })
$script:BtnMod3.Add_Click({ Invoke-ComProtecao { Invoke-Modulo3Analise } 'Modulo 3' })
$script:BtnGrandes.Add_Click({ Invoke-ComProtecao { Invoke-ArquivosGrandes } 'Modulo 4' })
$script:BtnSessao.Add_Click({ Invoke-ComProtecao { Invoke-Modulo7Sessao } 'Modulo 5' })
$script:BtnPersist.Add_Click({ Invoke-ComProtecao { Invoke-DiagPersistencia } 'Modulo 6' })
$script:BtnRestaurar.Add_Click({ Invoke-ComProtecao { Invoke-RestaurarSessao; Show-StatusSessao -Estado 'ORIGINAL' } 'Restaurar' })
$script:BtnLimparSel.Add_Click({ Invoke-Modulo3Limpeza })
$btnFecharPainel.Add_Click({ $script:PainelLimpeza.Visible = $false })

$btnMarcar.Add_Click({
    for ($i = 0; $i -lt $script:ListaLimpeza.Items.Count; $i++) { $script:ListaLimpeza.SetItemChecked($i, $true) }
})
$btnDesmarcar.Add_Click({
    for ($i = 0; $i -lt $script:ListaLimpeza.Items.Count; $i++) { $script:ListaLimpeza.SetItemChecked($i, $false) }
})
$script:ComboPeriodo.Add_SelectedIndexChanged({
    $script:DiasCorte = @(3, 7, 15, 30, 0)[$script:ComboPeriodo.SelectedIndex]
    if ($script:PainelLimpeza.Visible -and -not $script:Ocupado) {
        Write-Log ('Periodo alterado para: {0}. Recalculando...' -f $script:ComboPeriodo.SelectedItem) 'DADO'
        Invoke-ComProtecao { Invoke-Modulo3Analise } 'Modulo 3'
    }
})

$script:BtnReverterSes.Add_Click({
    Invoke-ComProtecao { Invoke-RestaurarSessao; Show-StatusSessao -Estado 'ORIGINAL' } 'Restaurar sessao'
})
$btnAtualizarSes.Add_Click({
    $estado = $(if ($script:StatusSessao) { 'OTIMIZADA' } else { 'ORIGINAL' })
    Show-StatusSessao -Estado $estado
})
$btnFecharSes.Add_Click({ $script:PainelSessao.Visible = $false })

$btnAbrir.Add_Click({ Open-ItemGrande })
$script:ListaGrandes.Add_DoubleClick({ Open-ItemGrande })
$btnFecharGr.Add_Click({ $script:PainelGrandes.Visible = $false })

$script:BtnDesfazer.Add_Click({ Invoke-ComProtecao { Restore-Otimizacoes } 'Desfazer' })
$script:BtnReiniciar.Add_Click({ Restart-Computador })

$script:BtnCancelar.Add_Click({
    $script:Cancelar = $true
    Set-Status 'Parando apos a etapa atual...'
})

$btnCopiar.Add_Click({
    try {
        [System.Windows.Forms.Clipboard]::SetText($script:Log.Text)
        Set-Status 'Log copiado. Cole no e-mail ou no chamado.'
    } catch { Set-Status 'Nao foi possivel copiar. Selecione o texto e use Ctrl+C.' }
})

$btnSalvar.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Arquivo de texto (*.txt)|*.txt'
    $dlg.FileName = ('Log_{0}_{1:yyyyMMdd_HHmm}.txt' -f $env:COMPUTERNAME, (Get-Date))
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:Log.Text | Out-File -FilePath $dlg.FileName -Encoding UTF8
            Set-Status ('Log salvo em ' + $dlg.FileName)
        } catch { Set-Status 'Nao foi possivel salvar o arquivo.' }
    }
})

$btnLimparLog.Add_Click({ $script:Log.Clear(); Set-Status 'Tela limpa.' })

$script:Form.Add_Shown({
    Write-Log 'Kit de Suporte pronto.' 'OK'
    Write-Log ('Maquina {0}  ·  usuario {1}  ·  {2:dd/MM/yyyy HH:mm}' -f $env:COMPUTERNAME, $env:USERNAME, (Get-Date)) 'DADO'
    Write-Log ''
    Write-Log 'Modulo 1: prepara o ambiente (pastas, portais, mapeamentos, PDFtk, fonte).' 'ACAO'
    Write-Log 'Modulo 2: inventario da maquina + diagnostico do que da para resolver aqui.' 'ACAO'
    Write-Log 'Modulo 3: analisa, deixa tudo marcado e resolve em um clique ("Aplicar tudo").' 'ACAO'
    Write-Log 'Modulo 4: os 10 maiores arquivos e as 10 maiores pastas do C:, com link para o Explorer.' 'ACAO'
    Write-Log 'Modulo 5: deixa a sessao de agora leve, preservando Citrix, prontuario e navegadores.' 'ACAO'
    Write-Log 'Citrix e Tasy sao testados dentro do Modulo 2.' 'DADO'
    Write-Log 'Este kit so faz o que roda sem administrador.' 'DADO'
    Write-Log 'Aplicativo autocontido: a rotina de preparacao vem dentro dele.' 'DADO'
})

[void]$script:Form.ShowDialog()
