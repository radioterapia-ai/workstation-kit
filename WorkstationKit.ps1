#requires -Version 5.1
#
#  WORKSTATION KIT  1.0
#  Diagnostics, cleanup and session optimisation for Windows.
#  Single file. No installer. No administrator privilege.
#
#  Nothing this tool does survives a restart, except the files you chose
#  to delete. It uninstalls nothing.
#
#  Support tool. Not validated for clinical use. See NOTICE.
#  https://github.com/radioterapia-ai/workstation-kit
#  Copyright 2026 Henrique Faria Braga - Apache License 2.0
#

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 } catch { }

$script:Versao       = '1.0'

$script:PastaClinica = ''

$script:PastaRelatorios = try { [Environment]::GetFolderPath('MyDocuments') } catch { $env:TEMP }

$script:MarcaEcossistema = 'RADIOTERAPIA_AI'

$script:SufixosDnsExtra = @()

$script:RaizesClinicas =
    'Vitrea|VitreaData|\\Patients\\|\\Patients$' +
    '|\\radioterapia\\TC_DATA' +
    '|Mirada|Medis|INVIA|Corridor4DM|NeuroQ|OleaSphere|TomTec' +
    '|\\ARIA|MOSAIQ|Monaco|Eclipse|RayStation|Velocity' +
    '|VspApp|VspMgmt' +
    '|Digitalcore|\\Onis'

$script:CitrixStores = @()
$script:TasyUrls = @()
$script:TasyDescobertas = $null

$script:Lim = @{
    DiscoCritPct    = 10
    DiscoAlertaPct  = 15
    RamCritPct      = 90
    RamAlertaPct    = 80

    MomentoRecemMin = 20
    UptimeAlertaH   = 72
    UptimeCritH     = 168
    CitrixMs        = 3000
    ArquivoGrandeMB = 100

    SessaoDesconhecidoMB = 300

    CacheFrescoMin = 15
}

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
$script:DiasCorte    = 7
$script:Cancelar     = $false

$script:SnapProc       = $null
$script:SnapSvc        = $null
$script:EmUsoOperacao  = $null

$script:SnapProcEm       = $null
$script:EmUsoOperacaoEm  = $null
$script:SnapMaxSeg       = 15

$script:SnapSessao       = @{}
$script:SnapGrupos       = $null
$script:Ocupado      = $false

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

function Pump { [System.Windows.Forms.Application]::DoEvents() }

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -lt 0)   { return 'n/d' }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} bytes' -f $Bytes)
}

$script:Idiomas = @(
    [pscustomobject]@{ Cod = 'en'; Nome = 'English' }
    [pscustomobject]@{ Cod = 'pt'; Nome = 'Portugues' }
    [pscustomobject]@{ Cod = 'es'; Nome = 'Espanol' }
)

$script:Textos = @{}
$script:Textos['pt'] = @{
    'idioma.01' = 'Idioma'
    'item.01' = 'OneDrive (sincronismo)'
    'nota.01' = 'Volta ao abrir o OneDrive ou no proximo logon.'
    'item.02' = 'Microsoft Teams'
    'nota.02' = 'As conversas ficam no servidor. Volta ao abrir.'
    'item.03' = 'Atualizadores automaticos'
    'nota.03' = 'Voltam sozinhos quando houver atualizacao.'
    'item.04' = 'Adobe Creative Cloud (fundo)'
    'item.05' = 'Xbox e Game Bar'
    'item.06' = 'Sobreposicao NVIDIA'
    'item.07' = 'Vincular ao telefone'
    'item.08' = 'Widgets do Windows'
    'item.09' = 'Pesquisa e Copilot'
    'nota.04' = 'O Windows recarrega sozinho quando precisar.'
    'item.10' = 'Assistentes do fabricante'
    'item.11' = 'Servicos Apple'
    'item.12' = 'Spotify'
    'item.13' = 'Nuvem pessoal'
    'item.14' = 'Comunicadores'
    'nota.05' = 'Encerrado apenas se nao houver janela aberta.'
    'item.15' = 'Lojas de jogos'
    'item.16' = 'OneDrive'
    'item.17' = 'Skype'
    'item.18' = 'Zoom'
    'item.19' = 'Comunicador'
    'item.20' = 'Loja de jogos'
    'item.21' = 'Componente Adobe'
    'item.22' = 'Atualizador automatico'
    'item.23' = 'Componente Apple'
    'item.24' = 'Assistente do fabricante'
    'item.25' = 'Extra do Windows'
    'item.26' = 'Jogos e sobreposicoes'
    'item.27' = 'Utilitario opcional'
    'item.28' = 'Edge abrindo sozinho no logon'
    'item.29' = 'protegido (seguranca, rede ou clinico)'
    'item.30' = 'nao reconhecido'
    'item.31' = 'Desligar animacoes, sombras e transparencia'
    'nota.06' = 'Deixa a resposta imediata em maquina fraca.'
    'item.32' = 'Impedir apps da Store de rodar em segundo plano'
    'nota.07' = 'Tira dezenas de processos do fundo.'
    'item.33' = 'Ligar a limpeza automatica do Windows'
    'nota.08' = 'Evita que o disco encha de novo.'
    'nota.09' = 'Faz o login unico funcionar e o .ica abrir sozinho.'
    'item.34' = 'Limpar a barra de tarefas (pesquisa, visao de tarefas, widgets)'
    'nota.10' = 'Some da barra e para de consumir memoria. Reversivel.'
    'item.35' = 'Tirar o icone do OneDrive do Explorer'
    'nota.11' = 'So esconde o atalho. Os arquivos e a conta continuam intactos.'
    'item.36' = 'Desligar sugestoes e apps promocionais do Windows'
    'nota.12' = 'Para de instalar aplicativo sozinho.'
    'alvo.01' = 'Lixeira'
    'item.37' = 'Registro do usuario (HKCU)'
    'item.38' = 'AppData\Local do perfil'
    'item.39' = 'AppData\Roaming do perfil'
    'item.40' = 'Disco fora do perfil'
    'alvo.02' = 'Arquivos temporarios do usuario'
    'obs.01' = 'Arquivos em uso sao ignorados automaticamente.'
    'alvo.03' = 'Cache do Microsoft Edge'
    'obs.02' = 'Nao apaga favoritos, senhas nem historico.'
    'alvo.04' = 'Cache do Google Chrome'
    'alvo.05' = 'Cache do Firefox'
    'obs.03' = 'Nao mexe no perfil com favoritos e senhas, que fica em Roaming.'
    'alvo.06' = 'Cache do Teams (versao classica)'
    'obs.04' = 'Feche o Teams antes. As conversas ficam no servidor.'
    'alvo.07' = 'Cache do Teams (versao nova)'
    'alvo.08' = 'Cache do Java Web Start (Tasy)'
    'obs.05' = 'Cache antigo do Java faz o Tasy falhar ao baixar arquivo. E baixado de novo no proximo acesso.'
    'alvo.09' = 'Cache do CentBrowser (navegador do prontuario)'
    'obs.06' = 'DESMARCADO de proposito: e o navegador do prontuario. Nao apaga favoritos nem senhas, so o cache. Marque se o Tasy estiver com tela em branco ou erro de arquivo.'
    'alvo.10' = 'Cache de imagens da Area de Trabalho Remota'
    'obs.07' = 'Miniaturas de tela das sessoes remotas. Sao recriadas na proxima conexao.'
    'obs.08' = 'Pasta de instalacao usada pelo TI. Arquivos em uso sao ignorados.'
    'alvo.11' = 'Cache de Opera e Vivaldi'
    'alvo.12' = 'Cookies legados do Windows (INetCookies)'
    'obs.09' = 'Pode pedir login de novo em sites internos antigos. Por isso vem desmarcado.'
    'alvo.13' = 'Cache de internet do Windows'
    'alvo.14' = 'Relatorios de erro do Windows'
    'alvo.15' = 'Despejos de memoria de travamentos'
    'alvo.16' = 'Cache de video (shaders)'
    'alvo.17' = 'Logs do OneDrive'
    'alvo.18' = 'Versoes antigas do Teams'
    'alvo.19' = 'Instaladores baixados do Edge/Chrome'
    'alvo.20' = 'Cache do Adobe Acrobat/Reader'
    'alvo.21' = 'Logs e diagnosticos do Office'
    'obs.10' = 'Apenas os lancadores baixados pelo navegador. O cache do Citrix nunca e tocado.'
    'alvo.22' = 'Cache de documentos do Office'
    'obs.11' = 'Feche todo o Office antes. Alteracoes ainda nao enviadas podem se perder.'
    'alvo.23' = 'Cache de miniaturas e icones'
    'obs.12' = 'As miniaturas sao recriadas sozinhas.'
    'alvo.24' = 'Lista de arquivos recentes'
    'obs.13' = 'Some a lista de recentes do Office e do Explorer. Nenhum arquivo e apagado.'
    'obs.14' = 'Feche o Excel antes. Sao arquivos de bloqueio que sobraram de sessoes travadas.'
    'alvo.25' = 'Esvaziar a Lixeira'
    'item.41' = 'Esvaziar a Lixeira (rotina)'
    'nota.13' = 'Feito antes dos Downloads, para que o que sair de la ainda possa ser recuperado.'
    'nota.14' = 'Vai para a Lixeira, da para recuperar. Troque o periodo na caixa acima da lista.'
    'item.42' = 'Dar prioridade de CPU ao Citrix, prontuario e Office'
    'nota.15' = 'Sobe o clinico para acima do normal e rebaixa o resto. Zera no proximo logon.'
    'item.43' = 'Compactar a memoria dos programas que ficam abertos'
    'nota.16' = 'Devolve ao Windows a memoria que o navegador e o Office reservaram sem usar. A sessao publicada do Citrix nao e mexida, para nao engasgar.'
    'item.44' = 'Google Drive'
    'item.45' = 'LibreOffice'
    'titulo.01' = '== {0} {1}'
    'achado.01' = '-> '
    'ajustezonascitrix.01' = '{0} adicionado aos sites da Intranet.'
    'diagcitrix.01' = 'Diagnostico do Citrix (ARIA, MOSAIQ e Monaco)'
    'diagcitrix.02' = 'Citrix Receiver antigo: versao {0}'
    'diagcitrix.03' = 'O Receiver 4.x saiu de suporte em 2018 e e substituido pelo Citrix Workspace. Atualizar exige o TI, mas resolve boa parte das falhas de abertura e de login unico.'
    'diagcitrix.04' = 'Citrix'
    'diagcitrix.05' = 'Citrix Workspace instalado (versao {0})'
    'diagcitrix.06' = 'Citrix Workspace nao encontrado nesta maquina'
    'diagcitrix.07' = 'Sem o cliente instalado o ARIA/MOSAIQ/Monaco nao abrem. Instalacao exige o TI.'
    'diagcitrix.08' = 'Citrix em execucao: {0} processo(s), {1} - {2}'
    'diagcitrix.09' = 'Ha sessao publicada aberta agora (ARIA, MOSAIQ ou Monaco)'
    'diagcitrix.10' = 'O kit nunca encerra o Citrix. Feche a sessao antes de limpar o cache do Workspace.'
    'diagcitrix.11' = 'Nenhum processo do Citrix em execucao.'
    'diagcitrix.12' = 'Stores ja configurados no Workspace deste usuario:'
    'diagcitrix.13' = 'Nenhum store gravado no perfil do usuario (acesso so pelo navegador).'
    'diagcitrix.14' = 'Endereco invalido na configuracao: {0}'
    'diagcitrix.15' = '--- {0} ---'
    'diagcitrix.16' = 'Aplicativos: {0}'
    'diagcitrix.17' = 'Endereco ..: {0}'
    'diagcitrix.18' = 'Testando '
    'diagcitrix.19' = '...'
    'diagcitrix.20' = 'DNS .......: {0} -> {1} ({2} ms)'
    'diagcitrix.21' = '{0}: o nome curto {1} nao resolve, so o completo {2}'
    'diagcitrix.22' = 'Troque o endereco do atalho e da configuracao para {0} . Sem isso o navegador nao acha o portal nesta estacao.'
    'diagcitrix.23' = '{0}: o nome {1} nao resolve nesta rede, nem com o dominio completo'
    'diagcitrix.24' = 'A estacao nao esta enxergando o DNS interno desse servidor. Confirme cabo/VPN e o sufixo DNS da placa antes de abrir chamado.'
    'diagcitrix.25' = 'Portas web : {0}'
    'diagcitrix.26' = 'nenhuma respondeu'
    'diagcitrix.27' = 'ICA 1494 ..: {0}   ·   Confiabilidade 2598: {1}{2}'
    'diagcitrix.28' = 'responde'
    'diagcitrix.29' = 'sem resposta'
    'diagcitrix.30' = '{0}: servidor {1} sem resposta nas portas web'
    'diagcitrix.31' = 'Enquanto isso nao voltar, os aplicativos desse destino nao abrem. Verifique a rede da estacao e, se as outras estacoes tambem falharem, avise o TI.'
    'diagcitrix.32' = '{0}: portal responde em {1} ms'
    'diagcitrix.33' = 'Acima de 3 segundos a lista de aplicativos demora a aparecer. Limpar o cache do Workspace no Modulo 3 costuma resolver do lado da estacao.'
    'diagcitrix.34' = '{0}: portal respondendo (HTTP {1}, {2} ms)'
    'diagcitrix.35' = '{0}: portal no ar, pedindo login (HTTP {1})'
    'diagcitrix.36' = '{0}: portal respondeu HTTP {1}'
    'diagcitrix.37' = 'O servidor esta no ar mas o store nao. Registre o codigo no chamado.'
    'diagcitrix.38' = '{0}: portal nao respondeu'
    'diagcitrix.39' = 'Detalhe: '
    'diagcitrix.40' = 'Zona ......: ja esta na Intranet (login unico funciona)'
    'diagcitrix.41' = '{0}: {1} nao esta na zona de Intranet'
    'diagcitrix.42' = 'E o que faz o navegador pedir senha de novo e nao abrir o arquivo .ica sozinho. O Modulo 3 corrige com um clique.'
    'diagcitrix.43' = 'Seguranca .: o servidor redireciona sozinho para HTTPS. O atalho pode ficar como esta.'
    'diagcitrix.44' = '{0}: o atalho usa HTTP, mas o servidor tambem atende em HTTPS'
    'diagcitrix.45' = 'Troque o atalho para {0} . Em HTTP o navegador bloqueia parte do login unico e o trafego vai sem criptografia.'
    'diagcitrix.46' = 'Seguranca .: endereco so em HTTP. O trafego vai sem criptografia.'
    'diagcitrix.47' = '{0}: certificado HTTPS nao confiavel nesta estacao'
    'diagcitrix.48' = 'O navegador mostra aviso de site nao seguro e pode bloquear download e impressao. Peca ao TI a instalacao do certificado da autoridade interna.'
    'diagcitrix.49' = 'Proxy ligado: {0}'
    'diagcitrix.50' = 'Excecoes ...: {0}'
    'diagcitrix.51' = 'Proxy sem excecao para: {0}'
    'diagcitrix.52' = 'O trafego do Citrix esta passando pelo proxy sem necessidade, o que deixa a abertura lenta. Peca ao TI para incluir nas excecoes.'
    'diagcitrix.53' = '{0,10}  {1}'
    'diagcitrix.54' = 'Arquivos .ica soltos: {0}'
    'diagcitrix.55' = 'Cache do Citrix com {0} - este kit nunca o limpa'
    'diagcitrix.56' = '{0} arquivos .ica soltos'
    'diagcitrix.57' = 'Sao apenas lancadores baixados pelo navegador. O Modulo 3 remove os de Downloads dentro do periodo escolhido; o cache do Citrix fica intacto.'
    'diagcitrix.58' = 'Se um destino falha e os outros funcionam, o problema e daquele servidor, nao da estacao.'
    'diagcitrix.59' = 'Se todos falham, e rede ou perfil da estacao.'
    'invidentificacao.01' = 'Inventario · identificacao'
    'invidentificacao.02' = 'Maquina .......: {0}   usuario: {1}\{2}'
    'invidentificacao.03' = 'Sistema .......: {0} build {1}'
    'invidentificacao.04' = 'Equipamento ...: {0} {1}   serie {2}'
    'invidentificacao.05' = 'Processador ...: {0} ({1} nucleos)'
    'invidentificacao.06' = 'Memoria .......: {0}'
    'invidentificacao.07' = 'Windows desde .: {0:dd/MM/yyyy}'
    'invidentificacao.08' = 'Memoria total de {0}'
    'invidentificacao.09' = 'Abaixo de 8 GB o Excel em rede mais o navegador ja saturam a maquina. Registre no chamado.'
    'invidentificacao.10' = 'Hardware'
    'invidentificacao.11' = '{0} -> {1}'
    'invsoftware.01' = 'Inventario · software instalado'
    'invsoftware.02' = '{0} programas instalados'
    'invsoftware.03' = '{0,-56} {1,-18} {2}'
    'invsoftware.04' = 'Software clinico:'
    'invsoftware.05' = '[X] {0,-14} {1}'
    'invsoftware.06' = '[ ] {0}'
    'invsoftware.07' = '[X] Papel de servidor: {0} ({1} processo(s), {2} servico(s))'
    'invsoftware.08' = 'Esta maquina hospeda servico (banco de dados ou aplicacao)'
    'invsoftware.09' = 'Limpeza e reinicio precisam de combinado previo com quem depende do servico.'
    'invsoftware.10' = 'Servidor'
    'invsoftware.11' = 'Agentes corporativos em execucao: {0}'
    'invinicializacaoco.01' = 'Inventario · tudo que abre com o Windows'
    'invinicializacaoco.02' = '{0} itens (usuario + maquina)'
    'invinicializacaoco.03' = '[{0,-13}] {1,-36} {2}'
    'invinicializacaoco.04' = '{0} itens de inicializacao sao da maquina, nao do seu usuario'
    'invinicializacaoco.05' = 'Esses so o TI desativa. Registre no chamado se o logon estiver longo.'
    'invinicializacaoco.06' = 'Inicializacao'
    'invservicos.01' = 'Inventario · servicos em execucao'
    'invservicos.02' = '{0} em execucao de {1} instalados'
    'invservicos.03' = '{0,-40} {1,-12} {2}'
    'invofficeexcel.01' = 'Inventario · Office e Excel'
    'invofficeexcel.02' = 'Versao do Office: {0}'
    'invofficeexcel.03' = 'Local confiavel: {0}'
    'invofficeexcel.04' = 'A pasta clinica nao esta nos locais confiaveis do Excel'
    'invofficeexcel.05' = 'Rode o Modulo 1 (Preparar ambiente) nesta maquina.'
    'invofficeexcel.06' = 'Excel'
    'invofficeexcel.07' = 'Suplemento: {0} (LoadBehavior {1})'
    'invofficeexcel.08' = 'Suplemento {0} carrega junto com o Excel'
    'invofficeexcel.09' = 'Arquivo > Opcoes > Suplementos > Suplementos de COM > desmarque. Atrasa a abertura de toda planilha.'
    'invpastaclinica.01' = 'Inventario · pasta clinica'
    'invpastaclinica.02' = '{0} nao existe'
    'invpastaclinica.03' = 'Verifique a pasta de trabalho configurada.'
    'invpastaclinica.04' = 'Ambiente'
    'invpastaclinica.05' = 'Total: {0}'
    'invpastaclinica.06' = '{0,-40} {1,10}'
    'invpastaclinica.07' = '{0,-40} {1,10}  {2:dd/MM/yyyy HH:mm}'
    'invpastaclinica.08' = 'Preparacao incompleta: falta {0}'
    'invredeimpressoras.01' = 'Inventario · rede e impressoras'
    'invredeimpressoras.02' = '{0,-32} {1,-12} {2}'
    'invredeimpressoras.03' = 'Estacao conectada por Wi-Fi ({0})'
    'invredeimpressoras.04' = 'Os portais tem 40 MB e sao abertos da rede. Cabo reduz de forma sensivel o tempo de abertura e evita corrupcao do arquivo.'
    'invredeimpressoras.05' = 'Rede'
    'invredeimpressoras.06' = 'Placa de rede negociando apenas {0}'
    'invredeimpressoras.07' = 'A placa e o switch sao de 1 Gbps: 100 Mbps quase sempre e cabo velho, mal crimpado ou porta ruim. Trocar o cabo custa nada e melhora Tasy, Citrix e a abertura dos portais de 40 MB.'
    'invredeimpressoras.08' = 'Sufixo DNS: {0}'
    'invredeimpressoras.09' = '{0}{1,-42} {2}'
    'invredeimpressoras.10' = 'Impressora padrao offline: {0}'
    'invredeimpressoras.11' = 'O Excel e o Word congelam ao abrir arquivo quando a impressora padrao nao responde. Troque a padrao para "Microsoft Print to PDF" enquanto nao resolvem.'
    'invredeimpressoras.12' = 'Impressora'
    'invmigracao.01' = 'Inventario · residuos de migracao de dominio'
    'invmigracao.02' = 'Agente de migracao instalado: {0} {1}'
    'invmigracao.03' = 'Agente de migracao de dominio ainda instalado ({0})'
    'invmigracao.04' = 'A migracao terminou ha anos. Remover exige administrador: peca ao TI a desinstalacao em massa. Ele roda servico e varre o perfil a cada logon.'
    'invmigracao.05' = 'Migracao'
    'invmigracao.06' = 'Nenhum agente de migracao instalado.'
    'invmigracao.07' = 'Usuario logado .....: {0}'
    'invmigracao.08' = 'Pasta do perfil ....: {0}'
    'invmigracao.09' = 'A pasta do perfil nao tem o nome do usuario: {0} usa a pasta {1}'
    'invmigracao.10' = 'Sinal de perfil herdado de outra conta, tipico de migracao de dominio. Funciona, mas confunde script, politica de grupo e permissao. Se esta maquina der problema estranho de perfil, e por aqui que se comeca.'
    'invmigracao.11' = 'Perfil com sufixo de dominio: {0}'
    'invmigracao.12' = 'Sinal de perfil recriado na migracao. Recriar de novo exige administrador.'
    'invmigracao.13' = '{0} credencial(is) salva(s) para servidores que nao existem mais:'
    'invmigracao.14' = '{0} credencial(is) salva(s) apontando para servidor inexistente'
    'invmigracao.15' = 'O Windows tenta usar cada uma ao abrir o Explorer e espera o tempo limite. O Modulo 3 remove - o item vem DESMARCADO porque nao tem desfazer.'
    'invmigracao.16' = 'Nenhuma credencial orfa no Gerenciador de Credenciais.'
    'inventario.01' = 'Relatorio'
    'inventario.02' = 'Servidor de logs fora de alcance: {0}'
    'inventario.03' = 'A planilha nao foi gravada. O relatorio e este log ("Copiar log" ou "Salvar log").'
    'inventario.04' = 'Pasta de logs indisponivel: {0}'
    'inventario.05' = 'Planilha ..: {0}'
    'inventario.06' = 'Gravada na pasta de LOGs da rede. O relatorio de texto e este proprio log.'
    'inventario.07' = 'Junte os .csv de varias maquinas numa planilha so para comparar o parque.'
    'inventario.08' = 'Nao foi possivel gravar o relatorio: {0}'
    'diagtasy.01' = 'Diagnostico do Tasy (prontuario eletronico)'
    'diagtasy.02' = 'Instalado: {0} {1}'
    'diagtasy.03' = 'Sem cliente instalado: o Tasy roda no navegador (Wheb HTML5).'
    'diagtasy.04' = 'Navegador {0,-14} {1,10} em {2} processo(s)'
    'diagtasy.05' = '{0} consumindo {1} em {2} processos'
    'diagtasy.06' = 'O Tasy roda dentro do navegador: com essa carga a tela trava mesmo com o servidor respondendo rapido. Feche as abas que nao usa antes de abrir chamado.'
    'diagtasy.07' = 'Tasy'
    'diagtasy.08' = 'Java em execucao: {0} · {1} · {2}'
    'diagtasy.09' = 'Cliente Java do Tasy rodando em 32 bits'
    'diagtasy.10' = 'Java 32 bits nao passa de ~1,5 GB de memoria. Relatorio grande trava ou fecha sozinho. Peca ao TI o Java 64 bits.'
    'diagtasy.11' = 'Nenhum endereco do Tasy configurado'
    'diagtasy.12' = 'Medindo '
    'diagtasy.13' = 'Tipo ......: {0}'
    'diagtasy.14' = 'servidor interno da rede'
    'diagtasy.15' = 'servico externo, sai pela internet'
    'diagtasy.16' = 'DNS .......: {0} ({1} ms)'
    'diagtasy.17' = '{0}: o nome {1} nao resolve nesta rede'
    'diagtasy.18' = 'A estacao nao esta enxergando o DNS. E problema de rede da estacao, nao do prontuario.'
    'diagtasy.19' = 'Portas ....: {0}'
    'diagtasy.20' = '{0}: servidor {1} sem resposta'
    'diagtasy.21' = 'Se as outras estacoes tambem nao abrem, o servidor caiu: chamado imediato.'
    'diagtasy.22' = 'Medindo o tempo de resposta (4 amostras)...'
    'diagtasy.23' = '{0}: portas abertas mas a pagina nao responde'
    'diagtasy.24' = 'O servidor web esta no ar e a aplicacao nao. Chamado com o horario exato.'
    'diagtasy.25' = 'Rede (TCP) ....: min {0} · tipico {1} · max {2} ms'
    'diagtasy.26' = 'Pagina (HTTP) .: min {0} · tipico {1} · max {2} ms   (HTTP {3})'
    'diagtasy.27' = 'Processamento .: ~{0} ms no servidor (HTTP menos rede)'
    'diagtasy.28' = 'Oscilacao: uma das amostras levou {0} ms contra {1} ms das outras. Rede instavel, nao lenta.'
    'diagtasy.29' = '{0}: pagina levando {1} ms para responder'
    'diagtasy.30' = 'A rede responde em {0} ms, entao a demora e do servidor de aplicacao. Anexe este log no chamado: nao adianta mexer na estacao.'
    'diagtasy.31' = 'Rede em {0} ms. A lentidao esta no servidor, nao na estacao.'
    'diagtasy.32' = '{0}: respondendo em {1} ms'
    'diagtasy.33' = '{0}: latencia tipica de {1} ms para um servidor interno'
    'diagtasy.34' = 'Acima de 60 ms dentro da rede interna e caminho ruim. Confira a velocidade negociada da placa e a porta do switch antes de culpar o servidor.'
    'diagtasy.35' = '{0}: resposta instavel (variou de {1} a {2} ms)'
    'diagtasy.36' = 'Oscilacao desse tamanho e o que faz a tela do Tasy "travar" de vez em quando. Registre no chamado com o horario.'
    'diagtasy.37' = 'Zona ......: fora da Intranet do Windows - sem efeito aqui, o Tasy tem login proprio.'
    'diagtasy.38' = 'COMPARATIVO ENTRE OS AMBIENTES'
    'diagtasy.39' = '{0,-34} {1,10} {2,10}  {3}'
    'diagtasy.40' = 'pagina'
    'diagtasy.41' = 'situacao'
    'diagtasy.42' = '{0} responde rapido e {1} nao, da mesma estacao e no mesmo minuto.'
    'diagtasy.43' = 'Isso descarta a estacao e a rede local: o problema esta no servidor lento.'
    'diagtasy.44' = 'Cache do Java Web Start: {0}'
    'diagtasy.45' = 'Cache do Java Web Start com {0}'
    'diagtasy.46' = 'Cache antigo faz o Tasy abrir versao errada e falhar ao baixar arquivo. O Modulo 3 limpa.'
    'diagtasy.47' = 'Cache do CentBrowser: {0}'
    'diagtasy.48' = 'Cache do CentBrowser com {0}'
    'diagtasy.49' = 'O Modulo 3 tem o item, mas vem DESMARCADO: e o navegador do prontuario e a decisao de limpar e sua.'
    'diagtasy.50' = 'Proxy sem excecao para o Tasy interno: {0}'
    'diagtasy.51' = 'O trafego do prontuario passa pelo proxy sem necessidade e isso soma tempo em cada tela. Peca ao TI para incluir nas excecoes.'
    'diagtasy.52' = 'COMO LER O RESULTADO'
    'diagtasy.53' = 'Rede baixa e pagina alta      -> servidor de aplicacao lento. Chamado, com estes numeros.'
    'diagtasy.54' = 'Rede alta e pagina alta       -> caminho de rede da estacao. Cabo em vez de Wi-Fi, e chamado de rede.'
    'diagtasy.55' = 'Tudo baixo e o Tasy travando  -> e a estacao: navegador com abas demais, cache velho ou memoria cheia.'
    'diagtasy.56' = 'Um ambiente rapido e outro lento -> o problema e daquele servidor, nao seu.'
    'diagtasy.57' = 'Peca ao TI a exclusao das pastas de cache do navegador e do Java na varredura do antivirus: e ganho direto na navegacao do Tasy.'
    'diagmemoria.01' = 'Memoria RAM'
    'diagmemoria.02' = 'Feche programas e abas do navegador. O computador esta usando disco como memoria.'
    'diagmemoria.03' = 'Feche o que nao estiver em uso, principalmente abas do navegador.'
    'diagmemoria.04' = 'Memoria total baixa para o uso clinico: {0}'
    'diagmemoria.05' = 'Abaixo de 8 GB, Excel + navegador + Teams travam o computador. Solicite ampliacao.'
    'diagmemoria.06' = 'Falha ao medir memoria: {0}'
    'diagmemoria.07' = 'Programas que mais consomem memoria:'
    'diagmemoria.08' = '{0,-24} {1,10}   ({2} processo(s))'
    'diagmemoria.09' = 'Feche as abas que nao estao em uso. Cada aba aberta e um processo com memoria propria.'
    'diagmemoria.10' = 'Teams consumindo muita memoria'
    'diagmemoria.11' = 'Feche e abra o Teams uma vez por dia, ou limpe o cache no Modulo 3.'
    'diagcaches.01' = 'Espaco recuperavel (previa da limpeza)'
    'diagcaches.02' = '{0} recuperaveis so com limpeza segura'
    'diagcaches.03' = 'Rode o Modulo 3 (Limpeza segura).'
    'diagcaches.04' = '{0} recuperaveis com limpeza segura'
    'diagsistemaacionav.01' = 'Estado da maquina'
    'diagsistemaacionav.02' = '{0} · build {1} · {2} de memoria'
    'diagsistemaacionav.03' = 'Reinicie o computador ao final da limpeza. Usar "Desligar" nao resolve: so "Reiniciar" limpa a memoria.'
    'diagsistemaacionav.04' = 'Reiniciar'
    'diagsistemaacionav.05' = 'Reinicie ao final da limpeza (botao no menu lateral).'
    'diagsistemaacionav.06' = 'Inicializacao Rapida ligada: "Desligar" nao limpa a memoria, so "Reiniciar".'
    'diagespaco.01' = 'Espaco em disco'
    'diagespaco.02' = 'Disco quase cheio e a causa numero um de lentidao. O Modulo 3 libera espaco agora.'
    'diagespaco.03' = 'Disco'
    'diagespaco.04' = 'Rode o Modulo 3 para liberar espaco.'
    'diagespaco.05' = 'Nao foi possivel ler as unidades.'
    'diagespaco.06' = 'Pasta Temp do Windows com {0}'
    'diagespaco.07' = 'Pasta do sistema: so o TI limpa, e a rotina de limpeza deles ja cobre. Registre o tamanho no chamado.'
    'diagespaco.08' = 'Pasta Temp do Windows: {0} (do sistema, fora do alcance do usuario)'
    'diaglixeira.01' = 'Lixeira'
    'diaglixeira.02' = 'Lixeira com {0}'
    'diaglixeira.03' = 'Esvaziar a Lixeira e rotina do Modulo 3. Confira antes se nao ha nada a recuperar.'
    'diagencerraveis.01' = 'Programas que podem ser encerrados agora com seguranca'
    'diagencerraveis.02' = 'Criterio: nao guardam documento aberto e voltam sozinhos quando voce abrir de novo.'
    'diagencerraveis.03' = 'Nenhum programa dispensavel rodando agora'
    'diagencerraveis.04' = 'Memoria'
    'diagencerraveis.05' = 'Nada a encerrar: o que esta aberto e trabalho ou sistema.'
    'diagencerraveis.06' = '{0,-34} {1,10}   ({2} processo(s))'
    'diagencerraveis.07' = '{0} de memoria presos em {1} programa(s) dispensavel(is)'
    'diagencerraveis.08' = 'O Modulo 3 encerra todos de uma vez. Eles voltam quando voce abrir.'
    'diagencerraveis.09' = '{0} recuperaveis encerrando programas dispensaveis'
    'diagencerraveis.10' = 'Programas de trabalho, clinicos, de seguranca e navegadores nunca sao encerrados pelo kit.'
    'diaginicioacionave.01' = 'Inicializacao do Windows'
    'diaginicioacionave.02' = 'A maquina ficou {0:N1} h ligada antes deste logon.'
    'diaginicioacionave.03' = 'Por isso nao da para medir o tempo de inicializacao nesta sessao.'
    'diaginicioacionave.04' = 'Para medir: reinicie e rode o Modulo 2 logo depois de entrar.'
    'diaginicioacionave.05' = 'Inicializacao levou {0} segundos'
    'diaginicioacionave.06' = 'Veja abaixo o que pesa: itens de inicializacao, unidades de rede e impressoras de rede. O Modulo 3 resolve a parte do usuario.'
    'diaginicioacionave.07' = 'O Modulo 3 reduz isso desativando os itens seguros.'
    'diaginicioacionave.08' = 'Inicializacao em {0} segundos'
    'diaginicioacionave.09' = 'Itens do seu usuario que abrem junto com o Windows: {0}'
    'diaginicioacionave.10' = 'Seguro NAO carregar de novo (o programa continua funcionando ao ser aberto):'
    'diaginicioacionave.11' = '[ desativar ] {0,-30} {1}'
    'diaginicioacionave.12' = 'Mantidos como estao:'
    'diaginicioacionave.13' = '[  manter  ] {0,-30} {1}'
    'diaginicioacionave.14' = 'O que mais pesa no logon desta estacao:'
    'diaginicioacionave.15' = '{0} unidade(s) de rede e {1} impressora(s) de rede sao reconectadas a cada logon.'
    'diaginicioacionave.16' = '{0} -> {1}   {2}'
    'diaginicioacionave.17' = 'impressora: {0}'
    'diaginicioacionave.18' = '{0} unidade(s) de rede sem resposta agora: {1}'
    'diaginicioacionave.19' = 'Cada uma dessas o Windows tenta reconectar no logon e espera o tempo limite. E a explicacao mais provavel do logon longo. Peca ao TI para remover o mapeamento ou liberar o acesso.'
    'diaginicioacionave.20' = 'Logon longo com apenas {0} item(ns) de inicializacao do usuario'
    'diaginicioacionave.21' = 'O tempo nao vem dos programas do seu usuario. Sao {0} unidade(s) de rede e {1} impressora(s) de rede reconectadas a cada logon, mais os agentes corporativos. Leve estes numeros ao TI.'
    'diaginicioacionave.22' = '{0} programas abrem sozinhos e podem ser desativados com seguranca'
    'diaginicioacionave.23' = 'Cada um deles disputa disco e CPU no logon. O Modulo 3 desativa todos de uma vez e o botao "Desfazer" reverte.'
    'diaginicioacionave.24' = '{0} programa(s) na inicializacao podem ser desativados'
    'diaginicioacionave.25' = 'O Modulo 3 desativa.'
    'diaginicioacionave.26' = 'Inicializacao do usuario ja esta enxuta'
    'diaginicioacionave.27' = 'Os itens da maquina (para todos os usuarios) e os de seguranca so o TI altera.'
    'diagajustespendent.01' = 'Ajustes de desempenho pendentes'
    'diagajustespendent.02' = 'Todos os ajustes de desempenho ja estao aplicados'
    'diagajustespendent.03' = 'Ajustes'
    'diagajustespendent.04' = '[ aplicar ] {0}'
    'diagajustespendent.05' = '{0} ajuste(s) de desempenho ainda nao aplicado(s)'
    'diagajustespendent.06' = 'O Modulo 3 aplica todos. Reversivel pelo botao "Desfazer otimizacoes".'
    'diagmapeamentos.01' = 'Unidades de rede'
    'diagmapeamentos.02' = 'Nenhuma unidade de rede desconectada'
    'diagmapeamentos.03' = '{0} -> {1}  [desconectada]'
    'diagmapeamentos.04' = '{0} unidade(s) de rede desconectada(s)'
    'diagmapeamentos.05' = 'Mapeamento morto congela o Explorer e o "Salvar como" do Office por 30 segundos ou mais. O Modulo 3 remove.'
    'diagpersistencia.01' = 'Persistencia - o que sobrevive ao logoff'
    'diagpersistencia.02' = 'Perfil ....: {0}'
    'diagpersistencia.03' = 'Copia movel: {0}'
    'diagpersistencia.04' = 'Perfil obrigatorio: o Windows descarta suas alteracoes em todo logoff'
    'diagpersistencia.05' = 'Nada que o Modulo 3 ajusta no registro sobrevive, e o Desfazer nunca acha o que desfazer. So o TI muda isso - leve este log ao chamado.'
    'diagpersistencia.06' = 'Persistencia'
    'diagpersistencia.07' = 'Perfil temporario: esta sessao inteira e descartada no logoff'
    'diagpersistencia.08' = 'O Windows nao conseguiu abrir o seu perfil e criou um descartavel. Nao adianta ajustar nada agora. Chamado para o TI recriar o perfil.'
    'diagpersistencia.09' = 'Politica apaga a copia local do perfil no logoff'
    'diagpersistencia.10' = 'O que nao subir para o servidor a tempo se perde. Vale conferir com o TI se o perfil movel esta salvando.'
    'diagpersistencia.11' = 'Perfil movel: o que vale e a copia do servidor'
    'diagpersistencia.12' = 'Se o logoff nao terminar de sincronizar, o ajuste se perde. Feche os programas antes de sair.'
    'diagpersistencia.13' = 'Perfil local: o registro do usuario deve sobreviver ao logoff'
    'diagpersistencia.14' = 'Filtro de escrita no disco: {0}'
    'diagpersistencia.15' = 'O disco volta ao estado anterior a cada reinicio. Nada instalado ou ajustado permanece. So o TI desliga - leve este log ao chamado.'
    'diagpersistencia.16' = 'Nenhum marcador anterior. Deixando um agora em cada lugar.'
    'diagpersistencia.17' = 'Faca logoff (ou reinicie) e rode este modulo de novo: ele dira exatamente o que sobreviveu.'
    'diagpersistencia.18' = 'Marcador de {0}, da mesma sessao de logon. Ainda nao houve logoff para comparar.'
    'diagpersistencia.19' = 'Faca logoff (ou reinicie) e rode de novo.'
    'diagpersistencia.20' = 'Marcador anterior de {0}. Houve logoff desde entao - este e o resultado:'
    'diagpersistencia.21' = '   SOBREVIVEU  {0,-30} guarda {1}'
    'diagpersistencia.22' = '   PERDEU      {0,-30} perde {1}'
    'diagpersistencia.23' = 'Tudo sobreviveu ao ultimo logoff'
    'diagpersistencia.24' = 'Nem o disco fora do perfil sobreviveu ao logoff'
    'diagpersistencia.25' = 'Isso e filtro de escrita ou congelador de disco, nao problema de perfil. So o TI desliga.'
    'diagpersistencia.26' = 'O perfil nao guarda alteracao: {0} de {1} lugares perderam o marcador'
    'diagpersistencia.27' = 'Perdeu: '
    'diagpersistencia.28' = '. O disco fora do perfil ficou. Isso e perfil descartado, e so o TI corrige.'
    'diagpersistencia.29' = 'Marcador renovado em {0} de {1} lugares.'
    'diagpersistencia.30' = 'Algum lugar nem aceitou gravar agora - o que ja e resposta.'
    'invecossistema.01' = 'Aplicativos do radioterapia.ai'
    'invecossistema.02' = 'Nenhum registro do ecossistema nesta estacao. E o estado normal de quem nao instalou nada.'
    'invecossistema.03' = 'Registro ..: {0}'
    'invecossistema.04' = 'Registro do ecossistema ilegivel: {0}'
    'invecossistema.05' = 'Enquanto isso durar nenhum aplicativo consegue registrar versao, e o suporte fica sem saber o que esta instalado. Leve este log a quem cuida do instalador.'
    'invecossistema.06' = 'Ecossistema'
    'invecossistema.07' = '{0} aplicativo(s) declarado(s):'
    'invecossistema.08' = '   {0,-24} v{1,-10} {2,-12} {3}'
    'invecossistema.09' = '      {0}'
    'invecossistema.10' = '{0} entrada(s) do registro apontam para pasta que nao existe'
    'invecossistema.11' = 'O registro declara mas o disco nao tem: '
    'invecossistema.12' = '. Registro que mente e pior que registro vazio: vazio faz procurar, errado faz concluir. Leve ao responsavel pelo instalador - o kit nao corrige registro de outro projeto.'
    'invecossistema.13' = '{0} aplicativo(s) instalado(s) e ausente(s) do registro'
    'invecossistema.14' = 'Tem entrega.txt no disco mas nao aparece no registro: '
    'invecossistema.15' = '. E a mentira na direcao oposta, e esconde do suporte a versao que esta em uso.'
    'invecossistema.16' = 'Registro do ecossistema confere com o disco'
    'inventariodiagnost.01' = 'Modulo 2 - Inventario e diagnostico'
    'inventariodiagnost.02' = 'Maquina {0} · usuario {1} · {2:dd/MM/yyyy HH:mm:ss}'
    'inventariodiagnost.03' = 'Parte 1: retrato da maquina. Parte 2: o que o Modulo 3 resolve sem administrador.'
    'inventariodiagnost.04' = 'Diagnostico interrompido pelo usuario.'
    'inventariodiagnost.05' = 'Diagnostico {0}/{1}: {2}'
    'inventariodiagnost.06' = 'Falha na etapa {0}: {1}'
    'inventariodiagnost.07' = 'Mapa dos pontos criticos'
    'inventariodiagnost.08' = 'Nada a corrigir por aqui: a maquina esta limpa do lado do usuario.'
    'inventariodiagnost.09' = 'Se a lentidao continuar, o gargalo exige administrador (disco, antivirus, rede, hardware).'
    'inventariodiagnost.10' = 'Copie este log e abra chamado: ele ja mostra que o lado do usuario esta limpo.'
    'inventariodiagnost.11' = 'Onde esta o problema, por area:'
    'inventariodiagnost.12' = '{0,-16} {1,-26} {2} achado(s), {3} de alto impacto'
    'inventariodiagnost.13' = '{0} problema(s) critico(s) e {1} alerta(s).'
    'inventariodiagnost.14' = 'RESOLVA NESTA ORDEM:'
    'inventariodiagnost.15' = '{0}. [{1}/{2}] {3}'
    'inventariodiagnost.16' = 'Tudo o que esta acima e resolvido pelo Modulo 3 (Limpeza). Abra e clique em "Aplicar tudo".'
    'inventariodiagnost.17' = 'O que precisa de administrador esta marcado como tal: use "Copiar log" no chamado.'
    'inventariodiagnost.18' = 'Este diagnostico ja incluiu os testes de Citrix (tres destinos) e de Tasy (tres ambientes).'
    'planolimpeza.01' = 'Medindo a Lixeira...'
    'planolimpeza.02' = 'Mapeamento de rede sem resposta: {0} aponta para {1}'
    'planolimpeza.03' = 'Para remover, rode no Prompt de Comando:  net use {0} /delete'
    'planolimpeza.04' = 'Credencial salva de servidor que nao responde: {0}'
    'planolimpeza.05' = 'Para remover, rode:  cmdkey /delete:{0}   -- confira o nome antes: se for erro de DNS, o servidor existe e a credencial e valida.'
    'planolimpeza.06' = 'Credenciais'
    'modulo3analise.01' = 'Modulo 3 - Limpeza (analise)'
    'modulo3analise.02' = 'Nada e alterado nesta etapa. Tudo o que e seguro ja vem marcado.'
    'modulo3analise.03' = 'O que vai acontecer ao clicar em "Aplicar tudo":'
    'modulo3analise.04' = 'Espaco liberado em disco ...... {0}'
    'modulo3analise.05' = 'Memoria devolvida agora ....... {0}'
    'modulo3analise.06' = 'Itens fora da inicializacao ... {0}'
    'modulo3analise.07' = 'Ajustes de desempenho ......... {0}'
    'modulo3analise.08' = 'Mapeamentos mortos removidos .. {0}'
    'modulo3analise.09' = '{0} item(ns) de inicializacao ficaram desmarcados por nao serem reconhecidos:'
    'modulo3analise.10' = 'Marque na lista se souber que pode desativar.'
    'modulo3analise.11' = 'Desmarque o que nao quiser e clique em "Aplicar tudo".'
    'modulo3limpeza.01' = 'Rode primeiro o Modulo 3 para analisar a maquina.'
    'modulo3limpeza.02' = 'Limpeza'
    'modulo3limpeza.03' = 'Information'
    'modulo3limpeza.04' = 'Nenhum item marcado.'
    'modulo3limpeza.05' = 'Aplicar limpeza'
    'modulo3limpeza.06' = 'YesNo'
    'modulo3limpeza.07' = 'Question'
    'modulo3limpeza.08' = 'Limpeza cancelada pelo usuario.'
    'modulo3limpeza.09' = 'SESSAO'
    'modulo3limpeza.10' = 'Modulo 5 - Aplicando na sessao'
    'modulo3limpeza.11' = 'Modulo 3 - Aplicando'
    'modulo3limpeza.12' = 'Encerrando programas dispensaveis...'
    'modulo3limpeza.13' = 'Limpando caches e temporarios...'
    'modulo3limpeza.14' = 'Esvaziando a Lixeira...'
    'modulo3limpeza.15' = 'Enviando Downloads antigos para a Lixeira...'
    'modulo3limpeza.16' = 'Tirando programas da inicializacao...'
    'modulo3limpeza.17' = 'Aplicando ajustes de desempenho...'
    'modulo3limpeza.18' = 'Encerrando programas dispensaveis da sessao...'
    'modulo3limpeza.19' = 'Parando tarefas agendadas em execucao...'
    'modulo3limpeza.20' = 'Ajustando prioridade de CPU...'
    'modulo3limpeza.21' = 'Compactando a memoria dos programas que ficam...'
    'modulo3limpeza.22' = 'Interrompido pelo usuario.'
    'modulo3limpeza.23' = 'Aplicando {0}/{1}: {2}'
    'modulo3limpeza.24' = '{0}: {1} processo(s) encerrado(s), {2} devolvidos'
    'modulo3limpeza.25' = '{0}: ja nao estava em execucao'
    'modulo3limpeza.26' = '{0}: {1}'
    'modulo3limpeza.27' = '{0}: nada a limpar'
    'modulo3limpeza.28' = '{0}: {1} liberados{2}'
    'modulo3limpeza.29' = 'Lixeira: {0}'
    'modulo3limpeza.30' = 'Lixeira: ja estava vazia'
    'modulo3limpeza.31' = 'Lixeira esvaziada: {0} liberados'
    'modulo3limpeza.32' = 'Downloads: {0} itens para a Lixeira, {1} liberados'
    'modulo3limpeza.33' = '{0} item(ns) ja tinham sido removidos por outra etapa.'
    'modulo3limpeza.34' = '{0} item(ns) em uso foram mantidos. Exemplos:'
    'modulo3limpeza.35' = '   ... e outros {0}.'
    'modulo3limpeza.36' = 'Confira a Lixeira antes de esvaziar de novo, caso queira recuperar algo.'
    'modulo3limpeza.37' = 'REGISTRO: a Area de Trabalho ou Documentos deste usuario estao dentro do OneDrive.'
    'modulo3limpeza.38' = 'Com o OneDrive fora da inicializacao, os arquivos so sobem para a nuvem quando ele for aberto.'
    'modulo3limpeza.39' = 'Nao vai mais abrir sozinho: {0}'
    'modulo3limpeza.40' = 'Nao foi possivel desativar {0}'
    'modulo3limpeza.41' = 'Aplicado: {0}'
    'modulo3limpeza.42' = 'Resultado'
    'modulo3limpeza.43' = 'Memoria devolvida ............. {0}'
    'modulo3limpeza.44' = 'Memoria compactada ............ {0}'
    'modulo3limpeza.45' = 'Total devolvido ............... {0}'
    'modulo3limpeza.46' = 'Memoria em uso agora .......... {0}% ({1} livres)'
    'modulo3limpeza.47' = 'A sessao de agora esta mais leve. O painel a direita mostra o status.'
    'modulo3limpeza.48' = 'Para voltar agora, sem reiniciar: botao "REVERTER PARA O ORIGINAL".'
    'modulo3limpeza.49' = 'Espaco liberado ............... {0}'
    'modulo3limpeza.50' = 'Livre em C: agora ............. {0} ({1}%)'
    'modulo3limpeza.51' = 'Tudo o que foi desativado volta pelo botao "Desfazer otimizacoes".'
    'modulo3limpeza.52' = 'Limpeza aplicada.

O ganho na inicializacao e nos ajustes so aparece depois de reiniciar.

Quer reiniciar agora? Salve seus arquivos antes.'
    'modulo3limpeza.53' = 'Reiniciar agora'
    'modulo3limpeza.54' = 'Reiniciar depois'
    'modulo3limpeza.55' = 'Reinicio agendado para 15 segundos. Use shutdown /a no Executar para cancelar.'
    'modulo3limpeza.56' = 'Erro durante a limpeza: {0}'
    'modulo3limpeza.57' = 'Limpeza concluida.'
    'modulo3limpeza.58' = 'Encerrando {0}...'
    'modulo3limpeza.59' = '{0}: nao foi possivel encerrar (protegido ou ja fechado)'
    'modulo3limpeza.60' = 'Tarefa parada: {0}{1}'
    'modulo3limpeza.61' = '{0}: sem permissao para parar (roda como SISTEMA - so o TI)'
    'restartcomputador.01' = 'Reiniciar o computador'
    'restartcomputador.02' = 'Warning'
    'restartcomputador.03' = 'Reiniciando o computador...'
    'restartcomputador.04' = 'Reinicio agendado para 10 segundos. Use shutdown /a para cancelar.'
    'restartcomputador.05' = 'O Windows nao aceitou o pedido de reinicio (codigo {0}). A maquina NAO vai reiniciar.'
    'restartcomputador.06' = 'Reinicie pelo menu Iniciar. Use "Reiniciar", nao "Desligar".'
    'diagdownloads.01' = 'Pasta Downloads'
    'diagdownloads.02' = 'Pasta Downloads nao encontrada.'
    'diagdownloads.03' = 'Pasta Downloads vazia'
    'diagdownloads.04' = '{0} arquivos · {1} no total'
    'diagdownloads.05' = '{0} arquivos com mais de 90 dias · {1}'
    'diagdownloads.06' = '{0,-10} {1,10}  ({2} arquivo(s))'
    'diagdownloads.07' = 'sem ext'
    'diagdownloads.08' = 'Os 10 maiores:'
    'diagdownloads.09' = '{0,10}  {1}  ({2:dd/MM/yyyy})'
    'diagdownloads.10' = 'Downloads ocupando {0}, sendo {1} com mais de 90 dias'
    'diagdownloads.11' = 'Use o botao "Esvaziar Downloads" no menu lateral. Vai para a Lixeira, da para recuperar.'
    'diagdownloads.12' = 'Downloads com {0}'
    'diagnuvem.01' = 'OneDrive e Teams'
    'diagnuvem.02' = 'OneDrive rodando · {0}'
    'diagnuvem.03' = 'Pasta do OneDrive: {0} · {1}'
    'diagnuvem.04' = 'OneDrive sincronizando {0} no disco local'
    'diagnuvem.05' = 'Ligue "Arquivos sob demanda" (clique no icone da nuvem > engrenagem > Configuracoes) para liberar espaco sem perder acesso.'
    'diagnuvem.06' = 'Nuvem'
    'diagnuvem.07' = 'OneDrive consumindo {0}'
    'diagnuvem.08' = 'Use o botao "Encerrar OneDrive agora" quando precisar de desempenho. Ele volta ao ser aberto de novo.'
    'diagnuvem.09' = 'OneDrive nao esta em execucao.'
    'diagnuvem.10' = 'Teams rodando · {0} em {1} processo(s)'
    'diagnuvem.11' = 'Teams consumindo {0}'
    'diagnuvem.12' = 'Use "Encerrar Teams agora" durante o trabalho pesado nas planilhas. Ele volta ao ser aberto de novo.'
    'diagnuvem.13' = 'Teams nao esta em execucao.'
    'diagnuvem.14' = 'Versoes antigas do Teams ocupando {0}'
    'diagnuvem.15' = 'O Modulo 3 remove com o item "Versoes antigas do Teams".'
    'otimizacoes.01' = 'Desfazer otimizacoes'
    'otimizacoes.02' = 'Nenhuma otimizacao registrada para desfazer.'
    'otimizacoes.03' = 'Desfazer'
    'otimizacoes.04' = 'Isso devolve os efeitos visuais, os aplicativos em segundo plano e a inicializacao automatica ao estado anterior.

Continuar?'
    'otimizacoes.05' = 'Desfazer tudo'
    'otimizacoes.06' = 'Cancelar'
    'otimizacoes.07' = 'Restaurado: {0}'
    'otimizacoes.08' = 'Reativado na inicializacao: {0}'
    'otimizacoes.09' = 'Inicializacao do {0} restaurada.'
    'otimizacoes.10' = 'Zona removida: {0}'
    'otimizacoes.11' = 'Tudo restaurado. Faca logoff para aplicar por completo.'
    'ajusteregistro.01' = '{0}: bloqueado por politica do TI (o Windows nao deixa o usuario mudar).'
    'ajusteregistro.02' = '{0}: nao foi possivel ajustar.'
    'ajusteregistro.03' = '{0} de {1} ajustes deste grupo estao travados por politica. Os demais foram aplicados.'
    'explorer.01' = 'Explorer reiniciado: a barra de tarefas ja aparece limpa.'
    'explorer.02' = 'Nao foi possivel reiniciar o Explorer (bloqueio do antivirus).'
    'explorer.03' = 'As mudancas da barra de tarefas aparecem no proximo logon.'
    'arquivosgrandes.01' = 'Modulo 4 - Arquivos e pastas grandes'
    'arquivosgrandes.02' = 'Somente leitura. O kit nao apaga nada aqui: voce decide item por item.'
    'arquivosgrandes.03' = 'So o disco C:. E o unico cujo espaco livre afeta o desempenho do Windows.'
    'arquivosgrandes.04' = 'Pastas do Windows e Program Files ficam de fora: nao ha o que fazer nelas sem administrador.'
    'arquivosgrandes.05' = 'Paginacao do Windows fica em {0} ({1}% livre). Enquanto esse disco tiver espaco, ele nao pesa no desempenho.'
    'arquivosgrandes.06' = 'Disco {0} da paginacao com apenas {1}% livre'
    'arquivosgrandes.07' = 'A paginacao mora nesse disco: se ele encher, a maquina inteira trava. Libere espaco nele tambem.'
    'arquivosgrandes.08' = 'Procurando arquivos grandes...'
    'arquivosgrandes.09' = 'Arquivos grandes em '
    'arquivosgrandes.10' = 'OS 10 MAIORES ARQUIVOS'
    'arquivosgrandes.11' = 'Nenhum arquivo acima de {0} MB.'
    'arquivosgrandes.12' = '{0,2}. [{1,-15}] {2,10}  {3}'
    'arquivosgrandes.13' = 'OS 10 MAIORES DIRETORIOS'
    'arquivosgrandes.14' = 'Medindo pasta {0}/{1}: {2}'
    'arquivosgrandes.15' = 'PODE APAGAR ..... {0} em {1} item(ns) - descartavel, o Windows recria ou nao usa mais'
    'arquivosgrandes.16' = 'ATENCAO-AVALIAR . {0} em {1} item(ns) - confira o conteudo antes de mover ou apagar'
    'arquivosgrandes.17' = 'NAO APAGAR ...... dado clinico, banco, navegador do prontuario ou arquivo de sistema'
    'arquivosgrandes.18' = 'Nenhum diretorio recebe "PODE APAGAR": pasta inteira sempre exige conferencia.'
    'arquivosgrandes.19' = 'Selecione qualquer linha no painel a direita e clique em "Abrir no Explorer".'
    'openitemgrande.01' = 'Selecione uma linha da lista.'
    'openitemgrande.02' = 'Abrir no Explorer'
    'openitemgrande.03' = 'Nao existe mais: {0}'
    'openitemgrande.04' = 'Aberto no Explorer: {0}'
    'openitemgrande.05' = 'Nao foi possivel abrir: {0}'
    'snapshotsvc.01' = 'Nao foi possivel ler a lista de servicos desta maquina: a verificacao de papel de servidor NAO foi feita.'
    'logemuso.01' = 'Sinalizacao de app em uso encontrada mas VELHA ({0:N0} min). Ignorada.'
    'logemuso.02' = '{0}: {1} protege {2} processo(s), {3}.'
    'logemuso.03' = '   Nenhum deles sera encerrado, rebaixado ou compactado. Porta {0}.'
    'logemuso.04' = '   A data do arquivo nao pode ser lida. Se o aplicativo nao estiver aberto, apague:'
    'momentosessao.01' = 'Este kit e um app de Windows: ele prepara a maquina para o trabalho comecar.'
    'tarefasemexecucao.01' = 'Lendo tarefas agendadas (alguns segundos)...'
    'diagsessaoativa.01' = 'Sessao atual: o que esta ativo agora'
    'diagsessaoativa.02' = 'Nao foi possivel mapear os processos da sessao.'
    'diagsessaoativa.03' = 'Medido na sessao {0}, de {1} sessao(oes) ativa(s) nesta maquina.'
    'diagsessaoativa.04' = '{0,-46} {1,4} processo(s)  {2,10}'
    'diagsessaoativa.05' = 'ATENCAO: ha carga do radioterapia.ai em OUTRA sessao ({0} processo(s)):'
    'diagsessaoativa.06' = '   sessao {0,-4} {1,-16} PID {2,-7} {3,10}'
    'diagsessaoativa.07' = 'Carga do ecossistema em execucao (protegida dos Modulos 3 e 5):'
    'diagsessaoativa.08' = '   {0,-16} PID {1,-7} {2,10}   {3}'
    'diagsessaoativa.09' = 'Trabalho do radioterapia.ai em andamento {0}: {1} usando {2}'
    'diagsessaoativa.10' = 'Nao encerre nem compacte agora. O kit nao toca nesses processos, mas fechar o programa por fora interrompe o lote - e num lote noturno nao ha ninguem para notar.'
    'diagsessaoativa.11' = 'Os maiores dispensaveis (j = com janela aberta):'
    'diagsessaoativa.12' = '[{0}] {1,-30} {2,10}  ({3} processo(s))'
    'diagsessaoativa.13' = 'Suporte a Bluetooth ligado numa estacao cabeada'
    'diagsessaoativa.14' = 'O servico roda como SISTEMA e so o TI desliga. O Modulo 5 encerra os utilitarios de Bluetooth da sua sessao, mas o servico continua. Vale pedir a desativacao em massa nas estacoes de consultorio.'
    'diagsessaoativa.15' = 'Tarefas agendadas em execucao agora: {0}'
    'diagsessaoativa.16' = '- {0}{1}'
    'diagsessaoativa.17' = '{0} em {1} programa(s) dispensavel(is) na sessao de agora'
    'diagsessaoativa.18' = 'O Modulo 7 encerra tudo isso de uma vez, preserva Citrix e navegadores e devolve a memoria para o trabalho clinico. Volta ao normal no proximo logon.'
    'diagsessaoativa.19' = 'Sessao enxuta: {0} em dispensaveis'
    'modulo7sessao.01' = 'Modulo 5 - Otimizar a sessao atual'
    'modulo7sessao.02' = 'Nada e desinstalado. O que este modulo faz volta no proximo logon.'
    'modulo7sessao.03' = 'Citrix, navegador do prontuario, navegador web, Office e acesso remoto ficam sempre de fora.'
    'modulo7sessao.04' = 'Audio e microfone sao encerrados: o som continua funcionando, sai so a camada de realce.'
    'modulo7sessao.05' = 'ATENCAO: esta maquina hospeda servico ({0}).'
    'modulo7sessao.06' = 'Os servicos rodam como SISTEMA e nao entram na lista, mas confirme com quem depende deles'
    'modulo7sessao.07' = 'antes de otimizar uma estacao que serve outras pessoas.'
    'modulo7sessao.08' = 'Lendo processos da sessao...'
    'modulo7sessao.09' = 'Montando o plano...'
    'modulo7sessao.10' = 'Programas encerrados .......... {0}'
    'modulo7sessao.11' = 'Tarefas agendadas paradas ..... {0}'
    'modulo7sessao.12' = 'Prioridade de CPU ............. clinico acima do normal, resto abaixo'
    'modulo7sessao.13' = 'Memoria compactada ............ navegador e Office (Citrix intocado)'
    'modulo7sessao.14' = '{0} programa(s) com janela aberta ficaram desmarcados por nao serem reconhecidos:'
    'modulo7sessao.15' = 'Marque na lista se souber que pode fechar.'
    'restaurarsessao.01' = 'Restaurar a sessao'
    'restaurarsessao.02' = 'Devolve as prioridades ao normal e reabre o que foi encerrado no Modulo 5.'
    'restaurarsessao.03' = 'Prioridade normalizada em {0} programa(s).'
    'restaurarsessao.04' = 'Nenhum registro de encerramento nesta sessao. Nada a reabrir.'
    'restaurarsessao.05' = 'Registro do Modulo 5: {0} programa(s) encerrado(s).'
    'restaurarsessao.06' = '{0}: ja esta aberto.'
    'restaurarsessao.07' = '{0}: reaberto.'
    'restaurarsessao.08' = '{0}: nao encontrei o executavel para reabrir. Abra pelo menu Iniciar.'
    'restaurarsessao.09' = '{0} item(ns) voltam sozinhos quando o Windows precisar deles:'
    'restaurarsessao.10' = 'A camada de audio e os icones de bandeja voltam no proximo logon.'
    'restaurarsessao.11' = 'Sessao restaurada: {0} programa(s) reabertos, prioridades normalizadas.'
    'restaurarsessao.12' = 'O que foi tirado da inicializacao pelo Modulo 3 volta pelo botao "Desfazer otimizacoes".'
    'acaoprioridade.01' = 'Prioridade ajustada: {0} programa(s) clinico(s) acima do normal, {1} abaixo do normal.'
    'acaoprioridade.02' = 'O Windows zera as prioridades no proximo logon.'
    'acaomemoria.01' = 'Nenhum programa aceitou a compactacao nesta maquina.'
    'acaomemoria.02' = 'Memoria compactada em {0} programa(s): {1} devolvidos ao Windows.'
    'acaomemoria.03' = 'A primeira acao dentro de cada programa pode demorar um instante: ele recarrega o que precisa.'
    'topo.01' = 'Workstation Kit'
    'topo.02' = 'Workstation Kit  ·  diagnostico, limpeza e otimizacao de sessao'
    'topo.03' = 'Somente acoes sem administrador  ·  versao {0}  ·  {1}\{2}'
    'topo.04' = 'Pronto.'
    'topo.05' = 'Copiar log'
    'topo.06' = 'Salvar log'
    'topo.07' = 'Limpar tela'
    'topo.08' = 'Parar'
    'topo.09' = 'Diagnostico e inventario'
    'topo.10' = 'Arquivos e pastas grandes'
    'topo.11' = 'Otimizar a sessao de agora'
    'topo.12' = 'O que sobrevive ao logoff'
    'topo.13' = 'REVERTER E FECHAR O CICLO'
    'topo.14' = 'Restaurar sessao (reabrir)'
    'topo.15' = 'Ferramenta de apoio tecnico.
Nao validada para uso clinico.
Nenhuma acao exige administrador.
Tudo o que e desativado volta pelo Desfazer.'
    'topo.16' = 'O que vai ser aplicado'
    'topo.17' = 'Tudo o que e seguro ja vem marcado. Desmarque o que nao quiser e clique em APLICAR TUDO.'
    'topo.18' = 'Remover o que tem mais de:'
    'topo.19' = 'Marcar tudo'
    'topo.20' = 'Desmarcar tudo'
    'topo.21' = 'APLICAR TUDO'
    'topo.22' = 'Fechar'
    'topo.23' = 'AVALIAR = pode render espaco · NAO APAGAR = deixe como esta. Duplo clique abre no Explorer.'
    'topo.24' = 'Status da sessao'
    'topo.25' = 'Vale so para a sessao de agora. O proximo logon devolve a maquina ao estado normal.'
    'topo.26' = 'REVERTER PARA O ORIGINAL'
    'topo.27' = 'Atualizar'
    'comprotecao.01' = 'Erro inesperado em {0}: {1}'
    'topo.28' = 'Periodo alterado para: {0}. Recalculando...'
    'topo.29' = 'Parando apos a etapa atual...'
    'topo.30' = 'Log copiado. Cole no e-mail ou no chamado.'
    'topo.31' = 'Nao foi possivel copiar. Selecione o texto e use Ctrl+C.'
    'topo.32' = 'Log salvo em '
    'topo.33' = 'Nao foi possivel salvar o arquivo.'
    'topo.34' = 'Tela limpa.'
    'topo.35' = 'Kit de Suporte pronto.'
    'topo.36' = 'Maquina {0}  ·  usuario {1}  ·  {2:dd/MM/yyyy HH:mm}'
    'topo.37' = 'Modulo 2: inventario da maquina + diagnostico do que da para resolver aqui.'
    'topo.38' = 'Modulo 3: analisa, deixa tudo marcado e resolve em um clique ("Aplicar tudo").'
    'topo.39' = 'Modulo 4: os 10 maiores arquivos e as 10 maiores pastas do C:, com link para o Explorer.'
    'topo.40' = 'Modulo 5: deixa a sessao de agora leve, preservando Citrix, prontuario e navegadores.'
    'topo.41' = 'Citrix e Tasy sao testados dentro do Modulo 2.'
    'topo.42' = 'Este kit so faz o que roda sem administrador.'
    'topo.43' = 'Aplicativo autocontido: a rotina de preparacao vem dentro dele.'
}

$script:Textos['en'] = @{
    'item.15' = 'Game stores'
    'diagcitrix.18' = 'Testing '
    'topo.37' = 'Module 2: machine inventory + diagnosis of what can be resolved here.'
    'invsoftware.04' = 'Clinical software:'
    'diagtasy.23' = '{0}: ports open but the page does not respond'
    'diagencerraveis.03' = 'No dispensable program running now'
    'diagespaco.03' = 'Disk'
    'invredeimpressoras.12' = 'Printer'
    'item.05' = 'Xbox and Game Bar'
    'diaginicioacionave.05' = 'Startup took {0} seconds'
    'invecossistema.09' = '      {0}'
    'arquivosgrandes.16' = 'CAUTION-REVIEW .. {0} in {1} item(s) - check the contents before moving or deleting'
    'planolimpeza.04' = 'Saved credential for a server that does not respond: {0}'
    'obs.05' = 'An old Java cache makes Tasy fail when downloading a file. It is downloaded again on the next visit.'
    'invredeimpressoras.07' = 'The adapter and the switch are 1 Gbps: 100 Mbps is nearly always an old cable, badly crimped, or a bad port. Replacing the cable costs nothing and improves Tasy, Citrix and the opening of the 40 MB portals.'
    'invidentificacao.07' = 'Windows since .: {0:dd/MM/yyyy}'
    'modulo7sessao.02' = 'Nothing is uninstalled. What this module does comes back at the next sign-in.'
    'diagpersistencia.28' = '. The disk outside the profile remained. This is a discarded profile, and only IT fixes it.'
    'diagsessaoativa.11' = 'The largest dispensable ones (j = with an open window):'
    'modulo3limpeza.01' = 'Run Module 3 first to analyse the machine.'
    'modulo3analise.02' = 'Nothing is changed in this step. Everything that is safe comes already selected.'
    'arquivosgrandes.07' = 'The paging file lives on that disk: if it fills up, the whole machine freezes. Free up space on it too.'
    'topo.38' = 'Module 3: analyses, leaves everything selected and resolves it in one click ("Apply all").'
    'diagcitrix.24' = 'The workstation is not seeing this server''s internal DNS. Check the cable/VPN and the adapter''s DNS suffix before opening a ticket.'
    'topo.25' = 'Applies only to the current session. The next sign-in returns the machine to its normal state.'
    'topo.18' = 'Remove anything over:'
    'otimizacoes.01' = 'Undo optimisations'
    'idioma.01' = 'Language'
    'diagtasy.57' = 'Ask IT to exclude the browser and Java cache folders from the antivirus scan: it is a direct gain when browsing Tasy.'
    'modulo3limpeza.42' = 'Result'
    'topo.03' = 'Only non-administrator actions  ·  version {0}  ·  {1}\{2}'
    'diagnuvem.04' = 'OneDrive syncing {0} on the local disk'
    'invmigracao.15' = 'Windows tries to use each one when opening Explorer and waits for the timeout. Module 3 removes them - the item comes UNSELECTED because there is no undo.'
    'invidentificacao.04' = 'Equipment .....: {0} {1}   serial {2}'
    'diagcitrix.30' = '{0}: server {1} not responding on the web ports'
    'diagpersistencia.21' = '   SURVIVED  {0,-30} holds {1}'
    'restaurarsessao.07' = '{0}: reopened.'
    'diagtasy.20' = '{0}: server {1} not responding'
    'invmigracao.06' = 'No migration agent installed.'
    'item.11' = 'Apple services'
    'invsoftware.02' = '{0} programs installed'
    'diagsessaoativa.14' = 'The service runs as SYSTEM and only IT turns it off. Module 5 closes the Bluetooth utilities in your session, but the service keeps running. It is worth asking for bulk deactivation on the consulting-room workstations.'
    'invmigracao.07' = 'Signed-in user .....: {0}'
    'item.03' = 'Automatic updaters'
    'diaglixeira.02' = 'Recycle Bin with {0}'
    'diagcitrix.21' = '{0}: the short name {1} does not resolve, only the full one {2}'
    'item.26' = 'Games and overlays'
    'invecossistema.02' = 'No record of the ecosystem on this workstation. This is the normal state for someone who has installed nothing.'
    'nota.05' = 'Closed only if there is no window open.'
    'diagpersistencia.15' = 'The disk reverts to its previous state at every restart. Nothing installed or adjusted persists. Only IT can turn this off - take this log to the support ticket.'
    'diagdownloads.09' = '{0,10}  {1}  ({2:dd/MM/yyyy})'
    'modulo3limpeza.22' = 'Interrupted by the user.'
    'diagcitrix.48' = 'The browser shows a not-secure site warning and may block downloads and printing. Ask IT to install the internal authority''s certificate.'
    'obs.01' = 'Files in use are skipped automatically.'
    'modulo3limpeza.39' = 'Will no longer open by itself: {0}'
    'diagcitrix.57' = 'These are just launchers downloaded by the browser. Module 3 removes those in Downloads within the chosen period; the Citrix cache is left intact.'
    'diagtasy.56' = 'One environment fast and another slow -> the problem is with that server, not yours.'
    'topo.23' = 'ASSESS = may yield space · DO NOT DELETE = leave as it is. Double click opens in Explorer.'
    'inventariodiagnost.12' = '{0,-16} {1,-26} {2} finding(s), {3} high impact'
    'diagpersistencia.07' = 'Temporary profile: this entire session is discarded at sign-out'
    'topo.41' = 'Citrix and Tasy are tested inside Module 2.'
    'otimizacoes.05' = 'Undo all'
    'invecossistema.05' = 'While this lasts no application can register a version, and support has no way of knowing what is installed. Take this log to whoever looks after the installer.'
    'diagcitrix.33' = 'Above 3 seconds the application list is slow to appear. Clearing the Workspace cache in Module 3 usually resolves it on the workstation side.'
    'diagtasy.08' = 'Java running: {0} · {1} · {2}'
    'alvo.07' = 'Teams cache (new version)'
    'topo.05' = 'Copy log'
    'explorer.03' = 'Taskbar changes appear at the next sign-in.'
    'modulo3limpeza.57' = 'Cleanup completed.'
    'modulo3limpeza.58' = 'Closing {0}...'
    'modulo3limpeza.59' = '{0}: could not be closed (protected or already closed)'
    'modulo3limpeza.60' = 'Task stopped: {0}{1}'
    'modulo3limpeza.61' = '{0}: no permission to stop it (runs as SYSTEM - IT only)'
    'invpastaclinica.08' = 'Incomplete preparation: {0} missing'
    'diagcitrix.06' = 'Citrix Workspace not found on this machine'
    'openitemgrande.03' = 'No longer exists: {0}'
    'modulo3limpeza.18' = 'Closing dispensable programs in the session...'
    'restartcomputador.04' = 'Restart scheduled for 10 seconds. Use shutdown /a to cancel.'
    'invecossistema.14' = 'There is entrega.txt on the disk but it does not appear in the registry: '
    'topo.39' = 'Module 4: the 10 largest files and the 10 largest folders on C:, with a link to Explorer.'
    'item.17' = 'Skype'
    'diagespaco.01' = 'Disk space'
    'invsoftware.06' = '[ ] {0}'
    'item.07' = 'Phone Link'
    'diagtasy.39' = '{0,-34} {1,10} {2,10}  {3}'
    'diagtasy.37' = 'Zone ......: outside the Windows Intranet - no effect here, Tasy has its own sign-in.'
    'item.22' = 'Automatic updater'
    'diagtasy.28' = 'Fluctuation: one of the samples took {0} ms against {1} ms for the others. Unstable network, not slow.'
    'modulo3analise.11' = 'Deselect what you do not want and click "Apply all".'
    'invidentificacao.03' = 'System ........: {0} build {1}'
    'diagpersistencia.26' = 'The profile does not keep changes: {0} of {1} places lost the marker'
    'diagcaches.02' = '{0} recoverable only with safe cleanup'
    'diagsessaoativa.04' = '{0,-46} {1,4} process(es)  {2,10}'
    'invecossistema.10' = '{0} registry entry(ies) point to a folder that does not exist'
    'diagsessaoativa.13' = 'Bluetooth support enabled on a wired workstation'
    'modulo7sessao.05' = 'ATTENTION: this machine hosts a service ({0}).'
    'topo.27' = 'Refresh'
    'inventariodiagnost.16' = 'Everything above is solved by Module 3 (Cleanup). Open it and click "Apply all".'
    'invmigracao.02' = 'Migration agent installed: {0} {1}'
    'restartcomputador.01' = 'Restart the computer'
    'otimizacoes.03' = 'Undo'
    'nota.15' = 'It raises the clinical ones above normal and lowers the rest. Resets at the next sign-in.'
    'diagnuvem.06' = 'Cloud'
    'topo.01' = 'Workstation Kit'
    'modulo3limpeza.29' = 'Recycle Bin: {0}'
    'invredeimpressoras.08' = 'DNS suffix: {0}'
    'diagcitrix.52' = 'Citrix traffic is passing through the proxy unnecessarily, which makes it slow to open. Ask IT to add it to the exceptions.'
    'arquivosgrandes.05' = 'Windows paging sits on {0} ({1}% free). As long as that disk has space, it does not burden performance.'
    'diagajustespendent.05' = '{0} performance adjustment(s) not yet applied'
    'modulo7sessao.09' = 'Building the plan...'
    'item.44' = 'Google Drive'
    'restaurarsessao.12' = 'What Module 3 removed from startup comes back with the "Undo optimisations" button.'
    'restaurarsessao.01' = 'Restore the session'
    'modulo3limpeza.08' = 'Cleanup cancelled by the user.'
    'diaginicioacionave.21' = 'The time does not come from your user''s programs. There are {0} network drive(s) and {1} network printer(s) reconnected at every sign-in, plus the corporate agents. Take these numbers to IT.'
    'item.13' = 'Personal cloud'
    'diagpersistencia.02' = 'Profile ...: {0}'
    'arquivosgrandes.19' = 'Select any row in the right-hand pane and click "Open in Explorer".'
    'diagcitrix.36' = '{0}: portal responded HTTP {1}'
    'item.24' = 'Manufacturer assistant'
    'alvo.02' = 'User temporary files'
    'diagtasy.47' = 'CentBrowser cache: {0}'
    'diagpersistencia.17' = 'Sign out (or restart) and run this module again: it will say exactly what survived.'
    'modulo7sessao.10' = 'Programs closed ............... {0}'
    'diagtasy.04' = 'Browser {0,-14} {1,10} in {2} process(es)'
    'restaurarsessao.08' = '{0}: could not find the executable to reopen. Open it from the Start menu.'
    'invredeimpressoras.10' = 'Default printer offline: {0}'
    'diagpersistencia.09' = 'Policy deletes the local copy of the profile at sign-out'
    'diagtasy.14' = 'internal network server'
    'invofficeexcel.05' = 'Run Module 1 (Prepare environment) on this machine.'
    'logemuso.02' = '{0}: {1} protects {2} process(es), {3}.'
    'arquivosgrandes.10' = 'THE 10 LARGEST FILES'
    'diagcitrix.31' = 'Until this comes back, the applications on that target do not open. Check the workstation''s network and, if the other workstations fail as well, notify IT.'
    'inventariodiagnost.09' = 'If the slowness continues, the bottleneck requires an administrator (disk, antivirus, network, hardware).'
    'otimizacoes.07' = 'Restored: {0}'
    'inventariodiagnost.02' = 'Machine {0} · user {1} · {2:dd/MM/yyyy HH:mm:ss}'
    'explorer.01' = 'Explorer restarted: the taskbar already looks clean.'
    'modulo3limpeza.55' = 'Restart scheduled for 15 seconds. Use shutdown /a in Run to cancel.'
    'alvo.05' = 'Firefox cache'
    'openitemgrande.01' = 'Select a row from the list.'
    'diagajustespendent.01' = 'Pending performance adjustments'
    'alvo.19' = 'Installers downloaded by Edge/Chrome'
    'planolimpeza.06' = 'Credentials'
    'topo.15' = 'Technical support tool.
Not validated for clinical use.
No action requires an administrator.
Everything disabled comes back with Undo.'
    'diagnuvem.14' = 'Old Teams versions taking up {0}'
    'diagsessaoativa.19' = 'Lean session: {0} in dispensables'
    'topo.13' = 'REVERT AND CLOSE THE CYCLE'
    'modulo3analise.04' = 'Disk space freed .............. {0}'
    'invsoftware.08' = 'This machine hosts a service (database or application)'
    'item.09' = 'Search and Copilot'
    'item.42' = 'Give CPU priority to Citrix, the medical record and Office'
    'diagtasy.35' = '{0}: unstable response (varied from {1} to {2} ms)'
    'modulo7sessao.13' = 'Memory compacted .............. browser and Office (Citrix untouched)'
    'diagtasy.38' = 'COMPARISON BETWEEN THE ENVIRONMENTS'
    'diaginicioacionave.09' = 'Items of your user account that open together with Windows: {0}'
    'diagdownloads.01' = 'Downloads folder'
    'invpastaclinica.07' = '{0,-40} {1,10}  {2:dd/MM/yyyy HH:mm}'
    'diagtasy.43' = 'This rules out the workstation and the local network: the problem is the slow server.'
    'diagsessaoativa.07' = 'Ecosystem load running (protected from Modules 3 and 5):'
    'invidentificacao.01' = 'Inventory · identification'
    'diagcitrix.09' = 'There is a published session open now (ARIA, MOSAIQ or Monaco)'
    'topo.30' = 'Log copied. Paste it into the e-mail or the ticket.'
    'diaginicioacionave.14' = 'What weighs most on this workstation''s sign-in:'
    'diagpersistencia.24' = 'Not even the disk outside the profile survived the sign-out'
    'diagencerraveis.08' = 'Module 3 closes them all at once. They come back when you open them.'
    'item.18' = 'Zoom'
    'diagsessaoativa.09' = 'radioterapia.ai work in progress {0}: {1} using {2}'
    'diagcitrix.15' = '--- {0} ---'
    'inventariodiagnost.14' = 'RESOLVE IN THIS ORDER:'
    'diagcitrix.28' = 'responds'
    'invofficeexcel.01' = 'Inventory · Office and Excel'
    'diagespaco.08' = 'C:\Windows\Temp: {0} (system, out of the user''s reach)'
    'diagencerraveis.06' = '{0,-34} {1,10}   ({2} process(es))'
    'diagtasy.27' = 'Processing ....: ~{0} ms on the server (HTTP minus network)'
    'diaginicioacionave.02' = 'The machine was on for {0:N1} h before this sign-in.'
    'modulo3limpeza.27' = '{0}: nothing to clean'
    'invidentificacao.08' = 'Total memory of {0}'
    'diagtasy.46' = 'An old cache makes Tasy open the wrong version and fail to download a file. Module 3 clears it.'
    'diagencerraveis.04' = 'Memory'
    'diagtasy.50' = 'Proxy without an exception for internal Tasy: {0}'
    'modulo3limpeza.21' = 'Compacting the memory of the programs that stay...'
    'diagmemoria.01' = 'Memory (RAM)'
    'inventariodiagnost.04' = 'Diagnostics interrupted by the user.'
    'diaginicioacionave.23' = 'Each of them competes for disk and CPU at sign-in. Module 3 disables them all at once and the "Undo" button reverses it.'
    'diaginicioacionave.04' = 'To measure: restart and run Module 2 right after signing in.'
    'ajusteregistro.01' = '{0}: blocked by IT policy (Windows does not let the user change it).'
    'item.39' = 'AppData\Roaming in the profile'
    'diagdownloads.11' = 'Use the "Empty Downloads" button in the side menu. It goes to the Recycle Bin, so it can be recovered.'
    'diagcitrix.34' = '{0}: portal responding (HTTP {1}, {2} ms)'
    'diagsistemaacionav.03' = 'Restart the computer at the end of the cleanup. Using "Shut down" does not solve it: only "Restart" clears the memory.'
    'diaginicioacionave.20' = 'Long sign-in with only {0} user startup item(s)'
    'modulo3limpeza.49' = 'Space freed ................... {0}'
    'modulo3limpeza.52' = 'Cleanup applied.

The gain in startup and in the settings only appears after a restart.

Restart now? Save your files first.'
    'titulo.01' = '== {0} {1}'
    'modulo3limpeza.35' = '   ... and {0} others.'
    'invredeimpressoras.01' = 'Inventory · network and printers'
    'diagcitrix.38' = '{0}: portal did not respond'
    'restaurarsessao.02' = 'Returns priorities to normal and reopens what was closed in Module 5.'
    'item.28' = 'Edge opening by itself at sign-in'
    'item.40' = 'Disk outside the profile'
    'diagtasy.06' = 'Tasy runs inside the browser: with this load the screen freezes even with the server responding quickly. Close the tabs you do not use before opening a ticket.'
    'diagtasy.51' = 'The medical record traffic goes through the proxy unnecessarily and that adds time on every screen. Ask IT to include it in the exceptions.'
    'diagmapeamentos.05' = 'A dead mapping freezes Explorer and Office''s "Save as" for 30 seconds or more. Module 3 removes it.'
    'invecossistema.12' = '. A record that lies is worse than an empty record: empty makes you search, wrong makes you conclude. Take it to whoever is responsible for the installer - the kit does not fix another project''s record.'
    'diagtasy.16' = 'DNS .......: {0} ({1} ms)'
    'invofficeexcel.07' = 'Add-in: {0} (LoadBehavior {1})'
    'diagmemoria.11' = 'Close and open Teams once a day, or clear the cache in Module 3.'
    'modulo3analise.01' = 'Module 3 - Cleanup (analysis)'
    'diagmemoria.08' = '{0,-24} {1,10}   ({2} process(es))'
    'inventariodiagnost.07' = 'Map of critical points'
    'otimizacoes.09' = 'Startup for {0} restored.'
    'diagtasy.41' = 'status'
    'nota.08' = 'Keeps the disk from filling up again.'
    'diagpersistencia.18' = 'Marker from {0}, from the same sign-in session. There has been no sign-out yet to compare.'
    'modulo3limpeza.32' = 'Downloads: {0} items to the Recycle Bin, {1} freed'
    'achado.01' = '-> '
    'diagajustespendent.03' = 'Adjustments'
    'diagcitrix.58' = 'If one destination fails and the others work, the problem is with that server, not with the workstation.'
    'modulo3limpeza.25' = '{0}: was no longer running'
    'modulo3limpeza.14' = 'Emptying the Recycle Bin...'
    'invinicializacaoco.04' = '{0} startup items belong to the machine, not to your user'
    'topo.17' = 'Everything safe is already selected. Deselect what you do not want and click APPLY ALL.'
    'ajustezonascitrix.01' = '{0} added to the Intranet sites.'
    'diagpersistencia.20' = 'Previous marker from {0}. There has been a sign-out since then - this is the result:'
    'modulo3limpeza.11' = 'Module 3 - Applying'
    'diagcitrix.50' = 'Exceptions .: {0}'
    'alvo.10' = 'Remote Desktop image cache'
    'modulo7sessao.11' = 'Scheduled tasks stopped ....... {0}'
    'topo.08' = 'Stop'
    'invsoftware.11' = 'Corporate agents running: {0}'
    'invpastaclinica.05' = 'Total: {0}'
    'invidentificacao.11' = '{0} -> {1}'
    'diagencerraveis.01' = 'Programs that can be closed safely now'
    'restaurarsessao.09' = '{0} item(s) come back by themselves when Windows needs them:'
    'logemuso.03' = '   None of them will be closed, lowered or compacted. Port {0}.'
    'inventario.06' = 'Saved in the network LOGs folder. The text report is this log itself.'
    'diagnuvem.15' = 'Module 3 removes it with the item "Old Teams versions".'
    'acaoprioridade.02' = 'Windows resets the priorities at the next sign-in.'
    'diagtasy.02' = 'Installed: {0} {1}'
    'arquivosgrandes.02' = 'Read-only. The kit deletes nothing here: you decide item by item.'
    'diagcitrix.17' = 'Address ...: {0}'
    'topo.34' = 'Screen cleared.'
    'diagcaches.04' = '{0} recoverable with safe cleanup'
    'inventariodiagnost.05' = 'Diagnostic {0}/{1}: {2}'
    'invofficeexcel.03' = 'Trusted location: {0}'
    'diagespaco.06' = 'C:\Windows\Temp with {0}'
    'diaginicioacionave.19' = 'Windows tries to reconnect each of these at sign-in and waits for the timeout. It is the most likely explanation for the long sign-in. Ask IT to remove the mapping or to grant access.'
    'modulo3analise.10' = 'Select it in the list if you know it can be disabled.'
    'diagtasy.25' = 'Network (TCP) .: min {0} · typical {1} · max {2} ms'
    'diagnuvem.02' = 'OneDrive running · {0}'
    'diagcitrix.41' = '{0}: {1} is not in the Intranet zone'
    'obs.08' = 'Installation folder used by IT. Files in use are skipped.'
    'invredeimpressoras.04' = 'The portals are 40 MB and are opened from the network. A cable noticeably reduces the opening time and prevents file corruption.'
    'item.10' = 'Manufacturer assistants'
    'inventario.03' = 'The spreadsheet was not written. The report is this log ("Copy log" or "Save log").'
    'modulo3limpeza.17' = 'Applying performance adjustments...'
    'inventariodiagnost.10' = 'Copy this log and open a ticket: it already shows that the user side is clean.'
    'invinicializacaoco.02' = '{0} items (user + machine)'
    'alvo.13' = 'Windows internet cache'
    'otimizacoes.11' = 'Everything restored. Sign out to apply it completely.'
    'diagcitrix.12' = 'Stores already configured in this user''s Workspace:'
    'modulo3limpeza.15' = 'Sending old Downloads to the Recycle Bin...'
    'item.37' = 'User registry (HKCU)'
    'diagtasy.54' = 'Network high and page high    -> workstation''s network path. Cable instead of Wi-Fi, and a network ticket.'
    'modulo3limpeza.07' = 'Question'
    'acaomemoria.01' = 'No program accepted compaction on this machine.'
    'alvo.25' = 'Empty the Recycle Bin'
    'diagnuvem.09' = 'OneDrive is not running.'
    'invredeimpressoras.03' = 'Workstation connected over Wi-Fi ({0})'
    'diagcitrix.46' = 'Security ..: address only over HTTP. The traffic travels unencrypted.'
    'modulo3limpeza.37' = 'RECORD: this user''s Desktop or Documents are inside OneDrive.'
    'diagtasy.40' = 'page'
    'restaurarsessao.04' = 'No record of closures in this session. Nothing to reopen.'
    'modulo3limpeza.05' = 'Apply cleanup'
    'diaginicioacionave.24' = '{0} startup program(s) can be disabled'
    'topo.32' = 'Log saved to '
    'invmigracao.04' = 'The migration finished years ago. Removing it requires administrator: ask IT for the bulk uninstall. It runs a service and scans the profile at every sign-in.'
    'diagtasy.18' = 'The workstation is not seeing the DNS. It is a network problem on the workstation, not on the medical record system.'
    'modulo3analise.03' = 'What will happen when you click "Apply all":'
    'diaglixeira.01' = 'Recycle Bin'
    'diagcitrix.22' = 'Change the address of the shortcut and of the configuration to {0} . Without this the browser does not find the portal on this workstation.'
    'item.29' = 'protected (security, network or clinical)'
    'diagnuvem.10' = 'Teams running · {0} in {1} process(es)'
    'nota.06' = 'Makes a slow machine respond immediately.'
    'diagdownloads.04' = '{0} files · {1} in total'
    'modulo3limpeza.45' = 'Total reclaimed ............... {0}'
    'invmigracao.11' = 'Profile with domain suffix: {0}'
    'diagmemoria.10' = 'Teams using a lot of memory'
    'alvo.22' = 'Office document cache'
    'topo.28' = 'Period changed to: {0}. Recalculating...'
    'openitemgrande.05' = 'Could not open: {0}'
    'obs.12' = 'The thumbnails are recreated by themselves.'
    'modulo3limpeza.20' = 'Adjusting CPU priority...'
    'acaomemoria.02' = 'Memory compacted in {0} program(s): {1} returned to Windows.'
    'invinicializacaoco.06' = 'Startup'
    'diagpersistencia.04' = 'Mandatory profile: Windows discards your changes at every sign-out'
    'arquivosgrandes.11' = 'No file above {0} MB.'
    'invservicos.03' = '{0,-40} {1,-12} {2}'
    'diagencerraveis.10' = 'Work programs, clinical programs, security programs and browsers are never closed by the kit.'
    'item.33' = 'Turn on Windows automatic cleanup'
    'invecossistema.06' = 'Ecosystem'
    'invservicos.01' = 'Inventory · running services'
    'diagpersistencia.11' = 'Roaming profile: the copy on the server is the one that counts'
    'modulo3limpeza.28' = '{0}: {1} freed{2}'
    'topo.06' = 'Save log'
    'modulo3limpeza.10' = 'Module 5 - Applying to the session'
    'topo.07' = 'Clear screen'
    'inventario.04' = 'Log folder unavailable: {0}'
    'diagcitrix.05' = 'Citrix Workspace installed (version {0})'
    'alvo.16' = 'Video cache (shaders)'
    'topo.36' = 'Machine {0}  ·  user {1}  ·  {2:dd/MM/yyyy HH:mm}'
    'item.14' = 'Messaging apps'
    'diagcitrix.19' = '...'
    'diagespaco.04' = 'Run Module 3 to free up space.'
    'tarefasemexecucao.01' = 'Reading scheduled tasks (a few seconds)...'
    'item.04' = 'Adobe Creative Cloud (background)'
    'invredeimpressoras.11' = 'Excel and Word freeze when opening a file if the default printer does not respond. Change the default to "Microsoft Print to PDF" until they resolve it.'
    'modulo3limpeza.40' = 'Could not disable {0}'
    'diaginicioacionave.06' = 'See below what is weighing it down: startup items, network drives and network printers. Module 3 handles the user''s part.'
    'diagcitrix.43' = 'Security ..: the server redirects to HTTPS on its own. The shortcut can stay as it is.'
    'invredeimpressoras.06' = 'Network adapter negotiating only {0}'
    'comprotecao.01' = 'Unexpected error in {0}: {1}'
    'obs.06' = 'UNCHECKED on purpose: this is the medical-record browser. It does not delete bookmarks or passwords, only the cache. Check it if Tasy shows a blank screen or a file error.'
    'topo.40' = 'Module 5: makes the current session light, preserving Citrix, the medical record system and browsers.'
    'diagpersistencia.29' = 'Marker renewed in {0} of {1} places.'
    'modulo7sessao.03' = 'Citrix, the medical record browser, the web browser, Office and remote access are always left out.'
    'diagcitrix.40' = 'Zone ......: already in the Intranet (single sign-on works)'
    'alvo.11' = 'Opera and Vivaldi cache'
    'diagcaches.01' = 'Recoverable space (cleanup preview)'
    'modulo3limpeza.02' = 'Cleanup'
    'diagcitrix.27' = 'ICA 1494 ..: {0}   ·   Reliability 2598: {1}{2}'
    'item.35' = 'Remove the OneDrive icon from Explorer'
    'nota.03' = 'They come back by themselves when there is an update.'
    'nota.16' = 'Returns to Windows the memory the browser and Office reserved without using. The published Citrix session is left alone, so it does not stutter.'
    'diagsessaoativa.17' = '{0} in {1} dispensable program(s) in the current session'
    'restaurarsessao.10' = 'The audio layer and the tray icons come back at the next sign-in.'
    'diagmemoria.07' = 'Programs using the most memory:'
    'otimizacoes.10' = 'Zone removed: {0}'
    'invmigracao.14' = '{0} saved credential(s) pointing to a non-existent server'
    'diagcitrix.51' = 'Proxy with no exception for: {0}'
    'arquivosgrandes.08' = 'Searching for large files...'
    'topo.19' = 'Select all'
    'diagtasy.31' = 'Network at {0} ms. The slowness is on the server, not on the workstation.'
    'diagajustespendent.06' = 'Module 3 applies them all. Reversible with the "Undo optimisations" button.'
    'restaurarsessao.06' = '{0}: is already open.'
    'nota.10' = 'It disappears from the bar and stops using memory. Reversible.'
    'invsoftware.01' = 'Inventory · installed software'
    'item.02' = 'Microsoft Teams'
    'diaglixeira.03' = 'Emptying the Recycle Bin is a Module 3 routine. Check first that there is nothing to recover.'
    'invpastaclinica.02' = '{0} does not exist'
    'item.27' = 'Optional utility'
    'nota.13' = 'Done before Downloads, so that what leaves there can still be recovered.'
    'modulo3limpeza.44' = 'Memory compacted .............. {0}'
    'nota.04' = 'Windows loads it again by itself when it needs to.'
    'invmigracao.13' = '{0} saved credential(s) for servers that no longer exist:'
    'planolimpeza.03' = 'To remove it, run in Command Prompt:  net use {0} /delete'
    'invmigracao.08' = 'Profile folder .....: {0}'
    'diagsessaoativa.01' = 'Current session: what is active now'
    'diagcitrix.54' = 'Loose .ica files: {0}'
    'obs.02' = 'It does not delete bookmarks, passwords or history.'
    'diagtasy.09' = 'Tasy Java client running in 32-bit'
    'topo.29' = 'Stopping after the current step...'
    'diaginicioacionave.27' = 'Only IT changes the machine items (for all users) and the security ones.'
    'arquivosgrandes.09' = 'Large files in '
    'inventariodiagnost.13' = '{0} critical problem(s) and {1} warning(s).'
    'diagpersistencia.06' = 'Persistence'
    'topo.42' = 'This kit only does what runs without administrator.'
    'invofficeexcel.08' = 'Add-in {0} loads together with Excel'
    'item.31' = 'Turn off animations, shadows and transparency'
    'diagtasy.52' = 'HOW TO READ THE RESULT'
    'modulo7sessao.15' = 'Select it in the list if you know it can be closed.'
    'invecossistema.04' = 'Ecosystem registry unreadable: {0}'
    'diagpersistencia.13' = 'Local profile: the user''s registry should survive sign-out'
    'alvo.08' = 'Java Web Start cache (Tasy)'
    'restaurarsessao.11' = 'Session restored: {0} program(s) reopened, priorities normalised.'
    'diagcitrix.07' = 'Without the client installed, ARIA/MOSAIQ/Monaco will not open. Installation requires IT.'
    'openitemgrande.02' = 'Open in Explorer'
    'alvo.14' = 'Windows error reports'
    'topo.02' = 'Workstation Kit  ·  diagnostics, cleanup and session optimisation'
    'item.16' = 'OneDrive'
    'invsoftware.05' = '[X] {0,-14} {1}'
    'item.06' = 'NVIDIA overlay'
    'item.41' = 'Empty the Recycle Bin (routine)'
    'diagespaco.02' = 'A nearly full disk is the number one cause of slowness. Module 3 frees space now.'
    'diagtasy.36' = 'Oscillation of this size is what makes the Tasy screen "freeze" from time to time. Record it in the ticket with the time.'
    'item.23' = 'Apple component'
    'diagtasy.29' = '{0}: page taking {1} ms to respond'
    'diagcitrix.10' = 'The kit never closes Citrix. Close the session before clearing the Workspace cache.'
    'planolimpeza.05' = 'To remove, run:  cmdkey /delete:{0}   -- check the name first: if it is a DNS error, the server exists and the credential is valid.'
    'arquivosgrandes.03' = 'Only the C: disk. It is the only one whose free space affects Windows performance.'
    'diagdownloads.05' = '{0} files more than 90 days old · {1}'
    'invidentificacao.02' = 'Machine .......: {0}   user: {1}\{2}'
    'obs.04' = 'Close Teams first. Conversations stay on the server.'
    'diagtasy.48' = 'CentBrowser cache with {0}'
    'diagpersistencia.27' = 'Lost: '
    'modulo7sessao.01' = 'Module 5 - Optimise the current session'
    'diagnuvem.11' = 'Teams using {0}'
    'diagcaches.03' = 'Run Module 3 (Safe cleanup).'
    'obs.14' = 'Close Excel first. These are lock files left over from sessions that hung.'
    'modulo3analise.05' = 'Memory returned now ........... {0}'
    'topo.24' = 'Session status'
    'inventariodiagnost.17' = 'What requires administrator is marked as such: use "Copy log" for the ticket.'
    'diagcitrix.25' = 'Web ports  : {0}'
    'logemuso.04' = '   The file date cannot be read. If the application is not open, delete:'
    'invmigracao.01' = 'Inventory · domain migration remnants'
    'nota.14' = 'It goes to the Recycle Bin, so it can be recovered. Change the period in the box above the list.'
    'diagsessaoativa.18' = 'Module 7 closes all of this at once, preserves Citrix and browsers and gives the memory back to clinical work. It returns to normal at the next sign-in.'
    'nota.01' = 'Comes back when you open OneDrive, or at the next sign-in.'
    'modulo3limpeza.43' = 'Memory returned ............... {0}'
    'topo.11' = 'Optimise the current session'
    'diagmemoria.05' = 'Below 8 GB, Excel + browser + Teams freeze the computer. Request an increase.'
    'invmigracao.16' = 'No orphaned credential in Credential Manager.'
    'invidentificacao.05' = 'Processor .....: {0} ({1} cores)'
    'arquivosgrandes.06' = 'Paging disk {0} with only {1}% free'
    'modulo3limpeza.13' = 'Cleaning caches and temporary files...'
    'diagencerraveis.09' = '{0} recoverable by closing dispensable programs'
    'diagcitrix.53' = '{0,10}  {1}'
    'invsoftware.10' = 'Server'
    'restaurarsessao.03' = 'Priority normalised in {0} program(s).'
    'restartcomputador.06' = 'Restart from the Start menu. Use "Restart", not "Shut down".'
    'modulo3limpeza.09' = 'SESSAO'
    'topo.04' = 'Done.'
    'diagsessaoativa.06' = '   session {0,-4} {1,-16} PID {2,-7} {3,10}'
    'item.12' = 'Spotify'
    'invsoftware.03' = '{0,-56} {1,-18} {2}'
    'diagcitrix.37' = 'The server is up but the store is not. Record the code in the ticket.'
    'item.25' = 'Windows extra'
    'topo.09' = 'Diagnostics and inventory'
    'diagtasy.32' = '{0}: responding in {1} ms'
    'nota.11' = 'Hides the shortcut only. Your files and your account stay intact.'
    'item.45' = 'LibreOffice'
    'alvo.03' = 'Microsoft Edge cache'
    'diagtasy.44' = 'Java Web Start cache: {0}'
    'diagpersistencia.16' = 'No previous marker. Leaving one now in each place.'
    'diagcitrix.02' = 'Old Citrix Receiver: version {0}'
    'modulo3limpeza.38' = 'With OneDrive out of startup, files only upload to the cloud when it is opened.'
    'diagcitrix.49' = 'Proxy enabled: {0}'
    'diagencerraveis.07' = '{0} of memory held in {1} dispensable program(s)'
    'diagencerraveis.02' = 'Criterion: they do not hold an open document and they come back on their own when you open them again.'
    'diagcitrix.56' = '{0} stray .ica files'
    'modulo3limpeza.46' = 'Memory in use now ............. {0}% ({1} free)'
    'inventariodiagnost.11' = 'Where the problem is, by area:'
    'diagpersistencia.08' = 'Windows could not open your profile and created a disposable one. There is no point adjusting anything now. Ticket for IT to recreate the profile.'
    'topo.20' = 'Deselect all'
    'topo.22' = 'Close'
    'invidentificacao.10' = 'Hardware'
    'logemuso.01' = 'App-in-use flag found but OLD ({0:N0} min). Ignored.'
    'ajusteregistro.03' = '{0} of {1} settings in this group are locked by policy. The rest were applied.'
    'diagcitrix.32' = '{0}: portal responds in {1} ms'
    'otimizacoes.04' = 'This returns the visual effects, the background applications and the automatic startup to their previous state.

Continue?'
    'modulo3limpeza.56' = 'Error during cleanup: {0}'
    'alvo.06' = 'Teams cache (classic version)'
    'modulo3limpeza.31' = 'Recycle Bin emptied: {0} freed'
    'inventario.08' = 'Could not write the report: {0}'
    'diaginicioacionave.26' = 'User startup is already lean'
    'modulo7sessao.12' = 'CPU priority .................. clinical above normal, rest below'
    'modulo3limpeza.19' = 'Stopping running scheduled tasks...'
    'invsoftware.07' = '[X] Server role: {0} ({1} process(es), {2} service(s))'
    'diagsistemaacionav.04' = 'Restart'
    'obs.10' = 'Only the launchers the browser downloaded. The Citrix cache is never touched.'
    'diaginicioacionave.13' = '[   keep   ] {0,-30} {1}'
    'item.08' = 'Windows widgets'
    'item.43' = 'Compact the memory of programs that stay open'
    'diagtasy.34' = 'Above 60 ms inside the internal network is a bad path. Check the adapter''s negotiated speed and the switch port before blaming the server.'
    'item.21' = 'Adobe component'
    'invecossistema.13' = '{0} application(s) installed and missing from the registry'
    'inventariodiagnost.03' = 'Part 1: portrait of the machine. Part 2: what Module 3 resolves without administrator.'
    'invmigracao.05' = 'Migration'
    'diagmemoria.03' = 'Close what is not in use, mainly browser tabs.'
    'topo.21' = 'APPLY ALL'
    'arquivosgrandes.01' = 'Module 4 - Large files and folders'
    'diagpersistencia.14' = 'Disk write filter: {0}'
    'diagpersistencia.25' = 'This is a write filter or a disk freezer, not a profile problem. Only IT can turn it off.'
    'invmigracao.10' = 'Sign of a profile inherited from another account, typical of a domain migration. It works, but it confuses scripts, group policy and permissions. If this machine has a strange profile problem, this is where you start.'
    'inventariodiagnost.01' = 'Module 2 - Inventory and diagnostics'
    'diagmemoria.02' = 'Close programs and browser tabs. The computer is using disk as memory.'
    'diagtasy.11' = 'No Tasy address configured'
    'diagpersistencia.01' = 'Persistence - what survives sign-out'
    'topo.26' = 'REVERT TO ORIGINAL'
    'inventariodiagnost.15' = '{0}. [{1}/{2}] {3}'
    'invmigracao.03' = 'Domain migration agent still installed ({0})'
    'diagpersistencia.30' = 'Some location did not even accept a write just now - which is already an answer.'
    'otimizacoes.02' = 'No optimisation recorded to undo.'
    'diagtasy.55' = 'Everything low and Tasy freezing  -> it is the workstation: browser with too many tabs, old cache or full memory.'
    'diaginicioacionave.03' = 'That is why the startup time cannot be measured in this session.'
    'modulo3limpeza.41' = 'Applied: {0}'
    'modulo3limpeza.26' = '{0}: {1}'
    'diagsistemaacionav.01' = 'Machine state'
    'diagdownloads.08' = 'The 10 largest:'
    'invredeimpressoras.09' = '{0}{1,-42} {2}'
    'invpastaclinica.03' = 'Check the configured working folder.'
    'arquivosgrandes.04' = 'The Windows and Program Files folders are left out: there is nothing that can be done in them without administrator.'
    'diagpersistencia.10' = 'Whatever does not upload to the server in time is lost. It is worth checking with IT whether the roaming profile is saving.'
    'diagsessaoativa.10' = 'Do not close or compact now. The kit does not touch these processes, but closing the program from outside interrupts the batch - and in a nightly batch there is nobody to notice.'
    'arquivosgrandes.13' = 'THE 10 LARGEST DIRECTORIES'
    'alvo.20' = 'Adobe Acrobat/Reader cache'
    'modulo7sessao.08' = 'Reading session processes...'
    'diagnuvem.07' = 'OneDrive consuming {0}'
    'diaginicioacionave.22' = '{0} programs open on their own and can be safely disabled'
    'invinicializacaoco.01' = 'Inventory · everything that opens with Windows'
    'diagtasy.13' = 'Type ......: {0}'
    'diagcitrix.11' = 'No Citrix process running.'
    'diagpersistencia.03' = 'Roaming copy: {0}'
    'topo.16' = 'What will be applied'
    'item.38' = 'AppData\Local in the profile'
    'diagcitrix.35' = '{0}: portal up, asking for sign-in (HTTP {1})'
    'modulo3limpeza.48' = 'To go back now, without restarting: "REVERT TO ORIGINAL" button.'
    'alvo.01' = 'Recycle Bin'
    'invmigracao.09' = 'The profile folder does not have the user name: {0} uses the folder {1}'
    'modulo3limpeza.53' = 'Restart now'
    'diagtasy.53' = 'Network low and page high     -> slow application server. Ticket, with these numbers.'
    'diagsessaoativa.05' = 'CAUTION: there is a radioterapia.ai load in ANOTHER session ({0} process(es)):'
    'modulo3limpeza.50' = 'Free on C: now ................ {0} ({1}%)'
    'diaginicioacionave.12' = 'Kept as they are:'
    'modulo3limpeza.51' = 'Everything that was disabled comes back with the "Undo optimisations" button.'
    'diagmapeamentos.03' = '{0} -> {1}  [disconnected]'
    'diagtasy.05' = '{0} consuming {1} across {2} processes'
    'invecossistema.15' = '. It is the lie in the opposite direction, and it hides the version that is in use from support.'
    'diagsessaoativa.02' = 'Could not map the session processes.'
    'invofficeexcel.04' = 'The clinical folder is not in Excel''s trusted locations'
    'topo.35' = 'Support Kit ready.'
    'diagtasy.15' = 'external service, goes out over the internet'
    'explorer.02' = 'Could not restart Explorer (antivirus block).'
    'diagpersistencia.22' = '   LOST        {0,-30} loses {1}'
    'otimizacoes.06' = 'Cancel'
    'invecossistema.08' = '   {0,-24} v{1,-10} {2,-12} {3}'
    'diagtasy.22' = 'Measuring response time (4 samples)...'
    'alvo.04' = 'Google Chrome cache'
    'modulo3limpeza.54' = 'Restart later'
    'modulo3limpeza.33' = '{0} item(s) had already been removed by another step.'
    'diagcitrix.59' = 'If all of them fail, it is the network or the workstation profile.'
    'diagajustespendent.02' = 'All the performance settings are already applied'
    'inventario.01' = 'Report'
    'alvo.18' = 'Old Teams versions'
    'diagtasy.21' = 'If the other workstations also cannot open it, the server is down: immediate ticket.'
    'invinicializacaoco.05' = 'Only IT disables these. Record it on the ticket if sign-in is long.'
    'invsoftware.09' = 'Cleanup and restart need to be agreed in advance with whoever depends on the service.'
    'modulo7sessao.07' = 'before optimising a workstation that serves other people.'
    'modulo3analise.07' = 'Performance adjustments ....... {0}'
    'restartcomputador.05' = 'Windows did not accept the restart request (code {0}). The machine will NOT restart.'
    'diagsistemaacionav.06' = 'Fast Startup on: "Shut down" does not clear the memory, only "Restart".'
    'arquivosgrandes.14' = 'Measuring folder {0}/{1}: {2}'
    'diaginicioacionave.11' = '[ disable ] {0,-30} {1}'
    'diagtasy.42' = '{0} responds fast and {1} does not, from the same workstation and in the same minute.'
    'diaginicioacionave.08' = 'Startup in {0} seconds'
    'invpastaclinica.06' = '{0,-40} {1,10}'
    'diagdownloads.12' = 'Downloads with {0}'
    'diagmapeamentos.02' = 'No disconnected network drives'
    'topo.14' = 'Restore session (reopen)'
    'diagcitrix.08' = 'Citrix running: {0} process(es), {1} - {2}'
    'modulo3limpeza.16' = 'Removing programs from startup...'
    'diagnuvem.05' = 'Turn on "Files On-Demand" (click the cloud icon > gear > Settings) to free up space without losing access.'
    'diagtasy.01' = 'Tasy diagnostics (electronic medical record)'
    'item.19' = 'Messaging app'
    'topo.43' = 'Self-contained application: the preparation routine comes inside it.'
    'diagcitrix.14' = 'Invalid address in the configuration: {0}'
    'diagespaco.07' = 'System folder: only IT cleans it, and their cleanup routine already covers it. Record the size in the ticket.'
    'invecossistema.03' = 'Registry ..: {0}'
    'diagcitrix.29' = 'no response'
    'diagdownloads.03' = 'Downloads folder empty'
    'diagmemoria.04' = 'Total memory low for clinical use: {0}'
    'diagtasy.26' = 'Page (HTTP) ...: min {0} · typical {1} · max {2} ms   (HTTP {3})'
    'diaginicioacionave.01' = 'Windows startup'
    'diagnuvem.01' = 'OneDrive and Teams'
    'diagencerraveis.05' = 'Nothing to close: what is open is work or system.'
    'invidentificacao.09' = 'Below 8 GB, Excel over the network plus the browser already saturate the machine. Record it in the ticket.'
    'obs.09' = 'It may ask you to sign in again on older internal sites. That is why it comes unchecked.'
    'restartcomputador.02' = 'Warning'
    'modulo7sessao.06' = 'Services run as SYSTEM and are not included in the list, but check with whoever depends on them'
    'diagdownloads.10' = 'Downloads taking up {0}, of which {1} is more than 90 days old'
    'invinicializacaoco.03' = '[{0,-13}] {1,-36} {2}'
    'ajusteregistro.02' = '{0}: could not be adjusted.'
    'diagcitrix.13' = 'No store saved in the user profile (access via browser only).'
    'item.36' = 'Turn off Windows suggestions and promoted apps'
    'invecossistema.11' = 'The registry declares it but the disk does not have it: '
    'diagsistemaacionav.02' = '{0} · build {1} · {2} of memory'
    'topo.33' = 'The file could not be saved.'
    'diagnuvem.08' = 'Use the "Close OneDrive now" button when you need performance. It comes back when opened again.'
    'alvo.24' = 'Recent files list'
    'diagcitrix.45' = 'Change the shortcut to {0} . Over HTTP the browser blocks part of single sign-on and the traffic is sent unencrypted.'
    'modulo3limpeza.34' = '{0} item(s) in use were kept. Examples:'
    'invredeimpressoras.02' = '{0,-32} {1,-12} {2}'
    'diagmemoria.09' = 'Close the tabs that are not in use. Each open tab is a process with its own memory.'
    'arquivosgrandes.15' = 'CAN DELETE ...... {0} in {1} item(s) - disposable, Windows recreates it or no longer uses it'
    'diagmapeamentos.01' = 'Network drives'
    'diagtasy.07' = 'Tasy'
    'diaginicioacionave.17' = 'printer: {0}'
    'item.30' = 'not recognised'
    'modulo3limpeza.06' = 'YesNo'
    'arquivosgrandes.17' = 'DO NOT DELETE ... clinical data, database, medical-record browser or system file'
    'diagtasy.17' = '{0}: the name {1} does not resolve on this network'
    'invofficeexcel.06' = 'Excel'
    'diaginicioacionave.15' = '{0} network drive(s) and {1} network printer(s) are reconnected at every sign-in.'
    'diagtasy.49' = 'Module 3 has the item, but it comes UNSELECTED: it is the medical record browser and the decision to clean it is yours.'
    'inventariodiagnost.06' = 'Failure in step {0}: {1}'
    'otimizacoes.08' = 'Re-enabled at startup: {0}'
    'topo.10' = 'Large files and folders'
    'nota.09' = 'Makes single sign-on work and the .ica file open by itself.'
    'diagpersistencia.19' = 'Sign out (or restart) and run it again.'
    'diagtasy.30' = 'The network responds in {0} ms, so the delay is in the application server. Attach this log to the ticket: changing anything on the workstation will not help.'
    'alvo.21' = 'Office logs and diagnostics'
    'diagsistemaacionav.05' = 'Restart at the end of the cleanup (button in the side menu).'
    'openitemgrande.04' = 'Opened in Explorer: {0}'
    'diagtasy.45' = 'Java Web Start cache with {0}'
    'obs.13' = 'The recent list in Office and in Explorer disappears. No file is deleted.'
    'diagcitrix.03' = 'Receiver 4.x went out of support in 2018 and is replaced by Citrix Workspace. Updating requires IT, but it fixes a good share of the launch and single sign-on failures.'
    'invecossistema.01' = 'radioterapia.ai applications'
    'acaomemoria.03' = 'The first action inside each program may take a moment: it reloads what it needs.'
    'planolimpeza.01' = 'Measuring the Recycle Bin...'
    'arquivosgrandes.12' = '{0,2}. [{1,-15}] {2,10}  {3}'
    'modulo3analise.09' = '{0} startup item(s) were left unselected because they are not recognised:'
    'item.32' = 'Stop Store apps from running in the background'
    'topo.12' = 'What survives sign-out'
    'invservicos.02' = '{0} running, out of {1} installed'
    'diagdownloads.02' = 'Downloads folder not found.'
    'diagsessaoativa.03' = 'Measured in session {0}, of {1} active session(s) on this machine.'
    'arquivosgrandes.18' = 'No directory gets "CAN DELETE": a whole folder always requires checking.'
    'invpastaclinica.04' = 'Environment'
    'inventario.07' = 'Combine the .csv files from several machines into a single spreadsheet to compare the fleet.'
    'diaginicioacionave.10' = 'Safe NOT to load again (the program still works when it is opened):'
    'invpastaclinica.01' = 'Inventory · clinical folder'
    'alvo.17' = 'OneDrive logs'
    'invecossistema.16' = 'Ecosystem register matches the disk'
    'diagsessaoativa.12' = '[{0}] {1,-30} {2,10}  ({3} process(es))'
    'diagsessaoativa.08' = '   {0,-16} PID {1,-7} {2,10}   {3}'
    'diagtasy.03' = 'No client installed: Tasy runs in the browser (Wheb HTML5).'
    'diagcitrix.16' = 'Applications: {0}'
    'modulo3limpeza.23' = 'Applying {0}/{1}: {2}'
    'invofficeexcel.02' = 'Office version: {0}'
    'diagespaco.05' = 'The drives could not be read.'
    'diaginicioacionave.18' = '{0} network drive(s) not responding now: {1}'
    'diagtasy.24' = 'The web server is up and the application is not. Ticket with the exact time.'
    'diaginicioacionave.07' = 'Module 3 reduces this by disabling the safe items.'
    'diagcitrix.42' = 'This is what makes the browser ask for the password again and not open the .ica file by itself. Module 3 fixes it with one click.'
    'diagcitrix.20' = 'DNS .......: {0} -> {1} ({2} ms)'
    'diagnuvem.03' = 'OneDrive folder: {0} · {1}'
    'obs.07' = 'Screen thumbnails from remote sessions. They are recreated on the next connection.'
    'invredeimpressoras.05' = 'Network'
    'inventariodiagnost.18' = 'This diagnostics run already included the Citrix tests (three destinations) and the Tasy tests (three environments).'
    'modulo7sessao.04' = 'Audio and microphone are closed: the sound keeps working, only the enhancement layer goes.'
    'alvo.12' = 'Legacy Windows cookies (INetCookies)'
    'modulo3analise.06' = 'Items out of startup .......... {0}'
    'modulo3limpeza.03' = 'Information'
    'diagmapeamentos.04' = '{0} network drive(s) disconnected'
    'diagcitrix.26' = 'none responded'
    'item.34' = 'Clear the taskbar (search, task view, widgets)'
    'inventario.02' = 'Log server out of reach: {0}'
    'diagcitrix.39' = 'Detail: '
    'restartcomputador.03' = 'Restarting the computer...'
    'snapshotsvc.01' = 'Could not read this machine''s service list: the server-role check was NOT performed.'
    'nota.02' = 'Conversations stay on the server. Comes back when you open it.'
    'modulo3limpeza.24' = '{0}: {1} process(es) closed, {2} returned'
    'diaginicioacionave.16' = '{0} -> {1}   {2}'
    'invidentificacao.06' = 'Memory ........: {0}'
    'modulo3limpeza.36' = 'Check the Recycle Bin before emptying it again, in case you want to recover something.'
    'diagcitrix.47' = '{0}: HTTPS certificate not trusted on this workstation'
    'diagpersistencia.23' = 'Everything survived the last sign-out'
    'modulo3limpeza.12' = 'Closing dispensable programs...'
    'diagmemoria.06' = 'Failed to measure memory: {0}'
    'restaurarsessao.05' = 'Module 5 record: {0} program(s) closed.'
    'modulo3limpeza.04' = 'No item selected.'
    'diaginicioacionave.25' = 'Module 3 disables it.'
    'diagsessaoativa.15' = 'Scheduled tasks running now: {0}'
    'diagtasy.19' = 'Ports .....: {0}'
    'diagcitrix.23' = '{0}: the name {1} does not resolve on this network, not even with the full domain'
    'item.01' = 'OneDrive (sync)'
    'inventariodiagnost.08' = 'Nothing to fix here: the machine is clean on the user''s side.'
    'diagnuvem.13' = 'Teams is not running.'
    'nota.12' = 'Stops installing applications on its own.'
    'modulo3limpeza.47' = 'The current session is lighter now. The panel on the right shows the status.'
    'diagnuvem.12' = 'Use "Close Teams now" during heavy spreadsheet work. It comes back when it is opened again.'
    'nota.07' = 'Takes dozens of processes out of the background.'
    'diagdownloads.07' = 'no ext'
    'diagcitrix.01' = 'Citrix diagnostics (ARIA, MOSAIQ and Monaco)'
    'alvo.23' = 'Thumbnail and icon cache'
    'invmigracao.12' = 'Sign of a profile recreated during the migration. Recreating it again requires administrator.'
    'planolimpeza.02' = 'Network mapping not responding: {0} points to {1}'
    'obs.03' = 'It does not touch the profile holding bookmarks and passwords, which lives in Roaming.'
    'topo.31' = 'Could not copy. Select the text and use Ctrl+C.'
    'diagcitrix.55' = 'Citrix cache with {0} - this kit never cleans it'
    'diagdownloads.06' = '{0,-10} {1,10}  ({2} file(s))'
    'obs.11' = 'Close all of Office first. Changes not yet uploaded can be lost.'
    'item.20' = 'Game store'
    'diagsessaoativa.16' = '- {0}{1}'
    'diagpersistencia.05' = 'Nothing Module 3 adjusts in the registry survives, and Undo never finds anything to undo. Only IT can change this - take this log to the support ticket.'
    'invofficeexcel.09' = 'File > Options > Add-ins > COM Add-ins > deselect. It delays the opening of every spreadsheet.'
    'momentosessao.01' = 'This kit is a Windows app: it prepares the machine for work to start.'
    'diagcitrix.44' = '{0}: the shortcut uses HTTP, but the server also answers on HTTPS'
    'modulo3limpeza.30' = 'Recycle Bin: it was already empty'
    'modulo3analise.08' = 'Dead mappings removed ......... {0}'
    'invecossistema.07' = '{0} declared application(s):'
    'diagtasy.12' = 'Measuring '
    'diagpersistencia.12' = 'If the sign-out does not finish syncing, the adjustment is lost. Close the programs before leaving.'
    'alvo.09' = 'CentBrowser cache (medical-record browser)'
    'diagtasy.10' = '32-bit Java does not go beyond ~1.5 GB of memory. A large report freezes or closes on its own. Ask IT for 64-bit Java.'
    'inventario.05' = 'Spreadsheet: {0}'
    'diagtasy.33' = '{0}: typical latency of {1} ms for an internal server'
    'diagcitrix.04' = 'Citrix'
    'diagajustespendent.04' = '[ apply ] {0}'
    'acaoprioridade.01' = 'Priority adjusted: {0} clinical program(s) above normal, {1} below normal.'
    'modulo7sessao.14' = '{0} program(s) with an open window were left unselected because they are not recognised:'
    'alvo.15' = 'Crash memory dumps'
}

$script:Textos['es'] = @{
    'item.15' = 'Tiendas de juegos'
    'diagcitrix.18' = 'Probando '
    'topo.37' = 'Módulo 2: inventario de la máquina + diagnóstico de lo que se puede resolver aquí.'
    'invsoftware.04' = 'Software clínico:'
    'diagtasy.23' = '{0}: puertos abiertos pero la página no responde'
    'diagencerraveis.03' = 'Ningún programa prescindible en ejecución ahora'
    'diagespaco.03' = 'Disco'
    'invredeimpressoras.12' = 'Impresora'
    'item.05' = 'Xbox y Game Bar'
    'diaginicioacionave.05' = 'El inicio llevó {0} segundos'
    'invecossistema.09' = '      {0}'
    'arquivosgrandes.16' = 'ATENCIÓN-EVALUAR. {0} en {1} elemento(s) - revise el contenido antes de mover o borrar'
    'planolimpeza.04' = 'Credencial guardada de un servidor que no responde: {0}'
    'obs.05' = 'Una caché antigua de Java hace que Tasy falle al descargar un archivo. Se descarga otra vez en el siguiente acceso.'
    'invredeimpressoras.07' = 'La tarjeta y el switch son de 1 Gbps: 100 Mbps casi siempre es cable viejo, mal crimpado o puerto defectuoso. Cambiar el cable no cuesta nada y mejora Tasy, Citrix y la apertura de los portales de 40 MB.'
    'invidentificacao.07' = 'Windows desde .: {0:dd/MM/yyyy}'
    'modulo7sessao.02' = 'No se desinstala nada. Lo que hace este módulo vuelve en el próximo inicio de sesión.'
    'diagpersistencia.28' = '. El disco fuera del perfil se mantuvo. Esto es un perfil descartado, y solo TI lo corrige.'
    'diagsessaoativa.11' = 'Los mayores prescindibles (j = con ventana abierta):'
    'modulo3limpeza.01' = 'Ejecute primero el Módulo 3 para analizar la máquina.'
    'modulo3analise.02' = 'Nada se modifica en este paso. Todo lo que es seguro ya viene marcado.'
    'arquivosgrandes.07' = 'El archivo de paginación reside en ese disco: si se llena, toda la máquina se bloquea. Libere espacio en él también.'
    'topo.38' = 'Módulo 3: analiza, deja todo marcado y resuelve en un clic ("Aplicar todo").'
    'diagcitrix.24' = 'La estación de trabajo no está viendo el DNS interno de ese servidor. Confirme el cable/VPN y el sufijo DNS de la tarjeta antes de abrir un ticket.'
    'topo.25' = 'Vale solo para la sesión actual. El siguiente inicio de sesión devuelve la máquina al estado normal.'
    'topo.18' = 'Quitar lo que tenga más de:'
    'otimizacoes.01' = 'Deshacer optimizaciones'
    'idioma.01' = 'Idioma'
    'diagtasy.57' = 'Pida a TI que excluya las carpetas de caché del navegador y de Java del análisis del antivirus: es una ganancia directa en la navegación de Tasy.'
    'modulo3limpeza.42' = 'Resultado'
    'topo.03' = 'Solo acciones sin administrador  ·  versión {0}  ·  {1}\{2}'
    'diagnuvem.04' = 'OneDrive sincronizando {0} en el disco local'
    'invmigracao.15' = 'Windows intenta usar cada una al abrir el Explorer y espera el tiempo límite. El Módulo 3 las elimina - el elemento viene DESMARCADO porque no tiene deshacer.'
    'invidentificacao.04' = 'Equipo ........: {0} {1}   serie {2}'
    'diagcitrix.30' = '{0}: el servidor {1} no responde en los puertos web'
    'diagpersistencia.21' = '   SOBREVIVIÓ  {0,-30} guarda {1}'
    'restaurarsessao.07' = '{0}: reabierto.'
    'diagtasy.20' = '{0}: servidor {1} sin respuesta'
    'invmigracao.06' = 'Ningún agente de migración instalado.'
    'item.11' = 'Servicios de Apple'
    'invsoftware.02' = '{0} programas instalados'
    'diagsessaoativa.14' = 'El servicio se ejecuta como SISTEMA y solo el área de TI lo apaga. El Módulo 5 cierra las utilidades de Bluetooth de su sesión, pero el servicio continúa. Vale pedir la desactivación masiva en las estaciones de trabajo de consultorio.'
    'invmigracao.07' = 'Usuario conectado ..: {0}'
    'item.03' = 'Actualizadores automáticos'
    'diaglixeira.02' = 'Papelera de reciclaje con {0}'
    'diagcitrix.21' = '{0}: el nombre corto {1} no resuelve, solo el completo {2}'
    'item.26' = 'Juegos y superposiciones'
    'invecossistema.02' = 'Ningún registro del ecosistema en esta estación de trabajo. Es el estado normal de quien no ha instalado nada.'
    'nota.05' = 'Se cierra solo si no hay ninguna ventana abierta.'
    'diagpersistencia.15' = 'El disco vuelve al estado anterior en cada reinicio. Nada instalado o ajustado permanece. Solo TI lo desactiva - lleve este log al ticket de soporte.'
    'diagdownloads.09' = '{0,10}  {1}  ({2:dd/MM/yyyy})'
    'modulo3limpeza.22' = 'Interrumpido por el usuario.'
    'diagcitrix.48' = 'El navegador muestra un aviso de sitio no seguro y puede bloquear la descarga y la impresión. Solicite a TI la instalación del certificado de la autoridad interna.'
    'obs.01' = 'Los archivos en uso se omiten automáticamente.'
    'modulo3limpeza.39' = 'Ya no se abrirá solo: {0}'
    'diagcitrix.57' = 'Son solo lanzadores descargados por el navegador. El Módulo 3 elimina los de Downloads dentro del período elegido; la caché de Citrix queda intacta.'
    'diagtasy.56' = 'Un entorno rápido y otro lento -> el problema es de ese servidor, no suyo.'
    'topo.23' = 'EVALUAR = puede rendir espacio · NO BORRAR = déjelo como está. Doble clic abre en Explorer.'
    'inventariodiagnost.12' = '{0,-16} {1,-26} {2} hallazgo(s), {3} de alto impacto'
    'diagpersistencia.07' = 'Perfil temporal: toda esta sesión se descarta al cerrar la sesión'
    'topo.41' = 'Citrix y Tasy se prueban dentro del Módulo 2.'
    'otimizacoes.05' = 'Deshacer todo'
    'invecossistema.05' = 'Mientras esto dure ninguna aplicación puede registrar la versión, y el soporte se queda sin saber qué está instalado. Lleve este log a quien se encarga del instalador.'
    'diagcitrix.33' = 'Por encima de 3 segundos la lista de aplicaciones tarda en aparecer. Limpiar la caché de Workspace en el Módulo 3 suele resolverlo del lado de la estación de trabajo.'
    'diagtasy.08' = 'Java en ejecución: {0} · {1} · {2}'
    'alvo.07' = 'Caché de Teams (versión nueva)'
    'topo.05' = 'Copiar log'
    'explorer.03' = 'Los cambios de la barra de tareas aparecen en el próximo inicio de sesión.'
    'modulo3limpeza.57' = 'Limpieza completada.'
    'modulo3limpeza.58' = 'Cerrando {0}...'
    'modulo3limpeza.59' = '{0}: no se pudo cerrar (protegido o ya cerrado)'
    'modulo3limpeza.60' = 'Tarea detenida: {0}{1}'
    'modulo3limpeza.61' = '{0}: sin permiso para detenerla (se ejecuta como SISTEMA - solo TI)'
    'invpastaclinica.08' = 'Preparación incompleta: falta {0}'
    'diagcitrix.06' = 'Citrix Workspace no encontrado en esta máquina'
    'openitemgrande.03' = 'Ya no existe: {0}'
    'modulo3limpeza.18' = 'Cerrando programas prescindibles de la sesión...'
    'restartcomputador.04' = 'Reinicio programado para 10 segundos. Use shutdown /a para cancelar.'
    'invecossistema.14' = 'Hay entrega.txt en el disco pero no aparece en el registro: '
    'topo.39' = 'Módulo 4: los 10 archivos más grandes y las 10 carpetas más grandes de C:, con enlace al Explorer.'
    'item.17' = 'Skype'
    'diagespaco.01' = 'Espacio en disco'
    'invsoftware.06' = '[ ] {0}'
    'item.07' = 'Vincular al teléfono'
    'diagtasy.39' = '{0,-34} {1,10} {2,10}  {3}'
    'diagtasy.37' = 'Zona ......: fuera de la Intranet de Windows - sin efecto aquí, Tasy tiene su propio inicio de sesión.'
    'item.22' = 'Actualizador automático'
    'diagtasy.28' = 'Oscilación: una de las muestras tardó {0} ms frente a {1} ms de las otras. Red inestable, no lenta.'
    'modulo3analise.11' = 'Desmarque lo que no quiera y haga clic en "Aplicar todo".'
    'invidentificacao.03' = 'Sistema .......: {0} compilación {1}'
    'diagpersistencia.26' = 'El perfil no guarda los cambios: {0} de {1} lugares perdieron el marcador'
    'diagcaches.02' = '{0} recuperables solo con limpieza segura'
    'diagsessaoativa.04' = '{0,-46} {1,4} proceso(s)  {2,10}'
    'invecossistema.10' = '{0} entrada(s) del registro apuntan a una carpeta que no existe'
    'diagsessaoativa.13' = 'Soporte de Bluetooth activado en una estación de trabajo cableada'
    'modulo7sessao.05' = 'ATENCIÓN: esta máquina hospeda un servicio ({0}).'
    'topo.27' = 'Actualizar'
    'inventariodiagnost.16' = 'Todo lo que está arriba lo resuelve el Módulo 3 (Limpieza). Ábralo y haga clic en "Aplicar todo".'
    'invmigracao.02' = 'Agente de migración instalado: {0} {1}'
    'restartcomputador.01' = 'Reiniciar el equipo'
    'otimizacoes.03' = 'Deshacer'
    'nota.15' = 'Sube lo clínico por encima de lo normal y baja el resto. Se restablece en el próximo inicio de sesión.'
    'diagnuvem.06' = 'Nube'
    'topo.01' = 'Workstation Kit'
    'modulo3limpeza.29' = 'Papelera de reciclaje: {0}'
    'invredeimpressoras.08' = 'Sufijo DNS: {0}'
    'diagcitrix.52' = 'El tráfico de Citrix está pasando por el proxy sin necesidad, lo que hace lenta la apertura. Pida a TI que lo incluya en las excepciones.'
    'arquivosgrandes.05' = 'La paginación de Windows está en {0} ({1}% libre). Mientras ese disco tenga espacio, no pesa en el rendimiento.'
    'diagajustespendent.05' = '{0} ajuste(s) de rendimiento aún no aplicado(s)'
    'modulo7sessao.09' = 'Montando el plan...'
    'item.44' = 'Google Drive'
    'restaurarsessao.12' = 'Lo que el Módulo 3 quitó del inicio automático vuelve con el botón "Deshacer optimizaciones".'
    'restaurarsessao.01' = 'Restaurar la sesión'
    'modulo3limpeza.08' = 'Limpieza cancelada por el usuario.'
    'diaginicioacionave.21' = 'El tiempo no viene de los programas de su usuario. Son {0} unidad(es) de red y {1} impresora(s) de red reconectadas en cada inicio de sesión, más los agentes corporativos. Lleve estos números al área de TI.'
    'item.13' = 'Nube personal'
    'diagpersistencia.02' = 'Perfil ....: {0}'
    'arquivosgrandes.19' = 'Seleccione cualquier fila en el panel de la derecha y haga clic en "Abrir en Explorer".'
    'diagcitrix.36' = '{0}: el portal respondió HTTP {1}'
    'item.24' = 'Asistente del fabricante'
    'alvo.02' = 'Archivos temporales del usuario'
    'diagtasy.47' = 'Caché de CentBrowser: {0}'
    'diagpersistencia.17' = 'Cierre la sesión (o reinicie) y ejecute este módulo de nuevo: dirá exactamente qué sobrevivió.'
    'modulo7sessao.10' = 'Programas cerrados ............ {0}'
    'diagtasy.04' = 'Navegador {0,-14} {1,10} en {2} proceso(s)'
    'restaurarsessao.08' = '{0}: no encontré el ejecutable para reabrir. Ábralo desde el menú Inicio.'
    'invredeimpressoras.10' = 'Impresora predeterminada sin conexión: {0}'
    'diagpersistencia.09' = 'La política borra la copia local del perfil al cerrar sesión'
    'diagtasy.14' = 'servidor interno de la red'
    'invofficeexcel.05' = 'Ejecute el Módulo 1 (Preparar entorno) en esta máquina.'
    'logemuso.02' = '{0}: {1} protege {2} proceso(s), {3}.'
    'arquivosgrandes.10' = 'LOS 10 ARCHIVOS MÁS GRANDES'
    'diagcitrix.31' = 'Mientras esto no vuelva, las aplicaciones de ese destino no abren. Verifique la red de la estación de trabajo y, si las otras estaciones también fallan, avise a TI.'
    'inventariodiagnost.09' = 'Si la lentitud continúa, el cuello de botella exige administrador (disco, antivirus, red, hardware).'
    'otimizacoes.07' = 'Restaurado: {0}'
    'inventariodiagnost.02' = 'Máquina {0} · usuario {1} · {2:dd/MM/yyyy HH:mm:ss}'
    'explorer.01' = 'Explorer reiniciado: la barra de tareas ya aparece limpia.'
    'modulo3limpeza.55' = 'Reinicio programado para 15 segundos. Use shutdown /a en Ejecutar para cancelar.'
    'alvo.05' = 'Caché de Firefox'
    'openitemgrande.01' = 'Seleccione una fila de la lista.'
    'diagajustespendent.01' = 'Ajustes de rendimiento pendientes'
    'alvo.19' = 'Instaladores descargados por Edge/Chrome'
    'planolimpeza.06' = 'Credenciales'
    'topo.15' = 'Herramienta de apoyo técnico.
No validada para uso clínico.
Ninguna acción exige administrador.
Todo lo que se desactiva vuelve con Deshacer.'
    'diagnuvem.14' = 'Versiones antiguas de Teams ocupando {0}'
    'diagsessaoativa.19' = 'Sesión reducida: {0} en prescindibles'
    'topo.13' = 'REVERTIR Y CERRAR EL CICLO'
    'modulo3analise.04' = 'Espacio liberado en disco ..... {0}'
    'invsoftware.08' = 'Esta máquina aloja un servicio (base de datos o aplicación)'
    'item.09' = 'Búsqueda y Copilot'
    'item.42' = 'Dar prioridad de CPU a Citrix, la historia clínica y Office'
    'diagtasy.35' = '{0}: respuesta inestable (varió de {1} a {2} ms)'
    'modulo7sessao.13' = 'Memoria compactada ............ navegador y Office (Citrix intacto)'
    'diagtasy.38' = 'COMPARACIÓN ENTRE LOS ENTORNOS'
    'diaginicioacionave.09' = 'Elementos de su usuario que se abren junto con Windows: {0}'
    'diagdownloads.01' = 'Carpeta Downloads'
    'invpastaclinica.07' = '{0,-40} {1,10}  {2:dd/MM/yyyy HH:mm}'
    'diagtasy.43' = 'Esto descarta la estación de trabajo y la red local: el problema está en el servidor lento.'
    'diagsessaoativa.07' = 'Carga del ecosistema en ejecución (protegida de los Módulos 3 y 5):'
    'invidentificacao.01' = 'Inventario · identificación'
    'diagcitrix.09' = 'Hay una sesión publicada abierta ahora (ARIA, MOSAIQ o Monaco)'
    'topo.30' = 'Log copiado. Péguelo en el correo o en el ticket.'
    'diaginicioacionave.14' = 'Lo que más pesa en el inicio de sesión de esta estación de trabajo:'
    'diagpersistencia.24' = 'Ni el disco fuera del perfil sobrevivió al cierre de sesión'
    'diagencerraveis.08' = 'El Módulo 3 los cierra todos a la vez. Vuelven cuando usted los abra.'
    'item.18' = 'Zoom'
    'diagsessaoativa.09' = 'Trabajo de radioterapia.ai en curso {0}: {1} usando {2}'
    'diagcitrix.15' = '--- {0} ---'
    'inventariodiagnost.14' = 'RESUELVA EN ESTE ORDEN:'
    'diagcitrix.28' = 'responde'
    'invofficeexcel.01' = 'Inventario · Office y Excel'
    'diagespaco.08' = 'C:\Windows\Temp: {0} (del sistema, fuera del alcance del usuario)'
    'diagencerraveis.06' = '{0,-34} {1,10}   ({2} proceso(s))'
    'diagtasy.27' = 'Procesamiento .: ~{0} ms en el servidor (HTTP menos red)'
    'diaginicioacionave.02' = 'La máquina estuvo encendida {0:N1} h antes de este inicio de sesión.'
    'modulo3limpeza.27' = '{0}: nada que limpiar'
    'invidentificacao.08' = 'Memoria total de {0}'
    'diagtasy.46' = 'Un caché antiguo hace que Tasy abra la versión equivocada y falle al descargar un archivo. El Módulo 3 lo limpia.'
    'diagencerraveis.04' = 'Memoria'
    'diagtasy.50' = 'Proxy sin excepción para el Tasy interno: {0}'
    'modulo3limpeza.21' = 'Compactando la memoria de los programas que se quedan...'
    'diagmemoria.01' = 'Memoria RAM'
    'inventariodiagnost.04' = 'Diagnóstico interrumpido por el usuario.'
    'diaginicioacionave.23' = 'Cada uno de ellos compite por disco y CPU en el inicio de sesión. El Módulo 3 los desactiva todos a la vez y el botón "Deshacer" lo revierte.'
    'diaginicioacionave.04' = 'Para medir: reinicie y ejecute el Módulo 2 justo después de iniciar sesión.'
    'ajusteregistro.01' = '{0}: bloqueado por política de TI (Windows no deja que el usuario lo cambie).'
    'item.39' = 'AppData\Roaming del perfil'
    'diagdownloads.11' = 'Use el botón "Vaciar Downloads" en el menú lateral. Va a la Papelera de reciclaje, se puede recuperar.'
    'diagcitrix.34' = '{0}: portal respondiendo (HTTP {1}, {2} ms)'
    'diagsistemaacionav.03' = 'Reinicie el equipo al final de la limpieza. Usar "Apagar" no lo resuelve: solo "Reiniciar" limpia la memoria.'
    'diaginicioacionave.20' = 'Inicio de sesión largo con solo {0} elemento(s) de inicio automático del usuario'
    'modulo3limpeza.49' = 'Espacio liberado .............. {0}'
    'modulo3limpeza.52' = 'Limpieza aplicada.

La mejora en el inicio automático y en los ajustes solo aparece después de reiniciar.

¿Reiniciar ahora? Guarde sus archivos antes.'
    'titulo.01' = '== {0} {1}'
    'modulo3limpeza.35' = '   ... y otros {0}.'
    'invredeimpressoras.01' = 'Inventario · red e impresoras'
    'diagcitrix.38' = '{0}: el portal no respondió'
    'restaurarsessao.02' = 'Devuelve las prioridades a la normalidad y reabre lo que se cerró en el Módulo 5.'
    'item.28' = 'Edge abriéndose solo al iniciar sesión'
    'item.40' = 'Disco fuera del perfil'
    'diagtasy.06' = 'Tasy se ejecuta dentro del navegador: con esa carga la pantalla se congela incluso con el servidor respondiendo rápido. Cierre las pestañas que no usa antes de abrir un ticket.'
    'diagtasy.51' = 'El tráfico de la historia clínica pasa por el proxy sin necesidad y eso suma tiempo en cada pantalla. Pida a TI que lo incluya en las excepciones.'
    'diagmapeamentos.05' = 'Una asignación muerta congela el Explorer y el "Guardar como" de Office durante 30 segundos o más. El Módulo 3 la elimina.'
    'invecossistema.12' = '. Un registro que miente es peor que un registro vacío: vacío hace buscar, equivocado hace concluir. Llévelo al responsable del instalador - el kit no corrige el registro de otro proyecto.'
    'diagtasy.16' = 'DNS .......: {0} ({1} ms)'
    'invofficeexcel.07' = 'Complemento: {0} (LoadBehavior {1})'
    'diagmemoria.11' = 'Cierre y abra Teams una vez al día, o limpie la caché en el Módulo 3.'
    'modulo3analise.01' = 'Módulo 3 - Limpieza (análisis)'
    'diagmemoria.08' = '{0,-24} {1,10}   ({2} proceso(s))'
    'inventariodiagnost.07' = 'Mapa de los puntos críticos'
    'otimizacoes.09' = 'Inicio automático de {0} restaurado.'
    'diagtasy.41' = 'situación'
    'nota.08' = 'Evita que el disco se vuelva a llenar.'
    'diagpersistencia.18' = 'Marcador de {0}, del mismo inicio de sesión. Todavía no hubo cierre de sesión para comparar.'
    'modulo3limpeza.32' = 'Downloads: {0} elementos a la Papelera de reciclaje, {1} liberados'
    'achado.01' = '-> '
    'diagajustespendent.03' = 'Ajustes'
    'diagcitrix.58' = 'Si un destino falla y los otros funcionan, el problema es de ese servidor, no de la estación de trabajo.'
    'modulo3limpeza.25' = '{0}: ya no estaba en ejecución'
    'modulo3limpeza.14' = 'Vaciando la Papelera de reciclaje...'
    'invinicializacaoco.04' = '{0} elementos de inicio automático pertenecen a la máquina, no a su usuario'
    'topo.17' = 'Todo lo que es seguro ya viene marcado. Desmarque lo que no quiera y haga clic en APLICAR TODO.'
    'ajustezonascitrix.01' = '{0} agregado a los sitios de la Intranet.'
    'diagpersistencia.20' = 'Marcador anterior de {0}. Hubo un cierre de sesión desde entonces - este es el resultado:'
    'modulo3limpeza.11' = 'Módulo 3 - Aplicando'
    'diagcitrix.50' = 'Excepciones.: {0}'
    'alvo.10' = 'Caché de imágenes de Escritorio remoto'
    'modulo7sessao.11' = 'Tareas programadas detenidas .. {0}'
    'topo.08' = 'Detener'
    'invsoftware.11' = 'Agentes corporativos en ejecución: {0}'
    'invpastaclinica.05' = 'Total: {0}'
    'invidentificacao.11' = '{0} -> {1}'
    'diagencerraveis.01' = 'Programas que se pueden cerrar ahora con seguridad'
    'restaurarsessao.09' = '{0} elemento(s) vuelven solos cuando Windows los necesite:'
    'logemuso.03' = '   Ninguno de ellos será cerrado, rebajado ni compactado. Puerto {0}.'
    'inventario.06' = 'Guardada en la carpeta de LOGs de la red. El informe de texto es este mismo log.'
    'diagnuvem.15' = 'El Módulo 3 lo quita con el elemento "Versiones antiguas de Teams".'
    'acaoprioridade.02' = 'Windows restablece las prioridades en el próximo inicio de sesión.'
    'diagtasy.02' = 'Instalado: {0} {1}'
    'arquivosgrandes.02' = 'Solo lectura. El kit no borra nada aquí: usted decide elemento por elemento.'
    'diagcitrix.17' = 'Dirección .: {0}'
    'topo.34' = 'Pantalla limpia.'
    'diagcaches.04' = '{0} recuperables con limpieza segura'
    'inventariodiagnost.05' = 'Diagnóstico {0}/{1}: {2}'
    'invofficeexcel.03' = 'Ubicación de confianza: {0}'
    'diagespaco.06' = 'C:\Windows\Temp con {0}'
    'diaginicioacionave.19' = 'Windows intenta reconectar cada una de ellas en el inicio de sesión y espera el tiempo límite. Es la explicación más probable del inicio de sesión largo. Pida a TI que quite la asignación o que libere el acceso.'
    'modulo3analise.10' = 'Márquelo en la lista si sabe que se puede desactivar.'
    'diagtasy.25' = 'Red (TCP) .....: mín {0} · típico {1} · máx {2} ms'
    'diagnuvem.02' = 'OneDrive en ejecución · {0}'
    'diagcitrix.41' = '{0}: {1} no está en la zona de Intranet'
    'obs.08' = 'Carpeta de instalación que usa TI. Los archivos en uso se omiten.'
    'invredeimpressoras.04' = 'Los portales tienen 40 MB y se abren desde la red. El cable reduce de forma notable el tiempo de apertura y evita la corrupción del archivo.'
    'item.10' = 'Asistentes del fabricante'
    'inventario.03' = 'La hoja de cálculo no se guardó. El informe es este log ("Copiar log" o "Guardar log").'
    'modulo3limpeza.17' = 'Aplicando ajustes de rendimiento...'
    'inventariodiagnost.10' = 'Copie este log y abra un ticket: ya muestra que el lado del usuario está limpio.'
    'invinicializacaoco.02' = '{0} elementos (usuario + máquina)'
    'alvo.13' = 'Caché de internet de Windows'
    'otimizacoes.11' = 'Todo restaurado. Cierre la sesión para aplicarlo por completo.'
    'diagcitrix.12' = 'Stores ya configurados en el Workspace de este usuario:'
    'modulo3limpeza.15' = 'Enviando Downloads antiguos a la Papelera de reciclaje...'
    'item.37' = 'Registro del usuario (HKCU)'
    'diagtasy.54' = 'Red alta y página alta        -> camino de red de la estación de trabajo. Cable en vez de Wi-Fi, y ticket de red.'
    'modulo3limpeza.07' = 'Question'
    'acaomemoria.01' = 'Ningún programa aceptó la compactación en esta máquina.'
    'alvo.25' = 'Vaciar la papelera de reciclaje'
    'diagnuvem.09' = 'OneDrive no está en ejecución.'
    'invredeimpressoras.03' = 'Estación de trabajo conectada por Wi-Fi ({0})'
    'diagcitrix.46' = 'Seguridad .: dirección solo en HTTP. El tráfico va sin cifrado.'
    'modulo3limpeza.37' = 'REGISTRO: el Escritorio o Documentos de este usuario están dentro de OneDrive.'
    'diagtasy.40' = 'página'
    'restaurarsessao.04' = 'Ningún registro de cierre en esta sesión. Nada que reabrir.'
    'modulo3limpeza.05' = 'Aplicar limpieza'
    'diaginicioacionave.24' = '{0} programa(s) en el inicio automático se pueden desactivar'
    'topo.32' = 'Log guardado en '
    'invmigracao.04' = 'La migración terminó hace años. Quitarlo requiere administrador: pida al área de TI la desinstalación masiva. Ejecuta un servicio y recorre el perfil en cada inicio de sesión.'
    'diagtasy.18' = 'La estación de trabajo no está viendo el DNS. Es un problema de red de la estación, no de la historia clínica.'
    'modulo3analise.03' = 'Lo que va a ocurrir al hacer clic en "Aplicar todo":'
    'diaglixeira.01' = 'Papelera de reciclaje'
    'diagcitrix.22' = 'Cambie la dirección del acceso directo y de la configuración a {0} . Sin esto el navegador no encuentra el portal en esta estación de trabajo.'
    'item.29' = 'protegido (seguridad, red o clínico)'
    'diagnuvem.10' = 'Teams en ejecución · {0} en {1} proceso(s)'
    'nota.06' = 'Hace que una máquina lenta responda de inmediato.'
    'diagdownloads.04' = '{0} archivos · {1} en total'
    'modulo3limpeza.45' = 'Total recuperado .............. {0}'
    'invmigracao.11' = 'Perfil con sufijo de dominio: {0}'
    'diagmemoria.10' = 'Teams consumiendo mucha memoria'
    'alvo.22' = 'Caché de documentos de Office'
    'topo.28' = 'Período cambiado a: {0}. Recalculando...'
    'openitemgrande.05' = 'No se pudo abrir: {0}'
    'obs.12' = 'Las miniaturas se recrean solas.'
    'modulo3limpeza.20' = 'Ajustando la prioridad de CPU...'
    'acaomemoria.02' = 'Memoria compactada en {0} programa(s): {1} devueltos a Windows.'
    'invinicializacaoco.06' = 'Inicio automático'
    'diagpersistencia.04' = 'Perfil obligatorio: Windows descarta sus cambios en cada cierre de sesión'
    'arquivosgrandes.11' = 'Ningún archivo por encima de {0} MB.'
    'invservicos.03' = '{0,-40} {1,-12} {2}'
    'diagencerraveis.10' = 'Los programas de trabajo, clínicos, de seguridad y los navegadores nunca son cerrados por el kit.'
    'item.33' = 'Activar la limpieza automática de Windows'
    'invecossistema.06' = 'Ecosistema'
    'invservicos.01' = 'Inventario · servicios en ejecución'
    'diagpersistencia.11' = 'Perfil móvil: lo que vale es la copia del servidor'
    'modulo3limpeza.28' = '{0}: {1} liberados{2}'
    'topo.06' = 'Guardar log'
    'modulo3limpeza.10' = 'Módulo 5 - Aplicando en la sesión'
    'topo.07' = 'Limpiar pantalla'
    'inventario.04' = 'Carpeta de logs no disponible: {0}'
    'diagcitrix.05' = 'Citrix Workspace instalado (versión {0})'
    'alvo.16' = 'Caché de vídeo (shaders)'
    'topo.36' = 'Máquina {0}  ·  usuario {1}  ·  {2:dd/MM/yyyy HH:mm}'
    'item.14' = 'Mensajería'
    'diagcitrix.19' = '...'
    'diagespaco.04' = 'Ejecute el Módulo 3 para liberar espacio.'
    'tarefasemexecucao.01' = 'Leyendo tareas programadas (algunos segundos)...'
    'item.04' = 'Adobe Creative Cloud (segundo plano)'
    'invredeimpressoras.11' = 'Excel y Word se congelan al abrir un archivo si la impresora predeterminada no responde. Cambie la predeterminada a "Microsoft Print to PDF" mientras no lo resuelvan.'
    'modulo3limpeza.40' = 'No se pudo desactivar {0}'
    'diaginicioacionave.06' = 'Vea a continuación lo que pesa: elementos de inicio automático, unidades de red e impresoras de red. El Módulo 3 resuelve la parte del usuario.'
    'diagcitrix.43' = 'Seguridad .: el servidor redirige por sí solo a HTTPS. El acceso directo puede quedar como está.'
    'invredeimpressoras.06' = 'Tarjeta de red negociando solo {0}'
    'comprotecao.01' = 'Error inesperado en {0}: {1}'
    'obs.06' = 'SIN MARCAR a propósito: es el navegador de la historia clínica. No borra favoritos ni contraseñas, solo la caché. Márquelo si Tasy muestra pantalla en blanco o error de archivo.'
    'topo.40' = 'Módulo 5: deja la sesión actual liviana, preservando Citrix, la historia clínica y los navegadores.'
    'diagpersistencia.29' = 'Marcador renovado en {0} de {1} lugares.'
    'modulo7sessao.03' = 'Citrix, el navegador de la historia clínica, el navegador web, Office y el acceso remoto quedan siempre fuera.'
    'diagcitrix.40' = 'Zona ......: ya está en la Intranet (el inicio de sesión único funciona)'
    'alvo.11' = 'Caché de Opera y Vivaldi'
    'diagcaches.01' = 'Espacio recuperable (vista previa de la limpieza)'
    'modulo3limpeza.02' = 'Limpieza'
    'diagcitrix.27' = 'ICA 1494 ..: {0}   ·   Fiabilidad 2598: {1}{2}'
    'item.35' = 'Quitar el icono de OneDrive del Explorer'
    'nota.03' = 'Vuelven solos cuando haya una actualización.'
    'nota.16' = 'Devuelve a Windows la memoria que el navegador y Office reservaron sin usar. La sesión publicada de Citrix no se toca, para que no se atasque.'
    'diagsessaoativa.17' = '{0} en {1} programa(s) prescindible(s) en la sesión actual'
    'restaurarsessao.10' = 'La capa de audio y los iconos de la bandeja vuelven en el próximo inicio de sesión.'
    'diagmemoria.07' = 'Programas que más consumen memoria:'
    'otimizacoes.10' = 'Zona eliminada: {0}'
    'invmigracao.14' = '{0} credencial(es) guardada(s) apuntando a un servidor inexistente'
    'diagcitrix.51' = 'Proxy sin excepción para: {0}'
    'arquivosgrandes.08' = 'Buscando archivos grandes...'
    'topo.19' = 'Marcar todo'
    'diagtasy.31' = 'Red en {0} ms. La lentitud está en el servidor, no en la estación de trabajo.'
    'diagajustespendent.06' = 'El Módulo 3 los aplica todos. Reversible con el botón "Deshacer optimizaciones".'
    'restaurarsessao.06' = '{0}: ya está abierto.'
    'nota.10' = 'Desaparece de la barra y deja de consumir memoria. Reversible.'
    'invsoftware.01' = 'Inventario · software instalado'
    'item.02' = 'Microsoft Teams'
    'diaglixeira.03' = 'Vaciar la Papelera de reciclaje es una rutina del Módulo 3. Compruebe antes que no haya nada que recuperar.'
    'invpastaclinica.02' = '{0} no existe'
    'item.27' = 'Utilidad opcional'
    'nota.13' = 'Se hace antes de Downloads, para que lo que salga de allí aún pueda recuperarse.'
    'modulo3limpeza.44' = 'Memoria compactada ............ {0}'
    'nota.04' = 'Windows lo vuelve a cargar solo cuando lo necesita.'
    'invmigracao.13' = '{0} credencial(es) guardada(s) para servidores que ya no existen:'
    'planolimpeza.03' = 'Para eliminarlo, ejecute en el Símbolo del sistema:  net use {0} /delete'
    'invmigracao.08' = 'Carpeta del perfil .: {0}'
    'diagsessaoativa.01' = 'Sesión actual: lo que está activo ahora'
    'diagcitrix.54' = 'Archivos .ica sueltos: {0}'
    'obs.02' = 'No borra favoritos, contraseñas ni historial.'
    'diagtasy.09' = 'Cliente Java de Tasy ejecutándose en 32 bits'
    'topo.29' = 'Deteniendo después del paso actual...'
    'diaginicioacionave.27' = 'Solo TI cambia los elementos de la máquina (para todos los usuarios) y los de seguridad.'
    'arquivosgrandes.09' = 'Archivos grandes en '
    'inventariodiagnost.13' = '{0} problema(s) crítico(s) y {1} advertencia(s).'
    'diagpersistencia.06' = 'Persistencia'
    'topo.42' = 'Este kit solo hace lo que se ejecuta sin administrador.'
    'invofficeexcel.08' = 'El complemento {0} se carga junto con Excel'
    'item.31' = 'Desactivar animaciones, sombras y transparencia'
    'diagtasy.52' = 'CÓMO LEER EL RESULTADO'
    'modulo7sessao.15' = 'Márquelo en la lista si sabe que se puede cerrar.'
    'invecossistema.04' = 'Registro del ecosistema ilegible: {0}'
    'diagpersistencia.13' = 'Perfil local: el registro del usuario debe sobrevivir al cierre de sesión'
    'alvo.08' = 'Caché de Java Web Start (Tasy)'
    'restaurarsessao.11' = 'Sesión restaurada: {0} programa(s) reabiertos, prioridades normalizadas.'
    'diagcitrix.07' = 'Sin el cliente instalado, ARIA/MOSAIQ/Monaco no abren. La instalación exige a TI.'
    'openitemgrande.02' = 'Abrir en Explorer'
    'alvo.14' = 'Informes de error de Windows'
    'topo.02' = 'Workstation Kit  ·  diagnóstico, limpieza y optimización de sesión'
    'item.16' = 'OneDrive'
    'invsoftware.05' = '[X] {0,-14} {1}'
    'item.06' = 'Superposición de NVIDIA'
    'item.41' = 'Vaciar la papelera de reciclaje (rutina)'
    'diagespaco.02' = 'Un disco casi lleno es la causa número uno de lentitud. El Módulo 3 libera espacio ahora.'
    'diagtasy.36' = 'Una oscilación de ese tamaño es lo que hace que la pantalla de Tasy se "congele" de vez en cuando. Regístrelo en el ticket con la hora.'
    'item.23' = 'Componente de Apple'
    'diagtasy.29' = '{0}: la página tarda {1} ms en responder'
    'diagcitrix.10' = 'El kit nunca cierra Citrix. Cierre la sesión antes de limpiar la caché de Workspace.'
    'planolimpeza.05' = 'Para eliminar, ejecute:  cmdkey /delete:{0}   -- compruebe el nombre antes: si es un error de DNS, el servidor existe y la credencial es válida.'
    'arquivosgrandes.03' = 'Solo el disco C:. Es el único cuyo espacio libre afecta el rendimiento de Windows.'
    'diagdownloads.05' = '{0} archivos con más de 90 días · {1}'
    'invidentificacao.02' = 'Máquina .......: {0}   usuario: {1}\{2}'
    'obs.04' = 'Cierre Teams antes. Las conversaciones quedan en el servidor.'
    'diagtasy.48' = 'Caché de CentBrowser con {0}'
    'diagpersistencia.27' = 'Perdió: '
    'modulo7sessao.01' = 'Módulo 5 - Optimizar la sesión actual'
    'diagnuvem.11' = 'Teams consumiendo {0}'
    'diagcaches.03' = 'Ejecute el Módulo 3 (Limpieza segura).'
    'obs.14' = 'Cierre Excel antes. Son archivos de bloqueo que quedaron de sesiones colgadas.'
    'modulo3analise.05' = 'Memoria devuelta ahora ........ {0}'
    'topo.24' = 'Estado de la sesión'
    'inventariodiagnost.17' = 'Lo que requiere administrador está marcado como tal: use "Copiar log" en el ticket.'
    'diagcitrix.25' = 'Puertos web: {0}'
    'logemuso.04' = '   No se puede leer la fecha del archivo. Si la aplicación no está abierta, elimine:'
    'invmigracao.01' = 'Inventario · residuos de migración de dominio'
    'nota.14' = 'Va a la papelera de reciclaje, se puede recuperar. Cambie el período en la casilla sobre la lista.'
    'diagsessaoativa.18' = 'El Módulo 7 cierra todo esto de una vez, preserva Citrix y los navegadores y devuelve la memoria al trabajo clínico. Vuelve a la normalidad en el próximo inicio de sesión.'
    'nota.01' = 'Vuelve al abrir OneDrive o en el próximo inicio de sesión.'
    'modulo3limpeza.43' = 'Memoria devuelta .............. {0}'
    'topo.11' = 'Optimizar la sesión actual'
    'diagmemoria.05' = 'Por debajo de 8 GB, Excel + navegador + Teams bloquean el equipo. Solicite una ampliación.'
    'invmigracao.16' = 'Ninguna credencial huérfana en el Administrador de credenciales.'
    'invidentificacao.05' = 'Procesador ....: {0} ({1} núcleos)'
    'arquivosgrandes.06' = 'Disco {0} de la paginación con solo {1}% libre'
    'modulo3limpeza.13' = 'Limpiando cachés y temporales...'
    'diagencerraveis.09' = '{0} recuperables cerrando programas prescindibles'
    'diagcitrix.53' = '{0,10}  {1}'
    'invsoftware.10' = 'Servidor'
    'restaurarsessao.03' = 'Prioridad normalizada en {0} programa(s).'
    'restartcomputador.06' = 'Reinicie desde el menú Inicio. Use "Reiniciar", no "Apagar".'
    'modulo3limpeza.09' = 'SESSAO'
    'topo.04' = 'Listo.'
    'diagsessaoativa.06' = '   sesión {0,-4} {1,-16} PID {2,-7} {3,10}'
    'item.12' = 'Spotify'
    'invsoftware.03' = '{0,-56} {1,-18} {2}'
    'diagcitrix.37' = 'El servidor está activo, pero el store no. Registre el código en el ticket.'
    'item.25' = 'Extra de Windows'
    'topo.09' = 'Diagnóstico e inventario'
    'diagtasy.32' = '{0}: respondiendo en {1} ms'
    'nota.11' = 'Solo oculta el acceso directo. Los archivos y la cuenta quedan intactos.'
    'item.45' = 'LibreOffice'
    'alvo.03' = 'Caché de Microsoft Edge'
    'diagtasy.44' = 'Caché de Java Web Start: {0}'
    'diagpersistencia.16' = 'Ningún marcador anterior. Dejando uno ahora en cada lugar.'
    'diagcitrix.02' = 'Citrix Receiver antiguo: versión {0}'
    'modulo3limpeza.38' = 'Con OneDrive fuera del inicio automático, los archivos solo suben a la nube cuando se abre.'
    'diagcitrix.49' = 'Proxy activado: {0}'
    'diagencerraveis.07' = '{0} de memoria retenidos en {1} programa(s) prescindible(s)'
    'diagencerraveis.02' = 'Criterio: no guardan documento abierto y vuelven solos cuando usted los abra de nuevo.'
    'diagcitrix.56' = '{0} archivos .ica sueltos'
    'modulo3limpeza.46' = 'Memoria en uso ahora .......... {0}% ({1} libres)'
    'inventariodiagnost.11' = 'Dónde está el problema, por área:'
    'diagpersistencia.08' = 'Windows no pudo abrir su perfil y creó uno descartable. No sirve de nada ajustar nada ahora. Ticket para que TI recree el perfil.'
    'topo.20' = 'Desmarcar todo'
    'topo.22' = 'Cerrar'
    'invidentificacao.10' = 'Hardware'
    'logemuso.01' = 'Señalización de app en uso encontrada pero VIEJA ({0:N0} min). Ignorada.'
    'ajusteregistro.03' = '{0} de {1} ajustes de este grupo están bloqueados por política. Los demás se aplicaron.'
    'diagcitrix.32' = '{0}: el portal responde en {1} ms'
    'otimizacoes.04' = 'Esto devuelve los efectos visuales, las aplicaciones en segundo plano y el inicio automático al estado anterior.

¿Continuar?'
    'modulo3limpeza.56' = 'Error durante la limpieza: {0}'
    'alvo.06' = 'Caché de Teams (versión clásica)'
    'modulo3limpeza.31' = 'Papelera de reciclaje vaciada: {0} liberados'
    'inventario.08' = 'No se pudo guardar el informe: {0}'
    'diaginicioacionave.26' = 'El inicio automático del usuario ya está reducido'
    'modulo7sessao.12' = 'Prioridad de CPU .............. clínico por encima de lo normal, el resto por debajo'
    'modulo3limpeza.19' = 'Deteniendo tareas programadas en ejecución...'
    'invsoftware.07' = '[X] Rol de servidor: {0} ({1} proceso(s), {2} servicio(s))'
    'diagsistemaacionav.04' = 'Reiniciar'
    'obs.10' = 'Solo los lanzadores que descargó el navegador. La caché de Citrix nunca se toca.'
    'diaginicioacionave.13' = '[ mantener ] {0,-30} {1}'
    'item.08' = 'Widgets de Windows'
    'item.43' = 'Compactar la memoria de los programas que quedan abiertos'
    'diagtasy.34' = 'Por encima de 60 ms dentro de la red interna es un mal camino. Compruebe la velocidad negociada de la tarjeta y el puerto del switch antes de culpar al servidor.'
    'item.21' = 'Componente de Adobe'
    'invecossistema.13' = '{0} aplicación(es) instalada(s) y ausente(s) del registro'
    'inventariodiagnost.03' = 'Parte 1: retrato de la máquina. Parte 2: lo que el Módulo 3 resuelve sin administrador.'
    'invmigracao.05' = 'Migración'
    'diagmemoria.03' = 'Cierre lo que no esté en uso, principalmente pestañas del navegador.'
    'topo.21' = 'APLICAR TODO'
    'arquivosgrandes.01' = 'Módulo 4 - Archivos y carpetas grandes'
    'diagpersistencia.14' = 'Filtro de escritura en el disco: {0}'
    'diagpersistencia.25' = 'Esto es un filtro de escritura o un congelador de disco, no un problema de perfil. Solo TI lo desactiva.'
    'invmigracao.10' = 'Señal de perfil heredado de otra cuenta, típico de migración de dominio. Funciona, pero confunde scripts, la política de grupo y los permisos. Si esta máquina da un problema extraño de perfil, es aquí donde se empieza.'
    'inventariodiagnost.01' = 'Módulo 2 - Inventario y diagnóstico'
    'diagmemoria.02' = 'Cierre programas y pestañas del navegador. El equipo está usando el disco como memoria.'
    'diagtasy.11' = 'Ninguna dirección de Tasy configurada'
    'diagpersistencia.01' = 'Persistencia - lo que sobrevive al cierre de sesión'
    'topo.26' = 'REVERTIR AL ORIGINAL'
    'inventariodiagnost.15' = '{0}. [{1}/{2}] {3}'
    'invmigracao.03' = 'Agente de migración de dominio aún instalado ({0})'
    'diagpersistencia.30' = 'Algún lugar no aceptó ni escribir ahora - lo que ya es una respuesta.'
    'otimizacoes.02' = 'Ninguna optimización registrada para deshacer.'
    'diagtasy.55' = 'Todo bajo y Tasy congelándose  -> es la estación de trabajo: navegador con demasiadas pestañas, caché viejo o memoria llena.'
    'diaginicioacionave.03' = 'Por eso no se puede medir el tiempo de inicio en esta sesión.'
    'modulo3limpeza.41' = 'Aplicado: {0}'
    'modulo3limpeza.26' = '{0}: {1}'
    'diagsistemaacionav.01' = 'Estado de la máquina'
    'diagdownloads.08' = 'Los 10 más grandes:'
    'invredeimpressoras.09' = '{0}{1,-42} {2}'
    'invpastaclinica.03' = 'Verifique la carpeta de trabajo configurada.'
    'arquivosgrandes.04' = 'Las carpetas de Windows y Program Files quedan fuera: no hay nada que hacer en ellas sin administrador.'
    'diagpersistencia.10' = 'Lo que no se suba al servidor a tiempo se pierde. Vale la pena confirmar con TI si el perfil móvil está guardando.'
    'diagsessaoativa.10' = 'No cierre ni compacte ahora. El kit no toca esos procesos, pero cerrar el programa por fuera interrumpe el lote - y en un lote nocturno no hay nadie para notarlo.'
    'arquivosgrandes.13' = 'LOS 10 DIRECTORIOS MÁS GRANDES'
    'alvo.20' = 'Caché de Adobe Acrobat/Reader'
    'modulo7sessao.08' = 'Leyendo los procesos de la sesión...'
    'diagnuvem.07' = 'OneDrive consumiendo {0}'
    'diaginicioacionave.22' = '{0} programas se abren solos y pueden desactivarse con seguridad'
    'invinicializacaoco.01' = 'Inventario · todo lo que se abre con Windows'
    'diagtasy.13' = 'Tipo ......: {0}'
    'diagcitrix.11' = 'Ningún proceso de Citrix en ejecución.'
    'diagpersistencia.03' = 'Copia móvil: {0}'
    'topo.16' = 'Lo que se va a aplicar'
    'item.38' = 'AppData\Local del perfil'
    'diagcitrix.35' = '{0}: portal activo, solicitando inicio de sesión (HTTP {1})'
    'modulo3limpeza.48' = 'Para volver ahora, sin reiniciar: botón "REVERTIR AL ORIGINAL".'
    'alvo.01' = 'Papelera de reciclaje'
    'invmigracao.09' = 'La carpeta del perfil no tiene el nombre del usuario: {0} usa la carpeta {1}'
    'modulo3limpeza.53' = 'Reiniciar ahora'
    'diagtasy.53' = 'Red baja y página alta        -> servidor de aplicación lento. Ticket, con estos números.'
    'diagsessaoativa.05' = 'ATENCIÓN: hay carga de radioterapia.ai en OTRA sesión ({0} proceso(s)):'
    'modulo3limpeza.50' = 'Libre en C: ahora ............. {0} ({1}%)'
    'diaginicioacionave.12' = 'Mantenidos como están:'
    'modulo3limpeza.51' = 'Todo lo que fue desactivado vuelve con el botón "Deshacer optimizaciones".'
    'diagmapeamentos.03' = '{0} -> {1}  [desconectada]'
    'diagtasy.05' = '{0} consumiendo {1} en {2} procesos'
    'invecossistema.15' = '. Es la mentira en la dirección opuesta, y oculta al soporte la versión que está en uso.'
    'diagsessaoativa.02' = 'No se pudieron mapear los procesos de la sesión.'
    'invofficeexcel.04' = 'La carpeta clínica no está en las ubicaciones de confianza de Excel'
    'topo.35' = 'Kit de Soporte listo.'
    'diagtasy.15' = 'servicio externo, sale por internet'
    'explorer.02' = 'No se pudo reiniciar el Explorer (bloqueo del antivirus).'
    'diagpersistencia.22' = '   PERDIÓ      {0,-30} pierde {1}'
    'otimizacoes.06' = 'Cancelar'
    'invecossistema.08' = '   {0,-24} v{1,-10} {2,-12} {3}'
    'diagtasy.22' = 'Midiendo el tiempo de respuesta (4 muestras)...'
    'alvo.04' = 'Caché de Google Chrome'
    'modulo3limpeza.54' = 'Reiniciar después'
    'modulo3limpeza.33' = '{0} elemento(s) ya habían sido eliminados por otra etapa.'
    'diagcitrix.59' = 'Si todos fallan, es la red o el perfil de la estación de trabajo.'
    'diagajustespendent.02' = 'Todos los ajustes de rendimiento ya están aplicados'
    'inventario.01' = 'Informe'
    'alvo.18' = 'Versiones antiguas de Teams'
    'diagtasy.21' = 'Si las otras estaciones de trabajo tampoco abren, el servidor se cayó: ticket inmediato.'
    'invinicializacaoco.05' = 'Estos solo los desactiva TI. Regístrelo en el ticket si el inicio de sesión es largo.'
    'invsoftware.09' = 'La limpieza y el reinicio requieren un acuerdo previo con quien depende del servicio.'
    'modulo7sessao.07' = 'antes de optimizar una estación de trabajo que sirve a otras personas.'
    'modulo3analise.07' = 'Ajustes de rendimiento ........ {0}'
    'restartcomputador.05' = 'Windows no aceptó la solicitud de reinicio (código {0}). La máquina NO se va a reiniciar.'
    'diagsistemaacionav.06' = 'Inicio rápido activado: "Apagar" no limpia la memoria, solo "Reiniciar".'
    'arquivosgrandes.14' = 'Midiendo carpeta {0}/{1}: {2}'
    'diaginicioacionave.11' = '[ desactivar ] {0,-30} {1}'
    'diagtasy.42' = '{0} responde rápido y {1} no, desde la misma estación de trabajo y en el mismo minuto.'
    'diaginicioacionave.08' = 'Inicialización en {0} segundos'
    'invpastaclinica.06' = '{0,-40} {1,10}'
    'diagdownloads.12' = 'Downloads con {0}'
    'diagmapeamentos.02' = 'Ninguna unidad de red desconectada'
    'topo.14' = 'Restaurar sesión (reabrir)'
    'diagcitrix.08' = 'Citrix en ejecución: {0} proceso(s), {1} - {2}'
    'modulo3limpeza.16' = 'Quitando programas del inicio automático...'
    'diagnuvem.05' = 'Active "Archivos a petición" (haga clic en el icono de la nube > engranaje > Configuración) para liberar espacio sin perder el acceso.'
    'diagtasy.01' = 'Diagnóstico de Tasy (historia clínica electrónica)'
    'item.19' = 'Mensajería'
    'topo.43' = 'Aplicación autocontenida: la rutina de preparación viene dentro de ella.'
    'diagcitrix.14' = 'Dirección no válida en la configuración: {0}'
    'diagespaco.07' = 'Carpeta del sistema: solo TI la limpia, y su rutina de limpieza ya la cubre. Registre el tamaño en el ticket.'
    'invecossistema.03' = 'Registro ..: {0}'
    'diagcitrix.29' = 'sin respuesta'
    'diagdownloads.03' = 'Carpeta Downloads vacía'
    'diagmemoria.04' = 'Memoria total baja para el uso clínico: {0}'
    'diagtasy.26' = 'Página (HTTP) .: min {0} · típico {1} · máx {2} ms   (HTTP {3})'
    'diaginicioacionave.01' = 'Inicio automático de Windows'
    'diagnuvem.01' = 'OneDrive y Teams'
    'diagencerraveis.05' = 'Nada que cerrar: lo que está abierto es trabajo o sistema.'
    'invidentificacao.09' = 'Por debajo de 8 GB, Excel en red más el navegador ya saturan la máquina. Regístrelo en el ticket.'
    'obs.09' = 'Puede pedir iniciar sesión otra vez en sitios internos antiguos. Por eso viene sin marcar.'
    'restartcomputador.02' = 'Warning'
    'modulo7sessao.06' = 'Los servicios se ejecutan como SISTEMA y no entran en la lista, pero confirme con quien depende de ellos'
    'diagdownloads.10' = 'Downloads ocupando {0}, de los cuales {1} tiene más de 90 días'
    'invinicializacaoco.03' = '[{0,-13}] {1,-36} {2}'
    'ajusteregistro.02' = '{0}: no se pudo ajustar.'
    'diagcitrix.13' = 'Ningún store guardado en el perfil del usuario (acceso solo por el navegador).'
    'item.36' = 'Desactivar las sugerencias y apps promocionadas de Windows'
    'invecossistema.11' = 'El registro lo declara pero el disco no lo tiene: '
    'diagsistemaacionav.02' = '{0} · build {1} · {2} de memoria'
    'topo.33' = 'No se pudo guardar el archivo.'
    'diagnuvem.08' = 'Use el botón "Cerrar OneDrive ahora" cuando necesite rendimiento. Vuelve al abrirlo de nuevo.'
    'alvo.24' = 'Lista de archivos recientes'
    'diagcitrix.45' = 'Cambie el acceso directo a {0} . En HTTP el navegador bloquea parte del inicio de sesión único y el tráfico va sin cifrado.'
    'modulo3limpeza.34' = '{0} elemento(s) en uso se mantuvieron. Ejemplos:'
    'invredeimpressoras.02' = '{0,-32} {1,-12} {2}'
    'diagmemoria.09' = 'Cierre las pestañas que no están en uso. Cada pestaña abierta es un proceso con memoria propia.'
    'arquivosgrandes.15' = 'PUEDE BORRAR .... {0} en {1} elemento(s) - descartable, Windows lo recrea o ya no lo usa'
    'diagmapeamentos.01' = 'Unidades de red'
    'diagtasy.07' = 'Tasy'
    'diaginicioacionave.17' = 'impresora: {0}'
    'item.30' = 'no reconocido'
    'modulo3limpeza.06' = 'YesNo'
    'arquivosgrandes.17' = 'NO BORRAR ....... dato clínico, base de datos, navegador de la historia clínica o archivo de sistema'
    'diagtasy.17' = '{0}: el nombre {1} no se resuelve en esta red'
    'invofficeexcel.06' = 'Excel'
    'diaginicioacionave.15' = '{0} unidad(es) de red y {1} impresora(s) de red se reconectan en cada inicio de sesión.'
    'diagtasy.49' = 'El Módulo 3 tiene el elemento, pero viene DESMARCADO: es el navegador de la historia clínica y la decisión de limpiar es suya.'
    'inventariodiagnost.06' = 'Fallo en el paso {0}: {1}'
    'otimizacoes.08' = 'Reactivado en el inicio automático: {0}'
    'topo.10' = 'Archivos y carpetas grandes'
    'nota.09' = 'Hace que el inicio de sesión único funcione y que el .ica se abra solo.'
    'diagpersistencia.19' = 'Cierre la sesión (o reinicie) y ejecute de nuevo.'
    'diagtasy.30' = 'La red responde en {0} ms, así que la demora es del servidor de aplicación. Adjunte este log al ticket: no sirve de nada tocar la estación de trabajo.'
    'alvo.21' = 'Registros y diagnósticos de Office'
    'diagsistemaacionav.05' = 'Reinicie al final de la limpieza (botón en el menú lateral).'
    'openitemgrande.04' = 'Abierto en Explorer: {0}'
    'diagtasy.45' = 'Caché de Java Web Start con {0}'
    'obs.13' = 'Desaparece la lista de recientes de Office y del Explorer. No se borra ningún archivo.'
    'diagcitrix.03' = 'Receiver 4.x dejó de tener soporte en 2018 y se reemplaza por Citrix Workspace. Actualizar requiere el área de TI, pero resuelve buena parte de las fallas de apertura y de inicio de sesión único.'
    'invecossistema.01' = 'Aplicaciones de radioterapia.ai'
    'acaomemoria.03' = 'La primera acción dentro de cada programa puede tardar un instante: recarga lo que necesita.'
    'planolimpeza.01' = 'Midiendo la Papelera de reciclaje...'
    'arquivosgrandes.12' = '{0,2}. [{1,-15}] {2,10}  {3}'
    'modulo3analise.09' = '{0} elemento(s) de inicio automático quedaron sin marcar por no ser reconocidos:'
    'item.32' = 'Impedir que las apps de la Store se ejecuten en segundo plano'
    'topo.12' = 'Qué sobrevive al cerrar sesión'
    'invservicos.02' = '{0} en ejecución de {1} instalados'
    'diagdownloads.02' = 'Carpeta Downloads no encontrada.'
    'diagsessaoativa.03' = 'Medido en la sesión {0}, de {1} sesión(es) activa(s) en esta máquina.'
    'arquivosgrandes.18' = 'Ningún directorio recibe "SE PUEDE BORRAR": una carpeta entera siempre exige comprobación.'
    'invpastaclinica.04' = 'Entorno'
    'inventario.07' = 'Junte los .csv de varias máquinas en una sola hoja de cálculo para comparar el parque de equipos.'
    'diaginicioacionave.10' = 'Es seguro NO cargar de nuevo (el programa sigue funcionando al abrirse):'
    'invpastaclinica.01' = 'Inventario · carpeta clínica'
    'alvo.17' = 'Logs de OneDrive'
    'invecossistema.16' = 'El registro del ecosistema coincide con el disco'
    'diagsessaoativa.12' = '[{0}] {1,-30} {2,10}  ({3} proceso(s))'
    'diagsessaoativa.08' = '   {0,-16} PID {1,-7} {2,10}   {3}'
    'diagtasy.03' = 'Sin cliente instalado: Tasy se ejecuta en el navegador (Wheb HTML5).'
    'diagcitrix.16' = 'Aplicaciones: {0}'
    'modulo3limpeza.23' = 'Aplicando {0}/{1}: {2}'
    'invofficeexcel.02' = 'Versión de Office: {0}'
    'diagespaco.05' = 'No se pudieron leer las unidades.'
    'diaginicioacionave.18' = '{0} unidad(es) de red sin respuesta ahora: {1}'
    'diagtasy.24' = 'El servidor web está activo y la aplicación no. Ticket con la hora exacta.'
    'diaginicioacionave.07' = 'El Módulo 3 reduce esto desactivando los elementos seguros.'
    'diagcitrix.42' = 'Es lo que hace que el navegador pida la contraseña de nuevo y no abra el archivo .ica por sí solo. El Módulo 3 lo corrige con un clic.'
    'diagcitrix.20' = 'DNS .......: {0} -> {1} ({2} ms)'
    'diagnuvem.03' = 'Carpeta de OneDrive: {0} · {1}'
    'obs.07' = 'Miniaturas de pantalla de las sesiones remotas. Se recrean en la siguiente conexión.'
    'invredeimpressoras.05' = 'Red'
    'inventariodiagnost.18' = 'Este diagnóstico ya incluyó las pruebas de Citrix (tres destinos) y de Tasy (tres entornos).'
    'modulo7sessao.04' = 'El audio y el micrófono se cierran: el sonido sigue funcionando, solo sale la capa de realce.'
    'alvo.12' = 'Cookies heredadas de Windows (INetCookies)'
    'modulo3analise.06' = 'Fuera del inicio automático ... {0}'
    'modulo3limpeza.03' = 'Información'
    'diagmapeamentos.04' = '{0} unidad(es) de red desconectada(s)'
    'diagcitrix.26' = 'ninguno respondió'
    'item.34' = 'Despejar la barra de tareas (búsqueda, vista de tareas, widgets)'
    'inventario.02' = 'Servidor de logs fuera de alcance: {0}'
    'diagcitrix.39' = 'Detalle: '
    'restartcomputador.03' = 'Reiniciando el equipo...'
    'snapshotsvc.01' = 'No se pudo leer la lista de servicios de esta máquina: la verificación de rol de servidor NO se realizó.'
    'nota.02' = 'Las conversaciones quedan en el servidor. Vuelve al abrirlo.'
    'modulo3limpeza.24' = '{0}: {1} proceso(s) cerrado(s), {2} devueltos'
    'diaginicioacionave.16' = '{0} -> {1}   {2}'
    'invidentificacao.06' = 'Memoria .......: {0}'
    'modulo3limpeza.36' = 'Revise la Papelera de reciclaje antes de vaciarla de nuevo, en caso de que quiera recuperar algo.'
    'diagcitrix.47' = '{0}: certificado HTTPS no confiable en esta estación de trabajo'
    'diagpersistencia.23' = 'Todo sobrevivió al último cierre de sesión'
    'modulo3limpeza.12' = 'Cerrando programas prescindibles...'
    'diagmemoria.06' = 'Fallo al medir la memoria: {0}'
    'restaurarsessao.05' = 'Registro del Módulo 5: {0} programa(s) cerrado(s).'
    'modulo3limpeza.04' = 'Ningún elemento marcado.'
    'diaginicioacionave.25' = 'El Módulo 3 lo desactiva.'
    'diagsessaoativa.15' = 'Tareas programadas en ejecución ahora: {0}'
    'diagtasy.19' = 'Puertos ...: {0}'
    'diagcitrix.23' = '{0}: el nombre {1} no se resuelve en esta red, ni con el dominio completo'
    'item.01' = 'OneDrive (sincronización)'
    'inventariodiagnost.08' = 'Nada que corregir por aquí: la máquina está limpia del lado del usuario.'
    'diagnuvem.13' = 'Teams no está en ejecución.'
    'nota.12' = 'Deja de instalar aplicaciones por su cuenta.'
    'modulo3limpeza.47' = 'La sesión actual está más ligera. El panel de la derecha muestra el estado.'
    'diagnuvem.12' = 'Use "Cerrar Teams ahora" durante el trabajo pesado en las hojas de cálculo. Vuelve al abrirlo de nuevo.'
    'nota.07' = 'Quita decenas de procesos del segundo plano.'
    'diagdownloads.07' = 'sin ext'
    'diagcitrix.01' = 'Diagnóstico de Citrix (ARIA, MOSAIQ y Monaco)'
    'alvo.23' = 'Caché de miniaturas e iconos'
    'invmigracao.12' = 'Señal de perfil recreado en la migración. Recrear de nuevo exige administrador.'
    'planolimpeza.02' = 'Asignación de red sin respuesta: {0} apunta a {1}'
    'obs.03' = 'No toca el perfil con favoritos y contraseñas, que está en Roaming.'
    'topo.31' = 'No se pudo copiar. Seleccione el texto y use Ctrl+C.'
    'diagcitrix.55' = 'Caché de Citrix con {0} - este kit nunca la limpia'
    'diagdownloads.06' = '{0,-10} {1,10}  ({2} archivo(s))'
    'obs.11' = 'Cierre todo Office antes. Los cambios aún no enviados pueden perderse.'
    'item.20' = 'Tienda de juegos'
    'diagsessaoativa.16' = '- {0}{1}'
    'diagpersistencia.05' = 'Nada de lo que el Módulo 3 ajusta en el registro sobrevive, y Deshacer nunca encuentra qué deshacer. Solo TI cambia esto - lleve este log al ticket de soporte.'
    'invofficeexcel.09' = 'Archivo > Opciones > Complementos > Complementos COM > desmarque. Retrasa la apertura de toda hoja de cálculo.'
    'momentosessao.01' = 'Este kit es una aplicación de Windows: prepara la máquina para que el trabajo empiece.'
    'diagcitrix.44' = '{0}: el acceso directo usa HTTP, pero el servidor también atiende en HTTPS'
    'modulo3limpeza.30' = 'Papelera de reciclaje: ya estaba vacía'
    'modulo3analise.08' = 'Asignaciones muertas quitadas . {0}'
    'invecossistema.07' = '{0} aplicación(es) declarada(s):'
    'diagtasy.12' = 'Midiendo '
    'diagpersistencia.12' = 'Si el cierre de sesión no termina de sincronizar, el ajuste se pierde. Cierre los programas antes de salir.'
    'alvo.09' = 'Caché de CentBrowser (navegador de la historia clínica)'
    'diagtasy.10' = 'Java de 32 bits no pasa de ~1,5 GB de memoria. Un informe grande se congela o se cierra solo. Pida al área de TI el Java de 64 bits.'
    'inventario.05' = 'Planilla ..: {0}'
    'diagtasy.33' = '{0}: latencia típica de {1} ms para un servidor interno'
    'diagcitrix.04' = 'Citrix'
    'diagajustespendent.04' = '[ aplicar ] {0}'
    'acaoprioridade.01' = 'Prioridad ajustada: {0} programa(s) clínico(s) por encima de lo normal, {1} por debajo de lo normal.'
    'modulo7sessao.14' = '{0} programa(s) con ventana abierta quedaron desmarcados por no ser reconocidos:'
    'alvo.15' = 'Volcados de memoria de bloqueos'
}

function T {
    param([string]$K)
    if (-not $K) { return '' }
    try {
        $tab = $script:Textos[$script:Idioma]
        if ($tab -and $tab.ContainsKey($K) -and $tab[$K]) { return [string]$tab[$K] }
        $en = $script:Textos['en']
        if ($en -and $en.ContainsKey($K) -and $en[$K]) { return [string]$en[$K] }
        $pt = $script:Textos['pt']
        if ($pt -and $pt.ContainsKey($K) -and $pt[$K]) { return [string]$pt[$K] }
    } catch { }
    return ('!!' + $K + '!!')
}

function Get-IdiomaSalvo {
    try {
        $v = (Get-ItemProperty -Path ('HKCU:\Software\' + $script:MarcaApp) -Name 'Idioma' -ErrorAction Stop).Idioma
        if ($v -and (@($script:Idiomas | ForEach-Object { $_.Cod }) -contains $v)) { return [string]$v }
    } catch { }
    return $null
}

function Save-Idioma {
    param([string]$Cod)
    try {
        $k = 'HKCU:\Software\' + $script:MarcaApp
        if (-not (Test-Path $k)) { New-Item -Path $k -Force -ErrorAction Stop | Out-Null }
        New-ItemProperty -Path $k -Name 'Idioma' -Value $Cod -PropertyType String -Force -ErrorAction Stop | Out-Null
    } catch { }
}

function Get-IdiomaInicial {
    $s = Get-IdiomaSalvo
    if ($s) { return $s }
    $c = ''
    try { $c = [System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName } catch { }
    if ($c -eq 'pt') { return 'pt' }
    if ($c -eq 'es') { return 'es' }
    return 'en'
}

$script:MarcaApp = 'WorkstationKit'
$script:Idioma   = Get-IdiomaInicial

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
    Write-Log ((T 'titulo.01') -f $Texto.ToUpper(), ('=' * [Math]::Max(4, 66 - $Texto.Length))) 'TITULO'
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
    if ($Recomendacao) { Write-Log ((T 'achado.01') + $Recomendacao) 'ACAO' }
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
    foreach ($b in @($script:BtnMod2, $script:BtnMod3,
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

function Test-ArquivoEmUso {
    param([string]$Caminho)
    if (-not (Test-Path -LiteralPath $Caminho)) { return $false }
    try {
        $fs = [System.IO.File]::Open($Caminho, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch { return $true }
}

$script:Protegidos = 'wfica32|wfcrun32|CDViewer|SelfService|Receiver|concentr|CtxWebHelper|AuthManSvr|redirector|HdxRtcEngine|CtxCFRUI|Citrix|tasy|TasyAgent|CentBrowser|javaw|^java$|jp2launcher|Wheb|Philips|CcmExec|ntrtscan|tmlisten|TMBM|PccNTMon|ShowMsg|smartscreen|unsecapp|SearchProtocolHost|SearchFilterHost|DSASvc|QualysAgent|stAgent|Cortex|cyserver|cytray|cyvera|traps|LsAgent|Quest|OnDemand|ODMActiveDirectory|SecureConnector|ARIA|Eclipse|Varian|Vitrea|MIM|MOSAIQ|RayStation|Monaco|Velocity|Osirix|Horos|Weasis|dicom|PACS|EXCEL|WINWORD|POWERPNT|OUTLOOK|MSACCESS|onenote|StickyNot|msedge|chrome|firefox|iexplore|notepad|wordpad|Acrobat|AcroRd32|AnyConnect|GlobalProtect|FortiClient|Pulse|CcmExec|CmRcService|ccmsetup|CSFalcon|CSAgent|Sophos|SAVService|^mfe|masvc|macmnsvc|McShield|ccSvcHst|SepMaster|ZSA|stAgent|nsdiag|ivanti|LANDesk|Forcepoint|splunk|nxlog|MsMpEng|NisSrv|SecurityHealth|Acronis|^mms$|Veeam|CommVault|^cvd$|^CvMountd$|TrueImage|Macrium|^Reflect|ShadowProtect|Arcserve|Datto|Carbonite|IDriveService|Realtek|RtkAud|IDTNC|Synaptics|igfx|nvcontainer|audiodg|System|Idle|Registry|smss|csrss|wininit|winlogon|^services$|lsass|svchost|fontdrvhost|dwm|explorer|RuntimeBroker|sihost|ctfmon|taskhostw|dllhost|conhost|WmiPrvSE|powershell|pwsh|LogonUI|SearchIndexer'

function Get-CatalogoProcessos {
    @(
        [pscustomobject]@{ Padrao = '^OneDrive$|^FileCoAuth$';                                    Rotulo = (T 'item.01');        Nivel = 'Sempre';   Nota = (T 'nota.01') }
        [pscustomobject]@{ Padrao = '^Teams$|^ms-teams$|^msteams';                                 Rotulo = (T 'item.02');               Nivel = 'Sempre';   Nota = (T 'nota.02') }
        [pscustomobject]@{ Padrao = 'Update$|Updater$|^GoogleUpdate|EdgeUpdate|^AdobeARM$|^AdobeGCClient$|^armsvc$|^jusched$|^SquirrelUpdate|^OfficeC2RClient$'; Rotulo = (T 'item.03'); Nivel = 'Sempre'; Nota = (T 'nota.03') }
        [pscustomobject]@{ Padrao = '^CCXProcess$|^Creative Cloud|^CCLibrary|^AdobeIPCBroker$|^AdobeNotificationClient$'; Rotulo = (T 'item.04'); Nivel = 'Sempre'; Nota = '' }
        [pscustomobject]@{ Padrao = '^GameBar|^XboxApp|^GamingServices|^XboxGame|^GameBarPresence'; Rotulo = (T 'item.05');              Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^NVIDIA Share$|^NVIDIA Web Helper';                           Rotulo = (T 'item.06');           Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^YourPhone$|^PhoneExperienceHost$';                           Rotulo = (T 'item.07');          Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^Widgets$|^WidgetService$';                                   Rotulo = (T 'item.08');            Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^SearchApp$|^Cortana$|^Copilot';                              Rotulo = (T 'item.09');            Nivel = 'Sempre';   Nota = (T 'nota.04') }
        [pscustomobject]@{ Padrao = '^SupportAssist|^DellSupport|^HPSupport|^HPPrintScan|^Vantage|^ImController$|^McUICnt$|^HPNotifications'; Rotulo = (T 'item.10'); Nivel = 'Sempre'; Nota = '' }
        [pscustomobject]@{ Padrao = '^iTunesHelper$|^AppleMobileDevice|^iPodService$';             Rotulo = (T 'item.11');                Nivel = 'Sempre';   Nota = '' }
        [pscustomobject]@{ Padrao = '^Spotify';                                                    Rotulo = (T 'item.12');                       Nivel = 'SemJanela'; Nota = '' }
        [pscustomobject]@{ Padrao = '^Dropbox$|^Box$|^GoogleDriveFS$';                             Rotulo = (T 'item.13');                 Nivel = 'SemJanela'; Nota = '' }
        [pscustomobject]@{ Padrao = '^Zoom$|^Slack$|^Discord$|^Skype';                             Rotulo = (T 'item.14');                 Nivel = 'SemJanela'; Nota = (T 'nota.05') }
        [pscustomobject]@{ Padrao = '^Steam|^EpicGames';                                           Rotulo = (T 'item.15');                Nivel = 'SemJanela'; Nota = '' }
    )
}

function Get-ProcessosEncerraveis {
    $saida = @()
    try { $todos = @(Get-Process -ErrorAction SilentlyContinue) } catch { return $saida }

    foreach ($c in (Get-CatalogoProcessos)) {
        $ps = @($todos | Where-Object { (Test-Padrao $_.ProcessName $c.Padrao) -and -not (Test-Padrao $_.ProcessName $script:Protegidos) })
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
        [pscustomobject]@{ Padrao = 'OneDrive';                          Rotulo = (T 'item.16') }
        [pscustomobject]@{ Padrao = 'Teams';                             Rotulo = (T 'item.02') }
        [pscustomobject]@{ Padrao = 'Spotify';                           Rotulo = (T 'item.12') }
        [pscustomobject]@{ Padrao = 'Skype';                             Rotulo = (T 'item.17') }
        [pscustomobject]@{ Padrao = 'Zoom';                              Rotulo = (T 'item.18') }
        [pscustomobject]@{ Padrao = 'Slack|Discord';                     Rotulo = (T 'item.19') }
        [pscustomobject]@{ Padrao = 'Steam|Epic';                        Rotulo = (T 'item.20') }
        [pscustomobject]@{ Padrao = 'Dropbox|Box Sync|GoogleDrive';      Rotulo = (T 'item.13') }
        [pscustomobject]@{ Padrao = 'Adobe|CCX|Creative|Acro';           Rotulo = (T 'item.21') }
        [pscustomobject]@{ Padrao = 'Update|Updater|jusched|Java';       Rotulo = (T 'item.22') }
        [pscustomobject]@{ Padrao = 'iTunes|Apple|QuickTime';            Rotulo = (T 'item.23') }
        [pscustomobject]@{ Padrao = 'SupportAssist|Dell|HP |Vantage|Lenovo'; Rotulo = (T 'item.24') }
        [pscustomobject]@{ Padrao = 'Cortana|Copilot|Widgets|YourPhone'; Rotulo = (T 'item.25') }
        [pscustomobject]@{ Padrao = 'Steam|Xbox|GameBar|NVIDIA';         Rotulo = (T 'item.26') }
        [pscustomobject]@{ Padrao = 'Spark|Grammarly|CCleaner|Toolbar';  Rotulo = (T 'item.27') }
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
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Seguro'; Rotulo = (T 'item.28'); Seguro = $true }
            continue
        }
        if (Test-Padrao $texto $script:Protegidos) {
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Protegido'; Rotulo = (T 'item.29'); Seguro = $false }
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
            $saida += [pscustomobject]@{ Nome = $i.Nome; Tipo = $i.Tipo; Classe = 'Desconhecido'; Rotulo = (T 'item.30'); Seguro = $false }
        }
    }
    return $saida
}

function Get-AjustesPendentes {
    $lista = @()

    $vfx = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting'
    if ($vfx -ne 2) {
        $lista += [pscustomobject]@{ Id = 'EFEITOS'; Rotulo = (T 'item.31'); Nota = (T 'nota.06') }
    }
    $bg = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' 'GlobalUserDisabled'
    if ($bg -ne 1) {
        $lista += [pscustomobject]@{ Id = 'SEGUNDOPLANO'; Rotulo = (T 'item.32'); Nota = (T 'nota.07') }
    }
    $ss = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' '01'
    if ($ss -ne 1) {
        $lista += [pscustomobject]@{ Id = 'STORAGE'; Rotulo = (T 'item.33'); Nota = (T 'nota.08') }
    }
    $zonas = Get-ZonasCitrixPendentes
    if ($zonas.Count -gt 0) {
        $lista += [pscustomobject]@{ Id = 'CITRIXZONA'; Rotulo = ('Confiar nos enderecos do Citrix ({0})' -f ($zonas -join ', ')); Nota = (T 'nota.09') }
    }

    $sb = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' 'SearchboxTaskbarMode'
    $tv = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowTaskViewButton'
    $wd = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarDa'
    if ($sb -ne 0 -or $tv -ne 0 -or $wd -ne 0) {
        $lista += [pscustomobject]@{ Id = 'BARRATAREFAS'; Rotulo = (T 'item.34'); Nota = (T 'nota.10') }
    }

    $odIco = Get-ValorReg 'HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' 'System.IsPinnedToNameSpaceTree'
    if ($odIco -ne 0) {
        $lista += [pscustomobject]@{ Id = 'ONEDRIVEICONE'; Rotulo = (T 'item.35'); Nota = (T 'nota.11') }
    }

    $sug = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' 'SilentInstalledAppsEnabled'
    if ($sug -ne 0) {
        $lista += [pscustomobject]@{ Id = 'SUGESTOES'; Rotulo = (T 'item.36'); Nota = (T 'nota.12') }
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

    foreach ($n in @('SelfService', 'SelfServicePlugin', 'Receiver', 'wfica32', 'wfcrun32', 'concentr', 'CDViewer')) {
        try {
            $p = Get-Process -Name $n -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($p -and $p.Path -and (Test-Path -LiteralPath $p.Path)) {
                $v = (Get-Item -LiteralPath $p.Path).VersionInfo
                return [pscustomobject]@{ Instalado = $true; Versao = "$($v.ProductVersion)"; Caminho = $p.Path }
            }
        } catch { }
    }

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
    if ($Servidor -notmatch '\.') { return $true }
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
            if ($partes.Count -lt 2) { return @() }
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
        Write-Log ((T 'ajustezonascitrix.01') -f $srv) 'DADO'
    }
    if ($criadas.Count -gt 0) {
        $e = Read-Estado
        $ant = @()
        if ($e.ContainsKey('zonas_citrix')) { $ant = @($e['zonas_citrix']) }
        Save-Estado 'zonas_citrix' ($ant + $criadas)
    }
}

function Invoke-DiagCitrix {
    Write-Titulo (T 'diagcitrix.01')

    $cli = Get-CitrixInstalado
    if ($cli.Instalado) {
        $maior = 0
        try { $maior = [int](("$($cli.Versao)" -split '\.')[0]) } catch { }
        if ($maior -gt 0 -and $maior -lt 19) {
            Add-Achado 'ALERTA' ((T 'diagcitrix.02') -f $cli.Versao) (T 'diagcitrix.03') (T 'diagcitrix.04') 'Alto'
        } else {
            Add-Achado 'OK' ((T 'diagcitrix.05') -f $cli.Versao) '' (T 'diagcitrix.04')
        }
    } else {
        Add-Achado 'CRITICO' (T 'diagcitrix.06') (T 'diagcitrix.07') (T 'diagcitrix.04') 'Alto'
    }

    $pr = Get-CitrixProcessos
    if ($pr.Rodando) {
        Write-Log ((T 'diagcitrix.08') -f $pr.Qtd, (Format-Bytes $pr.Memoria), ($pr.Nomes -join ', ')) 'DADO'
        if ($pr.Sessao) {
            Add-Achado 'ALERTA' (T 'diagcitrix.09') (T 'diagcitrix.10') (T 'diagcitrix.04') 'Medio'
        }
    } else {
        Write-Log (T 'diagcitrix.11') 'DADO'
    }

    $conf = Get-CitrixStoresConfigurados
    if ($conf.Count -gt 0) {
        Write-Log (T 'diagcitrix.12') 'DADO'
        foreach ($c in $conf) { Write-Log ('- ' + $c) 'DADO' }
    } else {
        Write-Log (T 'diagcitrix.13') 'DADO'
    }

    foreach ($s in $script:CitrixStores) {
        if ($script:Cancelar) { return }
        $i = Get-InfoUrl -Url $s.Url
        if (-not $i) { Write-Log ((T 'diagcitrix.14') -f $s.Url) 'ALERTA'; continue }

        Write-Log '' 'DADO'
        Write-Log ((T 'diagcitrix.15') -f $s.Nome) 'DADO'
        Write-Log ((T 'diagcitrix.16') -f $s.Apps) 'DADO'
        Write-Log ((T 'diagcitrix.17') -f $s.Url) 'DADO'
        Set-Status ((T 'diagcitrix.18') + $s.Nome + (T 'diagcitrix.19'))

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
                    Write-Log ((T 'diagcitrix.20') -f $c, ($ips -join ', '), $rel.ElapsedMilliseconds) 'DADO'
                    $alvo = $c
                    $resolvido = $true
                    if ($c -ne $i.Servidor) {
                        $urlTeste = $s.Url.Replace($i.Servidor, $c)
                        Add-Achado 'ALERTA' ((T 'diagcitrix.21') -f $s.Nome, $i.Servidor, $c) ((T 'diagcitrix.22') -f $urlTeste) (T 'diagcitrix.04') 'Alto'
                    }
                    break
                } catch { }
            }
            if (-not $resolvido) {
                Add-Achado 'CRITICO' ((T 'diagcitrix.23') -f $s.Nome, $i.Servidor) (T 'diagcitrix.24') (T 'diagcitrix.04') 'Alto'
                continue
            }
        }

        $portas = @(80, 443)
        $abertas = @()
        foreach ($p in $portas) {
            $t = Test-PortaTcp -Alvo $alvo -Porta $p -TimeoutMs 1500
            if ($t.Ok) { $abertas += ('{0} ({1} ms)' -f $p, $t.Ms) }
        }

        $ica = Test-PortaTcp -Alvo $alvo -Porta 1494 -TimeoutMs 1200
        $rel2598 = Test-PortaTcp -Alvo $alvo -Porta 2598 -TimeoutMs 1200
        Write-Log ((T 'diagcitrix.25') -f $(if ($abertas.Count -gt 0) { $abertas -join ' · ' } else { (T 'diagcitrix.26') })) 'DADO'
        $notaIca = $(if ($ica.Ok -or $rel2598.Ok) { '' } else { '   (normal: quem atende o ICA e o servidor da aplicacao, nao o portal)' })
        Write-Log ((T 'diagcitrix.27') -f $(if ($ica.Ok) { (T 'diagcitrix.28') } else { (T 'diagcitrix.29') }), $(if ($rel2598.Ok) { (T 'diagcitrix.28') } else { (T 'diagcitrix.29') }), $notaIca) 'DADO'

        if ($abertas.Count -eq 0) {
            Add-Achado 'CRITICO' ((T 'diagcitrix.30') -f $s.Nome, $alvo) (T 'diagcitrix.31') (T 'diagcitrix.04') 'Alto'
            continue
        }

        $r = Test-UrlHttp -Url $urlTeste
        if ($r.Codigo -ge 200 -and $r.Codigo -lt 400) {
            if ($r.Ms -gt $script:Lim.CitrixMs) {
                Add-Achado 'ALERTA' ((T 'diagcitrix.32') -f $s.Nome, $r.Ms) (T 'diagcitrix.33') (T 'diagcitrix.04') 'Medio'
            } else {
                Add-Achado 'OK' ((T 'diagcitrix.34') -f $s.Nome, $r.Codigo, $r.Ms) '' (T 'diagcitrix.04')
            }
        } elseif ($r.Codigo -eq 401 -or $r.Codigo -eq 403) {
            Add-Achado 'OK' ((T 'diagcitrix.35') -f $s.Nome, $r.Codigo) '' (T 'diagcitrix.04')
        } elseif ($r.Codigo -gt 0) {
            Add-Achado 'CRITICO' ((T 'diagcitrix.36') -f $s.Nome, $r.Codigo) (T 'diagcitrix.37') (T 'diagcitrix.04') 'Alto'
        } else {
            Add-Achado 'CRITICO' ((T 'diagcitrix.38') -f $s.Nome) ((T 'diagcitrix.39') + $r.Erro) (T 'diagcitrix.04') 'Alto'
        }

        if (Test-ZonaIntranet -Servidor $i.Servidor) {
            Write-Log (T 'diagcitrix.40') 'DADO'
        } else {
            Add-Achado 'ALERTA' ((T 'diagcitrix.41') -f $s.Nome, $i.Servidor) (T 'diagcitrix.42') (T 'diagcitrix.04') 'Alto'
        }

        if (-not $i.Https) {
            if ("$($r.UrlFinal)" -match '^https://') {
                Write-Log (T 'diagcitrix.43') 'DADO'
            } else {
                $alt = Test-HttpsAlternativo -Url $urlTeste
                if ($alt -and $alt.Disponivel) {
                    Add-Achado 'ALERTA' ((T 'diagcitrix.44') -f $s.Nome) ((T 'diagcitrix.45') -f $alt.Url) (T 'diagcitrix.04') 'Medio'
                } else {
                    Write-Log (T 'diagcitrix.46') 'DADO'
                }
            }
        }
        if ($r.CertificadoInvalido) {
            Add-Achado 'ALERTA' ((T 'diagcitrix.47') -f $s.Nome) (T 'diagcitrix.48') (T 'diagcitrix.04') 'Alto'
        }
    }

    $proxyOn = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyEnable'
    if ($proxyOn -eq 1) {
        $srv = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyServer'
        $exc = Get-ValorReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyOverride'
        Write-Log '' 'DADO'
        Write-Log ((T 'diagcitrix.49') -f $srv) 'DADO'
        Write-Log ((T 'diagcitrix.50') -f $exc) 'DADO'
        $faltando = @()
        foreach ($s in $script:CitrixStores) {
            $i = Get-InfoUrl -Url $s.Url
            if ($i -and ("$exc" -notlike ('*' + $i.Servidor + '*'))) { $faltando += $i.Servidor }
        }
        if ($faltando.Count -gt 0) {
            Add-Achado 'ALERTA' ((T 'diagcitrix.51') -f ($faltando -join ', ')) (T 'diagcitrix.52') (T 'diagcitrix.04') 'Medio'
        }
    }

    $caches = Get-CitrixCaches
    $tot = 0.0
    foreach ($c in $caches) {
        $b = Get-TamanhoPasta -Caminho $c -TimeoutSeg 45
        if ($b -gt 0) { $tot += $b; Write-Log ((T 'diagcitrix.53') -f (Format-Bytes $b), $c) 'DADO' }
    }
    $icas = @()
    foreach ($p in @($env:TEMP, (Join-Path $env:USERPROFILE 'Downloads'))) {
        try { $icas += @(Get-ChildItem -LiteralPath $p -Filter '*.ica' -File -Force -ErrorAction SilentlyContinue) } catch { }
    }
    if ($icas.Count -gt 0) { Write-Log ((T 'diagcitrix.54') -f $icas.Count) 'DADO' }

    if ($tot -gt 0) {
        Add-Achado 'OK' ((T 'diagcitrix.55') -f (Format-Bytes $tot)) '' (T 'diagcitrix.04')
    }
    if ($icas.Count -gt 5) {
        Add-Achado 'ALERTA' ((T 'diagcitrix.56') -f $icas.Count) (T 'diagcitrix.57') (T 'diagcitrix.04') 'Medio'
    }

    Write-Log '' 'DADO'
    Write-Log (T 'diagcitrix.58') 'DADO'
    Write-Log (T 'diagcitrix.59') 'DADO'
}

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
    Write-Titulo (T 'invidentificacao.01')
    try {
        $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $cp = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $bi = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue

        Write-Log ((T 'invidentificacao.02') -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME) 'DADO'
        Write-Log ((T 'invidentificacao.03') -f $so.Caption, $so.BuildNumber) 'DADO'
        Write-Log ((T 'invidentificacao.04') -f $cs.Manufacturer, $cs.Model, $bi.SerialNumber) 'DADO'
        Write-Log ((T 'invidentificacao.05') -f "$($cp.Name)".Trim(), $cp.NumberOfCores) 'DADO'
        Write-Log ((T 'invidentificacao.06') -f (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))) 'DADO'
        Write-Log ((T 'invidentificacao.07') -f $so.InstallDate) 'DADO'

        Add-ItemInv 'Identificacao' 'Maquina' $env:COMPUTERNAME
        Add-ItemInv 'Identificacao' 'Sistema' ("$($so.Caption) build $($so.BuildNumber)")
        Add-ItemInv 'Identificacao' 'Modelo' ("$($cs.Manufacturer) $($cs.Model)")
        Add-ItemInv 'Identificacao' 'Serie' ("$($bi.SerialNumber)")
        Add-ItemInv 'Identificacao' 'Processador' ("$($cp.Name)".Trim())
        Add-ItemInv 'Identificacao' 'Memoria' (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))

        if (($so.TotalVisibleMemorySize * 1KB) -lt 8GB) {
            Add-Achado 'ALERTA' ((T 'invidentificacao.08') -f (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))) (T 'invidentificacao.09') (T 'invidentificacao.10') 'Medio'
        }
    } catch { }

    try {
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)) {
            if (-not $d.Size) { continue }
            Add-ItemInv 'Disco' $d.DeviceID ('{0}% livre' -f [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1)) ('{0} de {1}' -f (Format-Bytes $d.FreeSpace), (Format-Bytes $d.Size))
        }
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=4' -ErrorAction SilentlyContinue)) {
            Write-Log ((T 'invidentificacao.11') -f $d.DeviceID, $d.ProviderName) 'DADO'
            Add-ItemInv 'Unidade de rede' $d.DeviceID "$($d.ProviderName)"
        }
    } catch { }
}

function Invoke-InvSoftware {
    Write-Titulo (T 'invsoftware.01')
    $prog = Get-ProgramasInstalados
    Write-Log ((T 'invsoftware.02') -f $prog.Count) 'DADO'
    foreach ($p in $prog) {
        Write-Log ((T 'invsoftware.03') -f $p.Nome, $p.Versao, $p.Fabricante) 'DADO'
        Add-ItemInv 'Programa' $p.Nome $p.Versao $p.Fabricante
    }

    $procs = @(Get-Process -ErrorAction SilentlyContinue)
    $svcs = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)

    Write-Log '' 'DADO'
    Write-Log (T 'invsoftware.04') 'DADO'
    foreach ($c in (Get-CatalogoClinico)) {
        $achou = @()
        $achou += @($prog  | Where-Object { $_.Nome -match $c.Padrao } | ForEach-Object { 'instalado: ' + $_.Nome })
        $achou += @($procs | Where-Object { $_.ProcessName -match $c.Padrao } | Select-Object -ExpandProperty ProcessName -Unique | ForEach-Object { 'processo: ' + $_ })
        $achou += @($svcs  | Where-Object { $_.Name -match $c.Padrao -or $_.DisplayName -match $c.Padrao } | Select-Object -ExpandProperty Name -Unique | ForEach-Object { 'servico: ' + $_ })
        $achou = @($achou | Select-Object -Unique)
        if ($achou.Count -gt 0) {
            Write-Log ((T 'invsoftware.05') -f $c.Chave, (($achou | Select-Object -First 4) -join ' · ')) 'DADO'
            Add-ItemInv 'Clinico' $c.Chave 'presente' (($achou | Select-Object -First 6) -join ' | ')
        } else {
            Write-Log ((T 'invsoftware.06') -f $c.Chave) 'DADO'
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
            Write-Log ((T 'invsoftware.07') -f $c.Chave, $ps.Count, $sv.Count) 'DADO'
            Add-ItemInv 'Servidor' $c.Chave 'presente' ('processos=' + $ps.Count + ' servicos=' + $sv.Count)
        }
    }
    if ($papel) {
        Add-Achado 'ALERTA' (T 'invsoftware.08') (T 'invsoftware.09') (T 'invsoftware.10') 'Alto'
    } else {
        Add-ItemInv 'Servidor' 'Nenhum' 'estacao comum'
    }

    $agentes = @($svcs | Where-Object { $_.State -eq 'Running' -and ($_.Name -match 'CcmExec|DSASvc|ntrtscan|tmlisten|TMBM|QualysAgent|stAgent|Cortex|cyserver|cyvera|LsAgent|Quest|OnDemand') })
    if ($agentes.Count -gt 0) {
        Write-Log ((T 'invsoftware.11') -f (($agentes | Select-Object -ExpandProperty Name) -join ', ')) 'DADO'
        Add-ItemInv 'Agentes' 'Em execucao' ("$($agentes.Count)") ((($agentes | Select-Object -ExpandProperty Name) -join ', '))
    }
}

function Invoke-InvInicializacaoCompleta {
    Write-Titulo (T 'invinicializacaoco.01')
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
    Write-Log ((T 'invinicializacaoco.02') -f $itens.Count) 'DADO'
    foreach ($i in ($itens | Sort-Object Origem, Nome)) {
        Write-Log ((T 'invinicializacaoco.03') -f $i.Origem, $i.Nome, $i.Comando) 'DADO'
        Add-ItemInv 'Inicializacao' $i.Nome $i.Origem $i.Comando
    }
    $daMaquina = @($itens | Where-Object { $_.Origem -like '*maquina*' }).Count
    if ($daMaquina -ge 8) {
        Add-Achado 'ALERTA' ((T 'invinicializacaoco.04') -f $daMaquina) (T 'invinicializacaoco.05') (T 'invinicializacaoco.06') 'Medio'
    }
}

function Invoke-InvServicos {
    Write-Titulo (T 'invservicos.01')
    try {
        $svcs = @(Get-CimInstance Win32_Service -ErrorAction Stop)
        $rod = @($svcs | Where-Object { $_.State -eq 'Running' } | Sort-Object DisplayName)
        Write-Log ((T 'invservicos.02') -f $rod.Count, $svcs.Count) 'DADO'
        foreach ($sv in $rod) {
            Write-Log ((T 'invservicos.03') -f $sv.Name, $sv.StartMode, $sv.DisplayName) 'DADO'
            Add-ItemInv 'Servico' $sv.Name $sv.StartMode "$($sv.DisplayName)"
        }
    } catch { }
}

function Invoke-InvOfficeExcel {
    Write-Titulo (T 'invofficeexcel.01')
    $ver = Get-ValorReg 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' 'VersionToReport'
    if ($ver) { Write-Log ((T 'invofficeexcel.02') -f $ver) 'DADO'; Add-ItemInv 'Office' 'Versao' "$ver" }

    $locais = @()
    try {
        foreach ($k in (Get-ChildItem -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security\Trusted Locations' -ErrorAction Stop)) {
            $cam = (Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue).Path
            if ($cam) { $locais += $cam; Write-Log ((T 'invofficeexcel.03') -f $cam) 'DADO'; Add-ItemInv 'Excel' 'Local confiavel' $cam }
        }
    } catch { }
    $temLocal = @($locais | Where-Object { $_ -like '*PRESCRI*' }).Count
    if ($temLocal -eq 0) {
        Add-Achado 'ALERTA' (T 'invofficeexcel.04') (T 'invofficeexcel.05') (T 'invofficeexcel.06') 'Alto'
    }

    foreach ($k in @('HKCU:\Software\Microsoft\Office\Excel\Addins', 'HKLM:\Software\Microsoft\Office\Excel\Addins')) {
        try {
            foreach ($a in (Get-ChildItem -Path $k -ErrorAction Stop)) {
                $lb = (Get-ItemProperty -Path $a.PSPath -ErrorAction SilentlyContinue).LoadBehavior
                Write-Log ((T 'invofficeexcel.07') -f $a.PSChildName, $lb) 'DADO'
                Add-ItemInv 'Excel' 'Suplemento' $a.PSChildName "LoadBehavior=$lb"
                if ($lb -eq 3 -and $a.PSChildName -match 'PDFMaker|Acrobat') {
                    Add-Achado 'ALERTA' ((T 'invofficeexcel.08') -f $a.PSChildName) (T 'invofficeexcel.09') (T 'invofficeexcel.06') 'Medio'
                }
            }
        } catch { }
    }
}

function Invoke-InvPastaClinica {
    Write-Titulo (T 'invpastaclinica.01')
    if (-not (Test-Path -LiteralPath $script:PastaClinica)) {
        Add-Achado 'ALERTA' ((T 'invpastaclinica.02') -f $script:PastaClinica) (T 'invpastaclinica.03') (T 'invpastaclinica.04') 'Alto'
        return
    }
    $tot = Get-TamanhoPasta -Caminho $script:PastaClinica -TimeoutSeg 120
    Write-Log ((T 'invpastaclinica.05') -f (Format-Bytes $tot)) 'DADO'
    Add-ItemInv 'Pasta clinica' 'Total' (Format-Bytes $tot)

    try {
        foreach ($d in (Get-ChildItem -LiteralPath $script:PastaClinica -Directory -ErrorAction Stop)) {
            if ($script:Cancelar) { break }
            $b = Get-TamanhoPasta -Caminho $d.FullName -TimeoutSeg 45
            Write-Log ((T 'invpastaclinica.06') -f $d.Name, (Format-Bytes $b)) 'DADO'
            Add-ItemInv 'Pasta clinica' $d.Name (Format-Bytes $b)
        }
        foreach ($f in (Get-ChildItem -LiteralPath $script:PastaClinica -File -Filter '*.xls*' -ErrorAction Stop)) {
            Write-Log ((T 'invpastaclinica.07') -f $f.Name, (Format-Bytes $f.Length), $f.LastWriteTime) 'DADO'
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
        Add-Achado 'ALERTA' ((T 'invpastaclinica.08') -f ($falta -join ' e ')) (T 'invofficeexcel.05') (T 'invpastaclinica.04') 'Alto'
    }
}

function Invoke-InvRedeImpressoras {
    Write-Titulo (T 'invredeimpressoras.01')
    try {
        foreach ($n in (Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' })) {
            $ip = (Get-NetIPAddress -InterfaceIndex $n.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress -join ', '
            Write-Log ((T 'invredeimpressoras.02') -f $n.InterfaceDescription, $n.LinkSpeed, $ip) 'DADO'
            Add-ItemInv 'Rede' $n.Name $n.LinkSpeed $ip
            if ($n.InterfaceDescription -match 'Wi-?Fi|Wireless|802\.11' -or $n.Name -match 'Wi-?Fi') {
                Add-Achado 'ALERTA' ((T 'invredeimpressoras.03') -f $n.LinkSpeed) (T 'invredeimpressoras.04') (T 'invredeimpressoras.05') 'Alto'
            } elseif ($n.LinkSpeed -match '^(10|100) Mbps') {
                Add-Achado 'ALERTA' ((T 'invredeimpressoras.06') -f $n.LinkSpeed) (T 'invredeimpressoras.07') (T 'invredeimpressoras.05') 'Alto'
            }
        }
    } catch { }
    try {
        $sufixo = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
        Write-Log ((T 'invredeimpressoras.08') -f $sufixo) 'DADO'
        Add-ItemInv 'Rede' 'Sufixo DNS' "$sufixo"
    } catch { }
    try {
        foreach ($imp in (Get-CimInstance Win32_Printer -ErrorAction Stop)) {
            $marca = if ($imp.Default) { '[padrao] ' } else { '' }
            Write-Log ((T 'invredeimpressoras.09') -f $marca, $imp.Name, $imp.PortName) 'DADO'
            Add-ItemInv 'Impressora' $imp.Name $(if ($imp.Default) { 'padrao' } else { '' }) "$($imp.PortName)"
            if ($imp.Default -and ($imp.WorkOffline -or $imp.PrinterStatus -eq 7)) {
                Add-Achado 'ALERTA' ((T 'invredeimpressoras.10') -f $imp.Name) (T 'invredeimpressoras.11') (T 'invredeimpressoras.12') 'Alto'
            }
        }
    } catch { }
}

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
    Write-Titulo (T 'invmigracao.01')

    $prog = Get-ProgramasInstalados
    $ag = @($prog | Where-Object { $_.Nome -match 'On Demand Migration|Quest.*Migration|Binary Tree|ADMT' })
    if ($ag.Count -gt 0) {
        foreach ($a in $ag) { Write-Log ((T 'invmigracao.02') -f $a.Nome, $a.Versao) 'DADO'; Add-ItemInv 'Migracao' 'Agente' $a.Nome $a.Versao }
        Add-Achado 'ALERTA' ((T 'invmigracao.03') -f $ag[0].Nome) (T 'invmigracao.04') (T 'invmigracao.05') 'Alto'
    } else {
        Write-Log (T 'invmigracao.06') 'DADO'
    }

    $pastaPerfil = Split-Path $env:USERPROFILE -Leaf
    if ($pastaPerfil -and ($pastaPerfil -ne $env:USERNAME)) {
        Write-Log ((T 'invmigracao.07') -f $env:USERNAME) 'DADO'
        Write-Log ((T 'invmigracao.08') -f $env:USERPROFILE) 'DADO'
        Add-ItemInv 'Migracao' 'Perfil com nome de outra conta' $pastaPerfil $env:USERNAME
        Add-Achado 'ALERTA' ((T 'invmigracao.09') -f $env:USERNAME, $pastaPerfil) (T 'invmigracao.10') (T 'invmigracao.05') 'Medio'
    }

    if ($env:USERPROFILE -match '\.[A-Za-z0-9]+$') {
        Write-Log ((T 'invmigracao.11') -f $env:USERPROFILE) 'DADO'
        Add-ItemInv 'Migracao' 'Perfil com sufixo' $env:USERPROFILE
        Write-Log (T 'invmigracao.12') 'DADO'
    }

    $orfas = Get-CredenciaisOrfas
    if ($orfas.Count -gt 0) {
        Write-Log ((T 'invmigracao.13') -f $orfas.Count) 'DADO'
        foreach ($o in $orfas) { Write-Log ('   ' + $o.Alvo) 'DADO'; Add-ItemInv 'Migracao' 'Credencial orfa' $o.Servidor $o.Alvo }
        Add-Achado 'ALERTA' ((T 'invmigracao.14') -f $orfas.Count) (T 'invmigracao.15') (T 'invmigracao.05') 'Alto'
    } else {
        Write-Log (T 'invmigracao.16') 'DADO'
    }
}

function Save-Inventario {
    Write-Titulo (T 'inventario.01')
    $carimbo = Get-Date -Format 'yyyyMMdd_HHmm'
    $base = 'Inventario_{0}_{1}' -f $env:COMPUTERNAME, $carimbo
    $pasta = $script:PastaRelatorios

    $acessivel = $false
    try {
        if ($pasta -like '\\*') {
            $srv = ($pasta.TrimStart('\') -split '\\')[0]
            $t = Test-PortaTcp -Alvo $srv -Porta 445 -TimeoutMs 1500
            if (-not $t.Ok) {
                Write-Log ((T 'inventario.02') -f $srv) 'DADO'
                Write-Log (T 'inventario.03') 'DADO'
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
        Write-Log ((T 'inventario.04') -f $pasta) 'DADO'
        Write-Log (T 'inventario.03') 'DADO'
        return
    }
    $csv = Join-Path $pasta ($base + '.csv')
    try {
        $script:ItensInv | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Force
        Write-Log ((T 'inventario.05') -f $csv) 'OK'
        Write-Log (T 'inventario.06') 'DADO'
        Write-Log (T 'inventario.07') 'ACAO'
    } catch {
        Write-Log ((T 'inventario.08') -f $_.Exception.Message) 'ALERTA'
    }
}

function Get-VeredictoArquivo {
    param($Arquivo)
    $p = "$($Arquivo.FullName)"
    $e = "$($Arquivo.Extension)".ToLower()

    if ($Arquivo.Name -match '^(hiberfil|pagefile|swapfile)\.sys$')      { return @{ V = 'MANTER';  M = 'Arquivo de sistema do Windows. Apagar quebra a maquina.' } }
    if ($e -match '^\.(mdf|ldf|ndf|bak|trn|dbf)$')                       { return @{ V = 'MANTER';  M = 'Arquivo de banco de dados. Apagar derruba o sistema.' } }
    if (Test-DentroDaPasta -Caminho $p -Pasta $script:PastaClinica)      { return @{ V = 'MANTER';  M = 'Esta na pasta clinica.' } }
    if (Test-Padrao $p $script:RaizesClinicas) { return @{ V = 'MANTER'; M = 'Dado de sistema clinico.' } }
    if ($p -match '\\CentBrowser')      { return @{ V = 'MANTER'; M = 'CentBrowser e o navegador do prontuario da clinica.' } }
    if ($p -match 'Digitalcore|\\Onis') { return @{ V = 'MANTER'; M = 'Dados do visualizador DICOM Onis.' } }
    if ($p -match '\\Citrix\\')         { return @{ V = 'MANTER'; M = 'Cache do Citrix. Este kit nunca o toca.' } }
    if ($e -match '^\.(dcm|dicom|nii|nrrd|mha|raw)$')                    { return @{ V = 'MANTER';  M = 'Imagem medica. Mova para a rede ou PACS, nao apague.' } }
    if ($e -match '^\.(ost|pst)$')                                       { return @{ V = 'MANTER';  M = 'Caixa de e-mail. Reduza pelo Outlook, nunca apague.' } }
    if ($p -match '\\Program Files|\\Windows\\|\\ProgramData\\|\\Microsoft SQL Server\\|\\Tomcat') { return @{ V = 'MANTER'; M = 'Parte de um programa instalado.' } }
    if ($e -match '^\.(exe|dll|msi|sys)$' -and $p -match '\\AppData\\Local\\' -and $p -notmatch '\\Temp\\|\\Downloads\\|[Cc]ache') { return @{ V = 'MANTER'; M = 'Programa instalado dentro do seu perfil.' } }
    if ($e -match '^\.(xlsb|xlsm)$')                                     { return @{ V = 'MANTER';  M = 'Planilha de trabalho.' } }

    if ($e -match '^\.(dmp|mdmp|hdmp)$')                                 { return @{ V = 'SEGURO';  M = 'Despejo de travamento. Nao tem uso.' } }
    if ($e -match '^\.(iso|img|vhd|vhdx|wim|esd)$')                      { return @{ V = 'SEGURO';  M = 'Imagem de instalacao, serve so uma vez.' } }
    if ($e -match '^\.(msi|msp|exe)$' -and $p -match '\\Downloads\\|\\Temp\\|\\Temporar|C:\\Temp') { return @{ V = 'SEGURO'; M = 'Instalador ja usado.' } }
    if ($p -match 'component_crx_cache|GrShaderCache|ShaderCache|ProvenanceData|EBWebView') { return @{ V = 'SEGURO'; M = 'Cache interno do navegador. E recriado sozinho.' } }
    if ($p -match '\\Temp\\|\\Temporar|[Cc]ache|\\CrashDumps\\|\\WER\\|\\INetCache\\|\\Downloaded Installations\\') { return @{ V = 'SEGURO'; M = 'Esta em pasta de cache ou temporaria.' } }
    if ($e -match '^\.(log|etl|old|tmp|chk|gid)$')                       { return @{ V = 'SEGURO';  M = 'Log ou sobra de instalacao.' } }

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

function Get-TasyUrlsDescobertas {
    $achadas = @()

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
    Write-Titulo (T 'diagtasy.01')

    $prog = Get-ProgramasInstalados
    $tasyProg = @($prog | Where-Object { $_.Nome -match 'Tasy|Philips.*Tasy|Wheb' })
    foreach ($t in $tasyProg) { Write-Log ((T 'diagtasy.02') -f $t.Nome, $t.Versao) 'DADO'; Add-ItemInv 'Tasy' 'Cliente' $t.Nome $t.Versao }
    if ($tasyProg.Count -eq 0) { Write-Log (T 'diagtasy.03') 'DADO' }

    $navProcs = @(Get-Process -Name 'CentBrowser', 'msedge', 'chrome' -ErrorAction SilentlyContinue)
    if ($navProcs.Count -gt 0) {
        $porNav = $navProcs | Group-Object ProcessName | ForEach-Object {
            [pscustomobject]@{ Nome = $_.Name; Qtd = $_.Count; Mem = ($_.Group | Measure-Object WorkingSet64 -Sum).Sum }
        } | Sort-Object Mem -Descending
        foreach ($n in $porNav) {
            Write-Log ((T 'diagtasy.04') -f $n.Nome, (Format-Bytes $n.Mem), $n.Qtd) 'DADO'
            Add-ItemInv 'Tasy' ('Navegador ' + $n.Nome) (Format-Bytes $n.Mem) ("$($n.Qtd) processos")
        }
        $pesado = $porNav | Select-Object -First 1
        if ($pesado -and $pesado.Mem -gt 2.5GB) {
            Add-Achado 'ALERTA' ((T 'diagtasy.05') -f $pesado.Nome, (Format-Bytes $pesado.Mem), $pesado.Qtd) (T 'diagtasy.06') (T 'diagtasy.07') 'Alto'
        }
    }

    try {
        $javas = @(Get-CimInstance Win32_Process -Filter "Name='javaw.exe' OR Name='java.exe'" -ErrorAction SilentlyContinue)
        foreach ($j in $javas) {
            $xmx = ''
            if ("$($j.CommandLine)" -match '-Xmx(\d+)([mMgG])') { $xmx = ('limite de memoria ' + $Matches[1] + $Matches[2].ToUpper()) }
            $bits = $(if ("$($j.ExecutablePath)" -match 'Program Files \(x86\)') { '32 bits' } else { '64 bits' })
            Write-Log ((T 'diagtasy.08') -f (Format-Bytes $j.WorkingSetSize), $bits, $xmx) 'DADO'
            if ($bits -eq '32 bits') {
                Add-Achado 'ALERTA' (T 'diagtasy.09') (T 'diagtasy.10') (T 'diagtasy.07') 'Medio'
            }
        }
    } catch { }

    $urls = Get-TasyUrlsEfetivas
    if ($urls.Count -eq 0) {
        Add-Achado 'ALERTA' (T 'diagtasy.11') 'Preencha $script:TasyUrls na secao 1 do aplicativo.' (T 'diagtasy.07') 'Medio'
        return
    }

    $resumo = @()
    foreach ($d in $urls) {
        if ($script:Cancelar) { return }
        $i = Get-InfoUrl -Url $d.Url
        if (-not $i) { continue }

        Write-Log '' 'DADO'
        Write-Log ((T 'diagcitrix.15') -f $d.Nome) 'DADO'
        Write-Log ((T 'diagcitrix.17') -f $d.Url) 'DADO'
        Set-Status ((T 'diagtasy.12') + $d.Nome + (T 'diagcitrix.19'))
        Add-ItemInv 'Tasy' $d.Nome $d.Url

        $interno = Test-EnderecoInterno -Servidor $i.Servidor
        Write-Log ((T 'diagtasy.13') -f $(if ($interno) { (T 'diagtasy.14') } else { (T 'diagtasy.15') })) 'DADO'

        $alvo = $i.Servidor
        if (-not $i.EhIp) {
            try {
                $rel = [System.Diagnostics.Stopwatch]::StartNew()
                $ips = [System.Net.Dns]::GetHostAddresses($i.Servidor) | Select-Object -ExpandProperty IPAddressToString
                $rel.Stop()
                Write-Log ((T 'diagtasy.16') -f ($ips -join ', '), $rel.ElapsedMilliseconds) 'DADO'
            } catch {
                Add-Achado 'CRITICO' ((T 'diagtasy.17') -f $d.Nome, $i.Servidor) (T 'diagtasy.18') (T 'diagtasy.07') 'Alto'
                $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = -1; Tcp = -1; Situacao = 'DNS nao resolve' }
                continue
            }
        }

        $porta = $i.Porta
        $abertas = @()
        foreach ($p in @($porta, 80, 443, 8080, 28080 | Select-Object -Unique)) {
            $t = Test-PortaTcp -Alvo $alvo -Porta $p -TimeoutMs 1200
            if ($t.Ok) { $abertas += ('{0} ({1} ms)' -f $p, $t.Ms) }
        }
        Write-Log ((T 'diagtasy.19') -f $(if ($abertas.Count -gt 0) { $abertas -join ' · ' } else { (T 'diagcitrix.26') })) 'DADO'
        if ($abertas.Count -eq 0) {
            Add-Achado 'CRITICO' ((T 'diagtasy.20') -f $d.Nome, $alvo) (T 'diagtasy.21') (T 'diagtasy.07') 'Alto'
            $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = -1; Tcp = -1; Situacao = 'servidor sem resposta' }
            continue
        }

        Write-Log (T 'diagtasy.22') 'DADO'
        $m = Measure-RespostaHttp -Url $d.Url -Alvo $alvo -Porta $porta -Amostras 4

        if ($m.Amostras -eq 0) {
            Add-Achado 'CRITICO' ((T 'diagtasy.23') -f $d.Nome) (T 'diagtasy.24') (T 'diagtasy.07') 'Alto'
            $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = -1; Tcp = $m.TcpMed; Situacao = 'pagina nao responde' }
            continue
        }

        Write-Log ((T 'diagtasy.25') -f $m.TcpMin, $m.TcpMed, $m.TcpMax) 'DADO'
        Write-Log ((T 'diagtasy.26') -f $m.HttpMin, $m.HttpMed, $m.HttpMax, $m.Codigo) 'DADO'
        Write-Log ((T 'diagtasy.27') -f $m.Processamento) 'DADO'
        Add-ItemInv 'Tasy' ($d.Nome + ' - resposta') ("$($m.HttpMed) ms") ("tcp=$($m.TcpMed)ms http=$($m.HttpMed)ms")
        if ($m.TcpMax -gt 0 -and $m.TcpMed -gt 0 -and $m.TcpMax -gt ($m.TcpMed * 4) -and $m.TcpMax -gt 300) {
            Write-Log ((T 'diagtasy.28') -f $m.TcpMax, $m.TcpMed) 'DADO'
        }

        $situacao = 'ok'
        if ($m.HttpMed -gt 3000) {
            $situacao = 'servidor lento'
            Add-Achado 'CRITICO' ((T 'diagtasy.29') -f $d.Nome, $m.HttpMed) ((T 'diagtasy.30') -f $m.TcpMed) (T 'diagtasy.07') 'Alto'
        } elseif ($m.HttpMed -gt 1500) {
            $situacao = 'servidor pesado'
            Add-Achado 'ALERTA' ((T 'diagtasy.29') -f $d.Nome, $m.HttpMed) ((T 'diagtasy.31') -f $m.TcpMed) (T 'diagtasy.07') 'Alto'
        } else {
            Add-Achado 'OK' ((T 'diagtasy.32') -f $d.Nome, $m.HttpMed) '' (T 'diagtasy.07')
        }

        if ($m.TcpMed -gt 60 -and $interno) {
            Add-Achado 'ALERTA' ((T 'diagtasy.33') -f $d.Nome, $m.TcpMed) (T 'diagtasy.34') (T 'diagtasy.07') 'Medio'
        }
        if ($m.Instavel) {
            $situacao = 'instavel'
            Add-Achado 'ALERTA' ((T 'diagtasy.35') -f $d.Nome, $m.HttpMin, $m.HttpMax) (T 'diagtasy.36') (T 'diagtasy.07') 'Alto'
        }

        if (-not $i.Https) {
            if ("$($m.UrlFinal)" -match '^https://') {
                Write-Log (T 'diagcitrix.43') 'DADO'
            } else {
                $alt = Test-HttpsAlternativo -Url $d.Url
                if ($alt -and $alt.Disponivel) {
                    Add-Achado 'ALERTA' ((T 'diagcitrix.44') -f $d.Nome) ((T 'diagcitrix.45') -f $alt.Url) (T 'diagtasy.07') 'Medio'
                } else {
                    Write-Log (T 'diagcitrix.46') 'DADO'
                }
            }
        }
        if ($m.CertificadoInvalido) {
            Add-Achado 'ALERTA' ((T 'diagcitrix.47') -f $d.Nome) (T 'diagcitrix.48') (T 'diagtasy.07') 'Alto'
        }

        if ($interno -and -not (Test-ZonaIntranet -Servidor $i.Servidor)) {
            Write-Log (T 'diagtasy.37') 'DADO'
        }

        $resumo += [pscustomobject]@{ Nome = $d.Nome; Http = $m.HttpMed; Tcp = $m.TcpMed; Situacao = $situacao }
    }

    if ($resumo.Count -gt 1) {
        Write-Log '' 'DADO'
        Write-Log (T 'diagtasy.38') 'TITULO'
        Write-Log ((T 'diagtasy.39') -f (T 'invpastaclinica.04'), (T 'invredeimpressoras.05'), (T 'diagtasy.40'), (T 'diagtasy.41')) 'DADO'
        foreach ($r in ($resumo | Sort-Object Http -Descending)) {
            $t = $(if ($r.Tcp -lt 0) { '-' } else { "$($r.Tcp) ms" })
            $h = $(if ($r.Http -lt 0) { '-' } else { "$($r.Http) ms" })
            Write-Log ((T 'diagtasy.39') -f $r.Nome, $t, $h, $r.Situacao) 'DADO'
        }
        $bons = @($resumo | Where-Object { $_.Http -ge 0 -and $_.Http -le 1500 })
        $ruins = @($resumo | Where-Object { $_.Http -gt 1500 })
        if ($bons.Count -gt 0 -and $ruins.Count -gt 0) {
            Write-Log '' 'DADO'
            Write-Log ((T 'diagtasy.42') -f $bons[0].Nome, $ruins[0].Nome) 'ACAO'
            Write-Log (T 'diagtasy.43') 'ACAO'
        }
    }

    Write-Log '' 'DADO'
    $cj = Get-CacheJava
    $totJava = 0.0
    foreach ($c in $cj) { $b = Get-TamanhoPasta -Caminho $c -TimeoutSeg 45; if ($b -gt 0) { $totJava += $b } }
    if ($totJava -gt 0) {
        Write-Log ((T 'diagtasy.44') -f (Format-Bytes $totJava)) 'DADO'
        Add-ItemInv 'Tasy' 'Cache Java Web Start' (Format-Bytes $totJava)
        if ($totJava -gt 200MB) {
            Add-Achado 'ALERTA' ((T 'diagtasy.45') -f (Format-Bytes $totJava)) (T 'diagtasy.46') (T 'diagtasy.07') 'Medio'
        }
    }

    $centCache = Join-Path $env:LOCALAPPDATA 'CentBrowser\User Data\Default\Cache'
    if (Test-Path -LiteralPath $centCache) {
        $b = Get-TamanhoPasta -Caminho $centCache -TimeoutSeg 45
        Write-Log ((T 'diagtasy.47') -f (Format-Bytes $b)) 'DADO'
        Add-ItemInv 'Tasy' 'Cache do CentBrowser' (Format-Bytes $b)
        if ($b -gt 300MB) {
            Add-Achado 'ALERTA' ((T 'diagtasy.48') -f (Format-Bytes $b)) (T 'diagtasy.49') (T 'diagtasy.07') 'Medio'
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
            Add-Achado 'ALERTA' ((T 'diagtasy.50') -f ($faltando -join ', ')) (T 'diagtasy.51') (T 'diagtasy.07') 'Alto'
        }
    }

    Write-Log '' 'DADO'
    Write-Log (T 'diagtasy.52') 'TITULO'
    Write-Log (T 'diagtasy.53') 'ACAO'
    Write-Log (T 'diagtasy.54') 'ACAO'
    Write-Log (T 'diagtasy.55') 'ACAO'
    Write-Log (T 'diagtasy.56') 'ACAO'
    Write-Log '' 'DADO'
    Write-Log (T 'diagtasy.57') 'ACAO'
}

function Invoke-DiagMemoria {
    Write-Titulo (T 'diagmemoria.01')
    try {
        $so     = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $totalB = $so.TotalVisibleMemorySize * 1KB
        $livreB = $so.FreePhysicalMemory * 1KB
        $usoPct = [Math]::Round((($totalB - $livreB) / $totalB) * 100, 1)

        $txt = ('Memoria em uso: {0}% ({1} de {2}, livres {3})' -f $usoPct, (Format-Bytes ($totalB - $livreB)), (Format-Bytes $totalB), (Format-Bytes $livreB))
        if ($usoPct -ge $script:Lim.RamCritPct)       { Add-Achado 'CRITICO' $txt (T 'diagmemoria.02') }
        elseif ($usoPct -ge $script:Lim.RamAlertaPct) { Add-Achado 'ALERTA'  $txt (T 'diagmemoria.03') }
        else                                          { Add-Achado 'OK' $txt }

        if (($totalB / 1GB) -lt 8) {
            Add-Achado 'ALERTA' ((T 'diagmemoria.04') -f (Format-Bytes $totalB)) (T 'diagmemoria.05')
        }
    } catch { Write-Log ((T 'diagmemoria.06') -f $_.Exception.Message) 'ALERTA' }

    Write-Log (T 'diagmemoria.07') 'DADO'
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
            Write-Log ((T 'diagmemoria.08') -f $g.Nome, (Format-Bytes $g.Memoria), $g.Processos) 'DADO'
        }

        $navegadores = $grupos | Where-Object { $_.Nome -match 'msedge|chrome|firefox' }
        foreach ($n in $navegadores) {
            if ($n.Memoria -gt 2GB) {
                Add-Achado 'ALERTA' ((T 'diagtasy.05') -f $n.Nome, (Format-Bytes $n.Memoria), $n.Processos) (T 'diagmemoria.09')
            }
        }
        $teams = $grupos | Where-Object { $_.Nome -match 'Teams|ms-teams' }
        if ($teams -and (($teams | Measure-Object Memoria -Sum).Sum -gt 1.5GB)) {
            Add-Achado 'ALERTA' (T 'diagmemoria.10') (T 'diagmemoria.11')
        }
    } catch { }
}

function Invoke-DiagCaches {
    Write-Titulo (T 'diagcaches.01')
    $alvos = Get-AlvosLimpeza
    $total = 0.0
    foreach ($a in $alvos) {
        if ($script:Cancelar) { return }
        Set-Status ((T 'diagtasy.12') + $a.Nome + (T 'diagcitrix.19'))
        $b = Measure-Alvo -Alvo $a
        $a.Bytes = $b
        if ($b -gt 0) {
            Write-Log ((T 'invpastaclinica.06') -f $a.Nome, (Format-Bytes $b)) 'DADO'
            if ($a.Padrao) { $total += $b }
        }
    }
    $script:AlvosAtuais = $alvos
    if ($total -gt 1GB) {
        Add-Achado 'ALERTA' ((T 'diagcaches.02') -f (Format-Bytes $total)) (T 'diagcaches.03')
    } else {
        Add-Achado 'OK' ((T 'diagcaches.04') -f (Format-Bytes $total))
    }
}

function Invoke-DiagSistemaAcionavel {
    Write-Titulo (T 'diagsistemaacionav.01')
    try {
        $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        Write-Log ((T 'diagsistemaacionav.02') -f $so.Caption, $so.BuildNumber, (Format-Bytes ($so.TotalVisibleMemorySize * 1KB))) 'DADO'

        $horas = ((Get-Date) - $so.LastBootUpTime).TotalHours
        $txt = ('{0:N0} dias e {1:N0} horas ligado sem reiniciar' -f [Math]::Floor($horas / 24), ($horas % 24))
        if ($horas -ge $script:Lim.UptimeCritH) {
            Add-Achado 'CRITICO' $txt (T 'diagsistemaacionav.03') (T 'diagsistemaacionav.04') 'Alto'
        } elseif ($horas -ge $script:Lim.UptimeAlertaH) {
            Add-Achado 'ALERTA' $txt (T 'diagsistemaacionav.05') (T 'diagsistemaacionav.04') 'Medio'
        } else {
            Add-Achado 'OK' $txt '' (T 'diagsistemaacionav.04')
        }

        if ((Get-ValorReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled') -eq 1) {
            Write-Log (T 'diagsistemaacionav.06') 'DADO'
        }
    } catch { }
}

function Invoke-DiagEspaco {
    Write-Titulo (T 'diagespaco.01')
    try {
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)) {
            if (-not $d.Size) { continue }
            $pct = [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
            $txt = ('Unidade {0}: {1} livres de {2} ({3}%)' -f $d.DeviceID, (Format-Bytes $d.FreeSpace), (Format-Bytes $d.Size), $pct)
            if ($pct -lt $script:Lim.DiscoCritPct) {
                Add-Achado 'CRITICO' $txt (T 'diagespaco.02') (T 'diagespaco.03') 'Alto'
            } elseif ($pct -lt $script:Lim.DiscoAlertaPct) {
                Add-Achado 'ALERTA' $txt (T 'diagespaco.04') (T 'diagespaco.03') 'Alto'
            } else {
                Add-Achado 'OK' $txt '' (T 'diagespaco.03')
            }
        }
    } catch { Write-Log (T 'diagespaco.05') 'ALERTA' }

    try {
        $wt = Get-TamanhoPasta -Caminho (Join-Path $env:SystemRoot 'Temp') -TimeoutSeg 45
        if ($wt -gt 1GB) {
            Add-Achado 'ALERTA' ((T 'diagespaco.06') -f (Format-Bytes $wt)) (T 'diagespaco.07') (T 'diagespaco.03') 'Medio'
        } elseif ($wt -gt 0) {
            Write-Log ((T 'diagespaco.08') -f (Format-Bytes $wt)) 'DADO'
        }
    } catch { }
}

function Invoke-DiagLixeira {
    Write-Titulo (T 'diaglixeira.01')
    $alvo = New-Alvo -Id 'LIXEIRA' -Nome (T 'alvo.01') -Modo 'Lixeira' -Caminhos @()
    $b = Measure-Alvo -Alvo $alvo
    if ($b -gt 500MB) {
        Add-Achado 'ALERTA' ((T 'diaglixeira.02') -f (Format-Bytes $b)) (T 'diaglixeira.03') (T 'diagespaco.03') 'Alto'
    } else {
        Add-Achado 'OK' ((T 'diaglixeira.02') -f (Format-Bytes $b)) '' (T 'diagespaco.03')
    }
}

function Invoke-DiagEncerraveis {
    Write-Titulo (T 'diagencerraveis.01')
    Write-Log (T 'diagencerraveis.02') 'DADO'

    $enc = Get-ProcessosEncerraveis
    if ($enc.Count -eq 0) {
        Add-Achado 'OK' (T 'diagencerraveis.03') '' (T 'diagencerraveis.04')
        Write-Log (T 'diagencerraveis.05') 'DADO'
        return
    }

    $total = ($enc | Measure-Object Memoria -Sum).Sum
    foreach ($e in ($enc | Sort-Object Memoria -Descending)) {
        Write-Log ((T 'diagencerraveis.06') -f $e.Rotulo, (Format-Bytes $e.Memoria), $e.Qtd) 'DADO'
        if ($e.Nota) { Write-Log ('   ' + $e.Nota) 'DADO' }
    }

    if ($total -gt 500MB) {
        Add-Achado 'ALERTA' ((T 'diagencerraveis.07') -f (Format-Bytes $total), $enc.Count) (T 'diagencerraveis.08') (T 'diagencerraveis.04') 'Alto'
    } else {
        Add-Achado 'OK' ((T 'diagencerraveis.09') -f (Format-Bytes $total)) '' (T 'diagencerraveis.04')
    }

    Write-Log '' 'DADO'
    Write-Log (T 'diagencerraveis.10') 'DADO'
}

function Invoke-DiagInicioAcionavel {
    Write-Titulo (T 'diaginicioacionave.01')

    try {
        $so  = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $exp = $null
        foreach ($e in @(Get-Process explorer -ErrorAction SilentlyContinue)) {
            try { $st = $e.StartTime } catch { continue }
            if (-not $exp -or $st -lt $exp.StartTime) { $exp = $e }
        }
        if ($exp -and $exp.StartTime -gt $so.LastBootUpTime) {
            $seg = [Math]::Round(($exp.StartTime - $so.LastBootUpTime).TotalSeconds)
            $script:LogonSegundos = $(if ($seg -le 1200) { $seg } else { 0 })
            if ($seg -gt 1200) {

                Write-Log ((T 'diaginicioacionave.02') -f ($seg / 3600)) 'DADO'
                Write-Log (T 'diaginicioacionave.03') 'DADO'
                Write-Log (T 'diaginicioacionave.04') 'ACAO'
            }
            elseif ($seg -gt 180) { Add-Achado 'CRITICO' ((T 'diaginicioacionave.05') -f $seg) (T 'diaginicioacionave.06') (T 'invinicializacaoco.06') 'Alto' }
            elseif ($seg -gt 90)  { Add-Achado 'ALERTA'  ((T 'diaginicioacionave.05') -f $seg) (T 'diaginicioacionave.07') (T 'invinicializacaoco.06') 'Alto' }
            else                  { Add-Achado 'OK' ((T 'diaginicioacionave.08') -f $seg) '' (T 'invinicializacaoco.06') }
        }
    } catch { }

    $itens = Get-ItensInicializacao
    $seguros = @($itens | Where-Object { $_.Seguro })
    $outros  = @($itens | Where-Object { -not $_.Seguro })

    Write-Log ((T 'diaginicioacionave.09') -f $itens.Count) 'DADO'
    Write-Log '' 'DADO'
    if ($seguros.Count -gt 0) {
        Write-Log (T 'diaginicioacionave.10') 'DADO'
        foreach ($i in $seguros) { Write-Log ((T 'diaginicioacionave.11') -f $i.Nome, $i.Rotulo) 'DADO' }
    }
    if ($outros.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log (T 'diaginicioacionave.12') 'DADO'
        foreach ($i in $outros) { Write-Log ((T 'diaginicioacionave.13') -f $i.Nome, $i.Rotulo) 'DADO' }
    }

    if ($script:LogonSegundos -gt 90) {
        $unidades = @()
        try { $unidades = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=4' -ErrorAction SilentlyContinue) } catch { }
        $impRede = @()
        try { $impRede = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | Where-Object { $_.Network -or "$($_.PortName)" -like '\\*' -or "$($_.Name)" -like '\\*' }) } catch { }

        if ($unidades.Count -gt 0 -or $impRede.Count -gt 0) {
            Write-Log '' 'DADO'
            Write-Log (T 'diaginicioacionave.14') 'DADO'
            Write-Log ((T 'diaginicioacionave.15') -f $unidades.Count, $impRede.Count) 'DADO'

            $lentos = @()
            foreach ($u in $unidades) {
                $srv = ("$($u.ProviderName)".Trim() -replace '^\\+', '' -split '\\')[0]
                if (-not $srv) { continue }
                $t = Test-PortaTcp -Alvo $srv -Porta 445 -TimeoutMs 700
                Write-Log ((T 'diaginicioacionave.16') -f $u.DeviceID, $u.ProviderName, $(if ($t.Ok) { "$($t.Ms) ms" } else { (T 'diagcitrix.29') })) 'DADO'
                if (-not $t.Ok) { $lentos += ('{0} ({1})' -f $u.DeviceID, $srv) }
            }
            foreach ($i in ($impRede | Select-Object -First 6)) { Write-Log ((T 'diaginicioacionave.17') -f $i.Name) 'DADO' }

            if ($lentos.Count -gt 0) {
                Add-Achado 'CRITICO' ((T 'diaginicioacionave.18') -f $lentos.Count, ($lentos -join ', ')) (T 'diaginicioacionave.19') (T 'invinicializacaoco.06') 'Alto'
            } elseif ($seguros.Count -lt 3) {
                Add-Achado 'ALERTA' ((T 'diaginicioacionave.20') -f $seguros.Count) ((T 'diaginicioacionave.21') -f $unidades.Count, $impRede.Count) (T 'invinicializacaoco.06') 'Alto'
            }
        }
    }

    if ($seguros.Count -ge 3) {
        Add-Achado 'ALERTA' ((T 'diaginicioacionave.22') -f $seguros.Count) (T 'diaginicioacionave.23') (T 'invinicializacaoco.06') 'Alto'
    } elseif ($seguros.Count -gt 0) {
        Add-Achado 'ALERTA' ((T 'diaginicioacionave.24') -f $seguros.Count) (T 'diaginicioacionave.25') (T 'invinicializacaoco.06') 'Medio'
    } else {
        Add-Achado 'OK' (T 'diaginicioacionave.26') '' (T 'invinicializacaoco.06')
    }
    if ($outros.Count -gt 0) {
        Write-Log (T 'diaginicioacionave.27') 'DADO'
    }
}

function Invoke-DiagAjustesPendentes {
    Write-Titulo (T 'diagajustespendent.01')
    $aj = Get-AjustesPendentes
    if ($aj.Count -eq 0) {
        Add-Achado 'OK' (T 'diagajustespendent.02') '' (T 'diagajustespendent.03')
        return
    }
    foreach ($a in $aj) {
        Write-Log ((T 'diagajustespendent.04') -f $a.Rotulo) 'DADO'
        Write-Log ('   ' + $a.Nota) 'DADO'
    }
    Add-Achado 'ALERTA' ((T 'diagajustespendent.05') -f $aj.Count) (T 'diagajustespendent.06') (T 'diagajustespendent.03') 'Alto'
}

function Invoke-DiagMapeamentos {
    Write-Titulo (T 'diagmapeamentos.01')
    $mortos = Get-MapeamentosMortos
    if ($mortos.Count -eq 0) {
        Add-Achado 'OK' (T 'diagmapeamentos.02') '' (T 'invredeimpressoras.05')
        return
    }
    foreach ($m in $mortos) { Write-Log ((T 'diagmapeamentos.03') -f $m.Letra, $m.Destino) 'DADO' }
    Add-Achado 'CRITICO' ((T 'diagmapeamentos.04') -f $mortos.Count) (T 'diagmapeamentos.05') (T 'invredeimpressoras.05') 'Alto'
}

$script:MarcadorChaveReg = 'HKCU:\Software\KitSuporteRT'

function Get-SessaoAtual {

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

    $lista = @(
        [pscustomobject]@{ Id='HKCU';    Rotulo=(T 'item.37'); Tipo='Reg'; Caminho=$script:MarcadorChaveReg; Prova='ajuste de registro do Modulo 3' }
        [pscustomobject]@{ Id='LOCAL';   Rotulo=(T 'item.38');    Tipo='Arq'; Caminho=(Join-Path $env:LOCALAPPDATA 'KitSuporteRT\marcador.txt'); Prova='estado do Desfazer' }
        [pscustomobject]@{ Id='ROAMING'; Rotulo=(T 'item.39');  Tipo='Arq'; Caminho=(Join-Path $env:APPDATA 'KitSuporteRT\marcador.txt');      Prova='autorrecuperacao do Excel' }
    )
    try {
        $util = Join-Path $script:PastaClinica 'UTILITARIOS'
        if (Test-Path -LiteralPath $util) {
            $lista += [pscustomobject]@{ Id='DISCO'; Rotulo=(T 'item.40'); Tipo='Arq'; Caminho=(Join-Path $util 'KitSuporteRT_marcador.txt'); Prova='portais, POPs e o proprio kit' }
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

    $r = [pscustomobject]@{
        Caminho = "$env:USERPROFILE"; Tipo = 'Local'; Obrigatorio = $false; Temporario = $false
        Movel = $false; CaminhoMovel = ''; Filtro = ''; ApagaCache = $false; Motivos = @()
    }
    $sid = ''
    try { $sid = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value } catch { }

    try {
        if (Test-Path -LiteralPath (Join-Path $env:USERPROFILE 'NTUSER.MAN')) {
            $r.Obrigatorio = $true; $r.Tipo = 'Obrigatorio'
            $r.Motivos += 'NTUSER.MAN presente na raiz do perfil'
        }
    } catch { }

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
    Write-Titulo (T 'diagpersistencia.01')
    $d = Get-DiagnosticoPerfil

    Write-Log ((T 'diagpersistencia.02') -f $d.Caminho) 'DADO'
    Write-Log ((T 'diagtasy.13') -f $d.Tipo) 'DADO'
    Add-ItemInv 'Persistencia' 'Tipo de perfil' $d.Tipo $d.Caminho
    if ($d.Movel)  { Write-Log ((T 'diagpersistencia.03') -f $d.CaminhoMovel) 'DADO'; Add-ItemInv 'Persistencia' 'Perfil movel' $d.CaminhoMovel }
    if ($d.Filtro) { Add-ItemInv 'Persistencia' 'Filtro de escrita' $d.Filtro }
    foreach ($m in $d.Motivos) { Write-Log ('   . ' + $m) 'DADO' }

    if ($d.Obrigatorio) {
        Add-Achado 'CRITICO' (T 'diagpersistencia.04') (T 'diagpersistencia.05') (T 'diagpersistencia.06') 'Alto'
    } elseif ($d.Temporario) {
        Add-Achado 'CRITICO' (T 'diagpersistencia.07') (T 'diagpersistencia.08') (T 'diagpersistencia.06') 'Alto'
    } elseif ($d.ApagaCache) {
        Add-Achado 'ALERTA' (T 'diagpersistencia.09') (T 'diagpersistencia.10') (T 'diagpersistencia.06') 'Alto'
    } elseif ($d.Movel) {
        Add-Achado 'ALERTA' (T 'diagpersistencia.11') (T 'diagpersistencia.12') (T 'diagpersistencia.06') 'Medio'
    } else {
        Add-Achado 'OK' (T 'diagpersistencia.13') '' (T 'diagpersistencia.06')
    }

    if ($d.Filtro) {
        Add-Achado 'CRITICO' ((T 'diagpersistencia.14') -f $d.Filtro) (T 'diagpersistencia.15') (T 'diagpersistencia.06') 'Alto'
    }

    $sessao = Get-SessaoAtual
    $locais = Get-LocaisMarcador
    $anterior = ''
    foreach ($id in @('DISCO','ROAMING','LOCAL','HKCU')) {
        $l = $locais | Where-Object { $_.Id -eq $id }
        if ($l) { $v = Read-Marcador $l; if ($v) { $anterior = $v; break } }
    }

    if (-not $anterior) {
        Write-Log '' 'DADO'
        Write-Log (T 'diagpersistencia.16') 'DADO'
        Write-Log (T 'diagpersistencia.17') 'ACAO'
    } else {
        $partes = $anterior -split '\|'
        $logonAntes = $(if ($partes.Count -ge 3) { $partes[2] } else { '' })
        Write-Log '' 'DADO'
        if ($logonAntes -and $logonAntes -eq $sessao.Logon) {
            Write-Log ((T 'diagpersistencia.18') -f $partes[0]) 'DADO'
            Write-Log (T 'diagpersistencia.19') 'ACAO'
        } else {
            Write-Log ((T 'diagpersistencia.20') -f $partes[0]) 'DADO'
            $perdidos = @()
            foreach ($l in $locais) {
                $v = Read-Marcador $l
                if ($v -eq $anterior) {
                    Write-Log ((T 'diagpersistencia.21') -f $l.Rotulo, $l.Prova) 'OK'
                } else {
                    Write-Log ((T 'diagpersistencia.22') -f $l.Rotulo, $l.Prova) 'ALERTA'
                    $perdidos += $l
                }
            }
            if ($perdidos.Count -eq 0) {
                Add-Achado 'OK' (T 'diagpersistencia.23') '' (T 'diagpersistencia.06')
            } else {
                $ids = @($perdidos | Select-Object -ExpandProperty Id)
                if ($ids -contains 'DISCO') {
                    Add-Achado 'CRITICO' (T 'diagpersistencia.24') (T 'diagpersistencia.25') (T 'diagpersistencia.06') 'Alto'
                } else {
                    Add-Achado 'CRITICO' ((T 'diagpersistencia.26') -f $perdidos.Count, $locais.Count) ((T 'diagpersistencia.27') + (($perdidos | Select-Object -ExpandProperty Rotulo) -join '; ') + (T 'diagpersistencia.28')) (T 'diagpersistencia.06') 'Alto'
                }
            }
        }
    }

    $carimbo = '{0:yyyy-MM-dd HH:mm:ss}|{1}|{2}' -f (Get-Date), $sessao.Boot, $sessao.Logon
    $ok = 0
    foreach ($l in $locais) { if (Write-Marcador $l $carimbo) { $ok++ } }
    Write-Log ((T 'diagpersistencia.29') -f $ok, $locais.Count) 'DADO'
    if ($ok -lt $locais.Count) { Write-Log (T 'diagpersistencia.30') 'ALERTA' }
}

function Get-ArquivoRegistroEcossistema {

    try {
        foreach ($d in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)) {
            $f = Join-Path ("$($d.DeviceID)\" + $script:MarcaEcossistema) '_instalados.json'
            if (Test-Path -LiteralPath $f) { return $f }
        }
    } catch { }
    return ''
}

function Test-CaminhoLocalExiste {

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
    Write-Titulo (T 'invecossistema.01')
    $r = Get-AppsEcossistema

    if ($r.Estado -eq 'ausente') {
        Write-Log (T 'invecossistema.02') 'DADO'
        Add-ItemInv 'Ecossistema' 'Registro' 'ausente'
        return
    }
    Write-Log ((T 'invecossistema.03') -f $r.Arquivo) 'DADO'

    if ($r.Estado -ne 'ok') {
        $porque = switch ($r.Estado) {
            'ilegivel' { 'nao foi possivel ler o arquivo' }
            'vazio'    { 'o arquivo esta vazio' }
            default    { 'o conteudo nao e JSON valido' }
        }
        Add-Achado 'ALERTA' ((T 'invecossistema.04') -f $porque) (T 'invecossistema.05') (T 'invecossistema.06') 'Medio'
        Add-ItemInv 'Ecossistema' 'Registro' $r.Estado $r.Arquivo
        return
    }

    Write-Log ((T 'invecossistema.07') -f $r.Apps.Count) 'DADO'
    $divergentes = @()
    foreach ($a in $r.Apps) {
        $situacao = if ($a.Existe -eq $true) { 'pasta ok' } elseif ($null -eq $a.Existe) { 'nao verificavel (caminho de rede)' } else { 'PASTA NAO EXISTE' }
        Write-Log ((T 'invecossistema.08') -f $a.Projeto, $a.Versao, $a.Data, $situacao) 'DADO'
        Write-Log ((T 'invecossistema.09') -f $a.Caminho) 'DADO'
        Add-ItemInv 'Ecossistema' $a.Projeto ('v' + $a.Versao) ('{0} | {1} | {2}' -f $a.Data, $a.Caminho, $situacao)
        if ($a.Existe -eq $false) { $divergentes += $a }
    }

    if ($divergentes.Count -gt 0) {
        $det = ($divergentes | ForEach-Object { '{0} v{1} -> {2}' -f $_.Projeto, $_.Versao, $_.Caminho }) -join '; '
        Add-Achado 'ALERTA' ((T 'invecossistema.10') -f $divergentes.Count) ((T 'invecossistema.11') + $det + (T 'invecossistema.12')) (T 'invecossistema.06') 'Medio'
    }
    if ($r.SemRegistro.Count -gt 0) {
        Add-Achado 'ALERTA' ((T 'invecossistema.13') -f $r.SemRegistro.Count) ((T 'invecossistema.14') + ($r.SemRegistro -join ', ') + (T 'invecossistema.15')) (T 'invecossistema.06') 'Medio'
    }
    if ($divergentes.Count -eq 0 -and $r.SemRegistro.Count -eq 0) {
        Add-Achado 'OK' (T 'invecossistema.16') '' (T 'invecossistema.06')
    }
}

function Invoke-InventarioDiagnostico {
    Clear-SnapshotsOperacao
    $script:Achados.Clear()
    $script:ItensInv.Clear()
    Write-Titulo (T 'inventariodiagnost.01')
    Write-Log ((T 'inventariodiagnost.02') -f $env:COMPUTERNAME, $env:USERNAME, (Get-Date)) 'DADO'
    Write-Log (T 'inventariodiagnost.03') 'DADO'

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
        if ($script:Cancelar) { Write-Log (T 'inventariodiagnost.04') 'ALERTA'; break }
        $i++
        Set-Status ((T 'inventariodiagnost.05') -f $i, $etapas.Count, $e.Nome)
        try { & $e.Fn } catch { Write-Log ((T 'inventariodiagnost.06') -f $e.Nome, $_.Exception.Message) 'ALERTA' }
    }

    Write-Titulo (T 'inventariodiagnost.07')
    $crit = @($script:Achados | Where-Object { $_.Severidade -eq 'CRITICO' })
    $alt  = @($script:Achados | Where-Object { $_.Severidade -eq 'ALERTA' })

    if ($crit.Count -eq 0 -and $alt.Count -eq 0) {
        Write-Log (T 'inventariodiagnost.08') 'OK'
        Write-Log (T 'inventariodiagnost.09') 'DADO'
        Write-Log (T 'inventariodiagnost.10') 'ACAO'
    } else {
        Write-Log (T 'inventariodiagnost.11') 'DADO'
        $porArea = $script:Achados | Group-Object Categoria | ForEach-Object {
            [pscustomobject]@{ Area = $_.Name; Alto = @($_.Group | Where-Object { $_.Impacto -eq 'Alto' }).Count; Total = $_.Count }
        } | Sort-Object Alto, Total -Descending
        foreach ($a in $porArea) {
            $barra = ('#' * [Math]::Min(24, ($a.Alto * 4 + $a.Total)))
            Write-Log ((T 'inventariodiagnost.12') -f $a.Area, $barra, $a.Total, $a.Alto) 'DADO'
        }

        Write-Log ''
        Write-Log ((T 'inventariodiagnost.13') -f $crit.Count, $alt.Count)
        Write-Log ''
        $ordem = @()
        $ordem += @($crit | Where-Object { $_.Impacto -eq 'Alto' })
        $ordem += @($crit | Where-Object { $_.Impacto -ne 'Alto' })
        $ordem += @($alt  | Where-Object { $_.Impacto -eq 'Alto' })
        $ordem += @($alt  | Where-Object { $_.Impacto -ne 'Alto' })

        Write-Log (T 'inventariodiagnost.14') 'TITULO'
        $n = 0
        foreach ($a in $ordem) {
            $n++
            Write-Log ((T 'inventariodiagnost.15') -f $n, $a.Categoria, $a.Impacto, $a.Titulo) $a.Severidade
            if ($a.Recomendacao) { Write-Log ((T 'achado.01') + $a.Recomendacao) 'ACAO' }
        }
    }

    Write-Log ''
    Write-Log (T 'inventariodiagnost.16') 'ACAO'
    Write-Log (T 'inventariodiagnost.17') 'DADO'
    Write-Log (T 'inventariodiagnost.18') 'DADO'
    Save-Inventario
}

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

    $al += New-Alvo -Id 'TEMP' -Nome (T 'alvo.02') -Caminhos @($env:TEMP) -Obs (T 'obs.01')
    $al += New-Alvo -Id 'EDGE' -Nome (T 'alvo.03') -Fechar @('msedge') -Caminhos @(
        "$lad\Microsoft\Edge\User Data\*\Cache",
        "$lad\Microsoft\Edge\User Data\*\Code Cache",
        "$lad\Microsoft\Edge\User Data\*\GPUCache",
        "$lad\Microsoft\Edge\User Data\*\Service Worker\CacheStorage",
        "$lad\Microsoft\Edge\User Data\component_crx_cache",
        "$lad\Microsoft\Edge\User Data\GrShaderCache") -Obs (T 'obs.02')
    $al += New-Alvo -Id 'CHROME' -Nome (T 'alvo.04') -Fechar @('chrome') -Caminhos @(
        "$lad\Google\Chrome\User Data\*\Cache",
        "$lad\Google\Chrome\User Data\*\Code Cache",
        "$lad\Google\Chrome\User Data\*\GPUCache",
        "$lad\Google\Chrome\User Data\*\Service Worker\CacheStorage") -Obs (T 'obs.02')
    $al += New-Alvo -Id 'FIREFOX' -Nome (T 'alvo.05') -Fechar @('firefox') -Caminhos @(
        "$lad\Mozilla\Firefox\Profiles\*\cache2",
        "$lad\Mozilla\Firefox\Profiles\*\startupCache",
        "$lad\Mozilla\Firefox\Profiles\*\OfflineCache") -Obs (T 'obs.03')
    $al += New-Alvo -Id 'TEAMS1' -Nome (T 'alvo.06') -Fechar @('Teams') -Caminhos @(
        "$ad\Microsoft\Teams\Cache", "$ad\Microsoft\Teams\blob_storage", "$ad\Microsoft\Teams\databases",
        "$ad\Microsoft\Teams\GPUCache", "$ad\Microsoft\Teams\IndexedDB", "$ad\Microsoft\Teams\Local Storage",
        "$ad\Microsoft\Teams\tmp", "$ad\Microsoft\Teams\Code Cache") -Obs (T 'obs.04')
    $al += New-Alvo -Id 'TEAMS2' -Nome (T 'alvo.07') -Fechar @('ms-teams') -Caminhos @(
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Cache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Code Cache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\GPUCache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Service Worker\CacheStorage",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Service Worker\ScriptCache",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Cache Storage",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\PreviousVersions",
        "$lad\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Temp")
    $al += New-Alvo -Id 'JAVAWS' -Nome (T 'alvo.08') -Fechar @('javaw','java','jp2launcher') -Caminhos @(
        "$lad\Sun\Java\Deployment\cache", "$ad\Sun\Java\Deployment\cache",
        "$up\AppData\LocalLow\Sun\Java\Deployment\cache") -Obs (T 'obs.05')
    $al += New-Alvo -Id 'CENTCACHE' -Nome (T 'alvo.09') -Padrao $false -Fechar @('CentBrowser') -Caminhos @(
        "$lad\CentBrowser\User Data\*\Cache", "$lad\CentBrowser\User Data\*\Code Cache",
        "$lad\CentBrowser\User Data\*\GPUCache") -Obs (T 'obs.06')
    $al += New-Alvo -Id 'RDPCACHE' -Nome (T 'alvo.10') -Caminhos @(
        "$lad\Microsoft\Terminal Server Client\Cache") -Obs (T 'obs.07')
    $al += New-Alvo -Id 'CTEMP' -Nome ('Pasta C:\Temp (' + $txtPeriodo + ')') -DiasMin $dias -Caminhos @(
        (Join-Path $env:SystemDrive 'Temp')) -Obs (T 'obs.08')
    $al += New-Alvo -Id 'OPERAVIVALDI' -Nome (T 'alvo.11') -Caminhos @(
        "$lad\Opera Software\Opera Next\Cache", "$lad\Opera Software\Opera Stable\Cache",
        "$lad\Vivaldi\User Data\*\Cache")
    $al += New-Alvo -Id 'INETCOOKIES' -Nome (T 'alvo.12') -Padrao $false -Caminhos @(
        "$lad\Microsoft\Windows\INetCookies") -Obs (T 'obs.09')
    $al += New-Alvo -Id 'INETCACHE' -Nome (T 'alvo.13') -Caminhos @("$lad\Microsoft\Windows\INetCache")
    $al += New-Alvo -Id 'WER' -Nome (T 'alvo.14') -Caminhos @(
        "$lad\Microsoft\Windows\WER\ReportQueue", "$lad\Microsoft\Windows\WER\ReportArchive",
        "$ad\Microsoft\Windows\WER")
    $al += New-Alvo -Id 'DUMPS' -Nome (T 'alvo.15') -Caminhos @("$lad\CrashDumps")
    $al += New-Alvo -Id 'GPU' -Nome (T 'alvo.16') -Caminhos @(
        "$lad\D3DSCache", "$lad\NVIDIA\DXCache", "$lad\NVIDIA\GLCache",
        "$lad\NVIDIA Corporation\NV_Cache", "$lad\AMD\DxCache", "$lad\Intel\ShaderCache")
    $al += New-Alvo -Id 'ONEDRIVE' -Nome (T 'alvo.17') -Caminhos @("$lad\Microsoft\OneDrive\logs")
    $al += New-Alvo -Id 'TEAMSANTIGO' -Nome (T 'alvo.18') -Caminhos @(
        "$lad\Microsoft\Teams\previous", "$lad\Microsoft\Teams\packages", "$lad\SquirrelTemp")
    $al += New-Alvo -Id 'EDGEUPDATE' -Nome (T 'alvo.19') -Caminhos @(
        "$lad\Microsoft\EdgeUpdate\Download", "$lad\Google\Update\Download")
    $al += New-Alvo -Id 'ADOBE' -Nome (T 'alvo.20') -Fechar @('Acrobat','AcroRd32') -Caminhos @(
        "$lad\Adobe\Acrobat\DC\Cache", "$lad\Adobe\Acrobat\DC\ConnectorIcons", "$lad\Adobe\Color\ACEC")
    $al += New-Alvo -Id 'OFFICELOG' -Nome (T 'alvo.21') -Caminhos @(
        "$lad\Microsoft\Office\16.0\Telemetry", "$lad\Temp\Diagnostics", "$lad\Microsoft\Office\Logs")

    $al += New-Alvo -Id 'CITRIXICA' -Nome ('Arquivos .ica soltos em Downloads (' + $txtPeriodo + ')') -Modo 'ArquivosRaiz' -DiasMin $dias -Caminhos @(
        "$up\Downloads") -Obs (T 'obs.10')
    $al += New-Alvo -Id 'LIXEIRA' -Nome (T 'alvo.01') -Modo 'Lixeira' -Caminhos @()

    $al += New-Alvo -Id 'OFFICECACHE' -Nome (T 'alvo.22') -Padrao $false -Fechar @('excel','winword','powerpnt','outlook') -Caminhos @(
        "$lad\Microsoft\Office\16.0\OfficeFileCache") -Obs (T 'obs.11')
    $al += New-Alvo -Id 'MINIATURAS' -Nome (T 'alvo.23') -Modo 'ArquivosRaiz' -Caminhos @(
        "$lad\Microsoft\Windows\Explorer") -Obs (T 'obs.12')
    $al += New-Alvo -Id 'EXCELRECUP' -Nome ('Autorrecuperacao do Excel (' + $txtPeriodo + ')') -Modo 'ArquivosRaiz' -DiasMin $dias -Caminhos @(
        "$ad\Microsoft\Excel") -Obs ('Nao mexe na pasta XLSTART. Com "tudo" selecionado, apaga tambem a recuperacao de hoje.')
    $al += New-Alvo -Id 'RECENTES' -Nome (T 'alvo.24') -Padrao $false -Caminhos @(
        "$ad\Microsoft\Windows\Recent") -Obs (T 'obs.13')
    $al += New-Alvo -Id 'ORFAOS' -Nome ('Arquivos travados do Office na pasta clinica (' + $txtPeriodo + ')') -Modo 'ArquivosRaiz' -DiasMin $dias -Fechar @('excel','winword') -Caminhos @(
        $script:PastaClinica) -Obs (T 'obs.14')

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

            $antesFiltro = $itens.Count
            $itens = @($itens | Where-Object { -not (Test-StagingEcossistema $_.Name) })

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

    foreach ($a in (Get-AlvosLimpeza)) {
        if ($script:Cancelar) { break }
        if ($a.Modo -eq 'Lixeira') { continue }
        Set-Status ((T 'diagtasy.12') + $a.Nome + (T 'diagcitrix.19'))
        $b = Measure-Alvo -Alvo $a
        if ($b -le 0) { continue }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'LIMPAR'; Rotulo = $a.Nome; Valor = (Format-Bytes $b); Bytes = $b
            Marcar = [bool]$a.Padrao; Dados = $a; Nota = $a.Obs
        })
    }

    $lix = New-Alvo -Id 'LIXEIRA' -Nome (T 'alvo.25') -Modo 'Lixeira' -Caminhos @()
    Set-Status (T 'planolimpeza.01')
    $bl = Measure-Alvo -Alvo $lix
    [void]$plano.Add([pscustomobject]@{
        Grupo = 'LIXEIRA'; Rotulo = (T 'item.41'); Valor = (Format-Bytes $bl); Bytes = $bl
        Marcar = $true; Dados = $lix; Nota = (T 'nota.13')
    })

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
            Marcar = ($itens.Count -gt 0); Dados = $itens; Nota = (T 'nota.14')
        })
    }

    foreach ($e in (Get-ProcessosEncerraveis)) {
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'FECHAR'; Rotulo = ('{0} ({1} processo(s))' -f $e.Rotulo, $e.Qtd); Valor = (Format-Bytes $e.Memoria); Bytes = 0
            Marcar = $true; Dados = $e; Nota = $e.Nota
        })
    }

    foreach ($i in (Get-ItensInicializacao)) {
        if ($i.Classe -eq 'Protegido') { continue }
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'INICIAR'; Rotulo = ('{0} - {1}' -f $i.Nome, $i.Rotulo); Valor = ''; Bytes = 0
            Marcar = [bool]$i.Seguro; Dados = $i
            Nota = $(if ($i.Seguro) { 'Continua funcionando quando voce abrir pelo menu Iniciar.' } else { 'Desmarcado por seguranca: o kit nao reconheceu este item.' })
        })
    }

    foreach ($a in (Get-AjustesPendentes)) {
        [void]$plano.Add([pscustomobject]@{
            Grupo = 'AJUSTE'; Rotulo = $a.Rotulo; Valor = ''; Bytes = 0
            Marcar = $true; Dados = $a; Nota = $a.Nota
        })
    }

    foreach ($m in (Get-MapeamentosMortos)) {
        Add-Achado 'ALERTA' ((T 'planolimpeza.02') -f $m.Letra, $m.Destino) `
            ((T 'planolimpeza.03') -f $m.Letra) (T 'invredeimpressoras.05') 'Baixo'
    }
    foreach ($c in (Get-CredenciaisOrfas)) {
        Add-Achado 'ALERTA' ((T 'planolimpeza.04') -f $c.Servidor) `
            ((T 'planolimpeza.05') -f $c.Alvo) (T 'planolimpeza.06') 'Baixo'
    }

    return $plano
}

function Invoke-Modulo3Analise {
    Clear-SnapshotsOperacao
    Write-Titulo (T 'modulo3analise.01')
    Write-Log (T 'modulo3analise.02') 'DADO'

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

    $bytes = ($plano | Where-Object { $_.Marcar -and $_.Grupo -match 'LIMPAR|LIXEIRA|BAIXADOS' } | Measure-Object Bytes -Sum).Sum
    $mem   = ($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'FECHAR' } | ForEach-Object { $_.Dados.Memoria } | Measure-Object -Sum).Sum
    $ini   = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'INICIAR' }).Count
    $aju   = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'AJUSTE' }).Count
    $red   = @($plano | Where-Object { $_.Marcar -and $_.Grupo -eq 'REDE' }).Count
    $naoRec = @($plano | Where-Object { $_.Grupo -eq 'INICIAR' -and -not $_.Marcar })

    Write-Log '' 'DADO'
    Write-Log (T 'modulo3analise.03') 'TITULO'
    Write-Log ((T 'modulo3analise.04') -f (Format-Bytes $bytes)) 'ACAO'
    Write-Log ((T 'modulo3analise.05') -f (Format-Bytes $mem)) 'ACAO'
    Write-Log ((T 'modulo3analise.06') -f $ini) 'ACAO'
    Write-Log ((T 'modulo3analise.07') -f $aju) 'ACAO'
    if ($red -gt 0) { Write-Log ((T 'modulo3analise.08') -f $red) 'ACAO' }
    if ($naoRec.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ((T 'modulo3analise.09') -f $naoRec.Count) 'DADO'
        foreach ($n in $naoRec) { Write-Log ('   ' + $n.Rotulo) 'DADO' }
        Write-Log (T 'modulo3analise.10') 'DADO'
    }
    Write-Log '' 'DADO'
    Write-Log (T 'modulo3analise.11') 'ACAO'
    $script:PainelGrandes.Visible = $false
    $script:PainelSessao.Visible = $false
    $script:PainelLimpeza.Visible = $true
}

function Invoke-Modulo3Limpeza {
    if (-not $script:PlanoAtual -or $script:PlanoAtual.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((T 'modulo3limpeza.01'), (T 'modulo3limpeza.02'), 'OK', (T 'modulo3limpeza.03')) | Out-Null
        return
    }

    $marcados = @()
    for ($i = 0; $i -lt $script:ListaLimpeza.Items.Count; $i++) {
        if ($script:ListaLimpeza.GetItemChecked($i)) { $marcados += $script:PlanoAtual[$i] }
    }
    if ($marcados.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((T 'modulo3limpeza.04'), (T 'modulo3limpeza.02'), 'OK', (T 'modulo3limpeza.03')) | Out-Null
        return
    }

    $bytes = ($marcados | Where-Object { $_.Grupo -match 'LIMPAR|LIXEIRA|BAIXADOS' } | Measure-Object Bytes -Sum).Sum
    $nFechar = @($marcados | Where-Object { $_.Grupo -eq 'FECHAR' }).Count
    $nIni    = @($marcados | Where-Object { $_.Grupo -eq 'INICIAR' }).Count
    $nAju    = @($marcados | Where-Object { $_.Grupo -eq 'AJUSTE' }).Count

    $momento = Get-MomentoSessao
    $aviso = ''
    if ($momento.Veredicto -ne 'RECEM') {
        $aviso = $momento.Texto + "`r`n"
        if ($momento.Detalhe) { $aviso += $momento.Detalhe + "`r`n" }
        $aviso += "`r`n"
    }

    if ($script:ModoPlano -eq 'SESSAO') {
        $nSes = @($marcados | Where-Object { $_.Grupo -eq 'SESSAO' }).Count
        $nTar = @($marcados | Where-Object { $_.Grupo -eq 'TAREFA' }).Count
        $memSes = ($marcados | Where-Object { $_.Grupo -eq 'SESSAO' } | Measure-Object Bytes -Sum).Sum
        $msg = $aviso + ("Otimizar a sessao de agora com {0} itens:`r`n`r`n" -f $marcados.Count)
        $msg += " · Encerrar {0} programa(s), devolvendo cerca de {1}`r`n" -f $nSes, (Format-Bytes $memSes)
        $msg += " · Parar {0} tarefa(s) agendada(s) em execucao`r`n" -f $nTar
        $msg += " · Ajustar prioridade de CPU e compactar memoria`r`n`r`n"
        $msg += "Citrix, navegador do prontuario, navegador web e Office nao sao tocados.`r`n"
        $msg += "Nada e desinstalado: o proximo logon devolve tudo ao normal.`r`n`r`nContinuar?"
    } else {
        $msg = $aviso + ("Vao ser aplicados {0} itens:`r`n`r`n" -f $marcados.Count)
        $msg += " · Liberar cerca de {0} em disco`r`n" -f (Format-Bytes $bytes)
        $msg += " · Encerrar {0} programa(s) dispensavel(is)`r`n" -f $nFechar
        $msg += " · Tirar {0} item(ns) da inicializacao`r`n" -f $nIni
        $msg += " · Aplicar {0} ajuste(s) de desempenho`r`n`r`n" -f $nAju

        $msg += "A Lixeira e esvaziada primeiro. O que sair de Downloads vai para ela depois, e continua recuperavel.`r`n"
        $msg += "Fora os arquivos apagados, nada aqui e definitivo: o botao Desfazer e o proximo logon devolvem o resto.`r`n`r`nContinuar?"
    }

    if ([System.Windows.Forms.MessageBox]::Show($msg, (T 'modulo3limpeza.05'), (T 'modulo3limpeza.06'), (T 'modulo3limpeza.07')) -ne 'Yes') {
        Write-Log (T 'modulo3limpeza.08') 'ALERTA'
        return
    }

    Clear-SnapshotsOperacao
    [void](Get-EmUsoOperacao -Renovar)

    Set-Ocupado $true
    $script:Cancelar = $false
    $script:SessaoEncerrados = @()
    $script:ExplorerAtualizado = $false
$script:LogonSegundos      = 0
    Write-Titulo $(if ($script:ModoPlano -eq 'SESSAO') { (T 'modulo3limpeza.10') } else { (T 'modulo3limpeza.11') })

    $liberado = 0.0
    $memoria  = 0.0
    $desativados = @()
    $n = 0

    try {

        $ordem = @('SESSAO', 'TAREFA', 'FECHAR', 'LIMPAR', 'LIXEIRA', 'BAIXADOS', 'INICIAR', 'AJUSTE', 'PRIORIDADE', 'MEMORIA')
        foreach ($g in $ordem) {
            $doGrupo = @($marcados | Where-Object { $_.Grupo -eq $g })
            if ($doGrupo.Count -eq 0) { continue }

            switch ($g) {
                'FECHAR'   { Write-Log '' ; Write-Log (T 'modulo3limpeza.12') 'TITULO' }
                'LIMPAR'   { Write-Log '' ; Write-Log (T 'modulo3limpeza.13') 'TITULO' }
                'LIXEIRA'  { Write-Log '' ; Write-Log (T 'modulo3limpeza.14') 'TITULO' }
                'BAIXADOS' { Write-Log '' ; Write-Log (T 'modulo3limpeza.15') 'TITULO' }
                'INICIAR'  { Write-Log '' ; Write-Log (T 'modulo3limpeza.16') 'TITULO' }
                'AJUSTE'   { Write-Log '' ; Write-Log (T 'modulo3limpeza.17') 'TITULO' }
                'SESSAO'     { Write-Log '' ; Write-Log (T 'modulo3limpeza.18') 'TITULO' }
                'TAREFA'     { Write-Log '' ; Write-Log (T 'modulo3limpeza.19') 'TITULO' }
                'PRIORIDADE' { Write-Log '' ; Write-Log (T 'modulo3limpeza.20') 'TITULO' }
                'MEMORIA'    { Write-Log '' ; Write-Log (T 'modulo3limpeza.21') 'TITULO' }
            }

            foreach ($item in $doGrupo) {
                if ($script:Cancelar) { Write-Log (T 'modulo3limpeza.22') 'ALERTA'; break }
                $n++
                Set-Status ((T 'modulo3limpeza.23') -f $n, $marcados.Count, $item.Rotulo)

                switch ($item.Grupo) {

                    'FECHAR' {
                        $antes = 0.0
                        $ok = 0
                        foreach ($id in $item.Dados.Ids) {
                            try {
                                $p = Get-Process -Id $id -ErrorAction Stop
                                if (Test-Padrao $p.ProcessName $script:Protegidos) { continue }
                                $antes += $p.WorkingSet64
                                if ($p.MainWindowHandle -ne 0) { [void]$p.CloseMainWindow(); Start-Sleep -Milliseconds 500 }
                                if (-not $p.HasExited) { Stop-Process -Id $id -Force -ErrorAction Stop }
                                $ok++
                            } catch { }
                        }
                        $memoria += $antes
                        if ($ok -gt 0) {
                            $script:SessaoEncerrados += @($item.Dados.Nomes)
                            Write-Log ((T 'modulo3limpeza.24') -f $item.Dados.Rotulo, $ok, (Format-Bytes $antes)) 'OK'
                        }
                        else { Write-Log ((T 'modulo3limpeza.25') -f $item.Dados.Rotulo) 'DADO' }
                    }

                    'LIMPAR' {
                        $r = Clear-Alvo -Alvo $item.Dados
                        $liberado += $r.Bytes
                        if ($r.Erro) { Write-Log ((T 'modulo3limpeza.26') -f $item.Rotulo, $r.Erro) 'ALERTA' }
                        elseif ($r.Bytes -le 0) { Write-Log ((T 'modulo3limpeza.27') -f $item.Rotulo) 'DADO' }
                        else {
                            $partes = @()
                            if ($r.Bloqueados -gt 0)  { $partes += ('{0} em uso' -f $r.Bloqueados) }
                            if ($r.Preservados -gt 0) { $partes += ('{0} preservado(s): trabalho do ecossistema ou mexido agora' -f $r.Preservados) }
                            $extra = if ($partes.Count -gt 0) { (' (' + ($partes -join '; ') + ')') } else { '' }
                            Write-Log ((T 'modulo3limpeza.28') -f $item.Rotulo, (Format-Bytes $r.Bytes), $extra) 'OK'
                        }
                    }

                    'LIXEIRA' {
                        $r = Clear-Alvo -Alvo $item.Dados
                        $liberado += $r.Bytes
                        if ($r.Erro) { Write-Log ((T 'modulo3limpeza.29') -f $r.Erro) 'DADO' }
                        elseif ($r.Bytes -le 0) { Write-Log (T 'modulo3limpeza.30') 'DADO' }
                        else { Write-Log ((T 'modulo3limpeza.31') -f (Format-Bytes $r.Bytes)) 'OK' }
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
                        Write-Log ((T 'modulo3limpeza.32') -f $qtd, (Format-Bytes $lib)) 'OK'
                        if ($sumidos -gt 0) { Write-Log ((T 'modulo3limpeza.33') -f $sumidos) 'DADO' }
                        if ($falhas.Count -gt 0) {
                            Write-Log ((T 'modulo3limpeza.34') -f $falhas.Count) 'ALERTA'
                            foreach ($n in ($falhas | Select-Object -First 5)) { Write-Log ('   ' + $n) 'DADO' }
                            if ($falhas.Count -gt 5) { Write-Log ((T 'modulo3limpeza.35') -f ($falhas.Count - 5)) 'DADO' }
                        }
                        Write-Log (T 'modulo3limpeza.36') 'ACAO'
                    }

                    'INICIAR' {
                        if ($item.Dados.Nome -match 'OneDrive' -and (Test-PastasNoOneDrive)) {
                            Write-Log (T 'modulo3limpeza.37') 'ALERTA'
                            Write-Log (T 'modulo3limpeza.38') 'ALERTA'
                        }
                        if (Disable-ItemInicializacao -Nome $item.Dados.Nome -Tipo $item.Dados.Tipo) {
                            Write-Log ((T 'modulo3limpeza.39') -f $item.Dados.Nome) 'OK'
                            $desativados += @{ Nome = $item.Dados.Nome; Tipo = $item.Dados.Tipo }
                        } else {
                            Write-Log ((T 'modulo3limpeza.40') -f $item.Dados.Nome) 'ALERTA'
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
                        Write-Log ((T 'modulo3limpeza.41') -f $item.Rotulo) 'OK'
                    }

                    'SESSAO' {
                        Set-Status ((T 'modulo3limpeza.58') -f $item.Dados.Nome)
                        $liberou = 0.0
                        $ok = 0
                        foreach ($id in $item.Dados.Ids) {
                            try {
                                $pr = Get-Process -Id $id -ErrorAction Stop
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
                            Write-Log ((T 'modulo3limpeza.24') -f $item.Dados.Nome, $ok, (Format-Bytes $liberou)) 'OK'
                        }
                        else { Write-Log ((T 'modulo3limpeza.59') -f $item.Dados.Nome) 'DADO' }
                    }

                    'TAREFA' {
                        try {
                            Stop-ScheduledTask -TaskName $item.Dados.Nome -TaskPath $item.Dados.Caminho -ErrorAction Stop
                            Write-Log ((T 'modulo3limpeza.60') -f $item.Dados.Caminho, $item.Dados.Nome) 'OK'
                        } catch {
                            Write-Log ((T 'modulo3limpeza.61') -f $item.Dados.Nome) 'DADO'
                        }
                    }

                    'PRIORIDADE' { Invoke-AcaoPrioridade }

                    'MEMORIA'    { Invoke-AcaoMemoria }

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

        Write-Titulo (T 'modulo3limpeza.42')
        if ($script:ModoPlano -eq 'SESSAO') {
            Write-Log ((T 'modulo3limpeza.43') -f (Format-Bytes $memoria)) 'OK'
            if ($script:UltimoGanhoMemoria -gt 0) {
                Write-Log ((T 'modulo3limpeza.44') -f (Format-Bytes $script:UltimoGanhoMemoria)) 'OK'
                Write-Log ((T 'modulo3limpeza.45') -f (Format-Bytes ($memoria + $script:UltimoGanhoMemoria))) 'OK'
            }
            try {
                $so = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
                $pct = [Math]::Round(((($so.TotalVisibleMemorySize - $so.FreePhysicalMemory) / $so.TotalVisibleMemorySize) * 100), 1)
                Write-Log ((T 'modulo3limpeza.46') -f $pct, (Format-Bytes ($so.FreePhysicalMemory * 1KB))) 'DADO'
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
            Write-Log (T 'modulo3limpeza.47') 'DADO'
            Write-Log (T 'modulo3limpeza.48') 'ACAO'
            Show-StatusSessao -Estado 'OTIMIZADA'
            return
        }
        if ($script:SessaoEncerrados.Count -gt 0) {
            $eL = Read-Estado
            $antL = @()
            if ($eL.ContainsKey('sessao_encerrados')) { $antL = @($eL['sessao_encerrados']) }
            Save-Estado 'sessao_encerrados' (@($antL + $script:SessaoEncerrados) | Select-Object -Unique)
        }
        Write-Log ((T 'modulo3limpeza.49') -f (Format-Bytes $liberado)) 'OK'
        Write-Log ((T 'modulo3limpeza.43') -f (Format-Bytes $memoria)) 'OK'
        Write-Log ((T 'modulo3analise.06') -f $desativados.Count) 'OK'
        try {
            $c = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
            Write-Log ((T 'modulo3limpeza.50') -f (Format-Bytes $c.FreeSpace), [Math]::Round(($c.FreeSpace / $c.Size) * 100, 1)) 'DADO'
        } catch { }
        Write-Log ''
        Write-Log (T 'modulo3limpeza.51') 'DADO'

        $r = Show-Escolha -Titulo (T 'diagsistemaacionav.04') `
            -Mensagem (T 'modulo3limpeza.52') `
            -Opcoes @((T 'modulo3limpeza.53'), (T 'modulo3limpeza.54'))
        if ($r -eq 0) {
            Write-Log (T 'modulo3limpeza.55') 'ALERTA'
            & shutdown /r /t 15 /c "Reinicio solicitado pelo Kit de Suporte" | Out-Null
        }
    } catch {
        Write-Log ((T 'modulo3limpeza.56') -f $_.Exception.Message) 'CRITICO'
    } finally {
        Set-Ocupado $false
        Set-Status (T 'modulo3limpeza.57')
        $script:Cancelar = $false
    }
}

function Restart-Computador {
    $msg = "O computador sera reiniciado agora.`r`n`r`nSalve os arquivos abertos antes de continuar.`r`n`r`nReiniciar (e nao Desligar) e o que realmente limpa a memoria.`r`n`r`nConfirmar?"
    if ([System.Windows.Forms.MessageBox]::Show($msg, (T 'restartcomputador.01'), (T 'modulo3limpeza.06'), (T 'restartcomputador.02')) -eq 'Yes') {
        Write-Log (T 'restartcomputador.03') 'ALERTA'

        $saidaSd = (& shutdown /r /t 10 /c "Reinicio solicitado pelo Kit de Suporte" 2>&1 | Out-String).Trim()
        if ($LASTEXITCODE -eq 0) {
            Write-Log (T 'restartcomputador.04') 'DADO'
        } else {
            Write-Log ((T 'restartcomputador.05') -f $LASTEXITCODE) 'ALERTA'
            if ($saidaSd) { Write-Log ('   ' + (@($saidaSd -split "`r?`n")[0])) 'DADO' }
            Write-Log (T 'restartcomputador.06') 'ACAO'
        }
    }
}

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

function Show-Escolha {
    param([string]$Titulo, [string]$Mensagem, [string[]]$Opcoes)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Titulo; $f.FormBorderStyle = 'FixedDialog'; $f.StartPosition = 'CenterParent'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.BackColor = $script:Cor.Painel; $f.ForeColor = $script:Cor.Texto
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

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

function Invoke-DiagDownloads {
    Write-Titulo (T 'diagdownloads.01')
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (-not (Test-Path -LiteralPath $dl)) { Write-Log (T 'diagdownloads.02') 'DADO'; return }

    try {
        $arq = @(Get-ChildItem -LiteralPath $dl -File -Force -Recurse -ErrorAction SilentlyContinue)
        if ($arq.Count -eq 0) { Add-Achado 'OK' (T 'diagdownloads.03') '' (T 'diagespaco.03'); return }

        $total  = ($arq | Measure-Object Length -Sum).Sum
        $velhos = @($arq | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) })
        $bVelho = ($velhos | Measure-Object Length -Sum).Sum

        Write-Log ((T 'diagdownloads.04') -f $arq.Count, (Format-Bytes $total)) 'DADO'
        Write-Log ((T 'diagdownloads.05') -f $velhos.Count, (Format-Bytes $bVelho)) 'DADO'

        $porTipo = $arq | Group-Object Extension | ForEach-Object {
            [pscustomobject]@{ Tipo = $_.Name; Qtd = $_.Count; Bytes = ($_.Group | Measure-Object Length -Sum).Sum }
        } | Sort-Object Bytes -Descending | Select-Object -First 8
        foreach ($t in $porTipo) {
            Write-Log ((T 'diagdownloads.06') -f $(if ($t.Tipo) { $t.Tipo } else { (T 'diagdownloads.07') }), (Format-Bytes $t.Bytes), $t.Qtd) 'DADO'
        }

        Write-Log (T 'diagdownloads.08') 'DADO'
        foreach ($f in ($arq | Sort-Object Length -Descending | Select-Object -First 10)) {
            Write-Log ((T 'diagdownloads.09') -f (Format-Bytes $f.Length), $f.Name, $f.LastWriteTime) 'DADO'
        }

        if ($total -gt 5GB -or $bVelho -gt 2GB) {
            Add-Achado 'ALERTA' ((T 'diagdownloads.10') -f (Format-Bytes $total), (Format-Bytes $bVelho)) (T 'diagdownloads.11') (T 'diagespaco.03') 'Alto'
        } else {
            Add-Achado 'OK' ((T 'diagdownloads.12') -f (Format-Bytes $total)) '' (T 'diagespaco.03')
        }
    } catch { }
}

function Invoke-DiagNuvem {
    Write-Titulo (T 'diagnuvem.01')

    $od = @(Get-Process OneDrive -ErrorAction SilentlyContinue)
    if ($od.Count -gt 0) {
        $mem = ($od | Measure-Object WorkingSet64 -Sum).Sum
        Write-Log ((T 'diagnuvem.02') -f (Format-Bytes $mem)) 'DADO'
        $pastaOD = Get-ValorReg 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' 'UserFolder'
        if (-not $pastaOD) { $pastaOD = $env:OneDrive }
        if ($pastaOD -and (Test-Path -LiteralPath $pastaOD)) {
            $t = Get-TamanhoPasta -Caminho $pastaOD -TimeoutSeg 90
            Write-Log ((T 'diagnuvem.03') -f $pastaOD, (Format-Bytes $t)) 'DADO'
            if ($t -gt 20GB) {
                Add-Achado 'ALERTA' ((T 'diagnuvem.04') -f (Format-Bytes $t)) (T 'diagnuvem.05') (T 'diagnuvem.06') 'Alto'
            }
        }
        if ($mem -gt 500MB) {
            Add-Achado 'ALERTA' ((T 'diagnuvem.07') -f (Format-Bytes $mem)) (T 'diagnuvem.08') (T 'diagnuvem.06') 'Medio'
        }
    } else {
        Write-Log (T 'diagnuvem.09') 'DADO'
    }

    $tm = @(Get-Process -Name 'Teams', 'ms-teams', 'msteams' -ErrorAction SilentlyContinue)
    if ($tm.Count -gt 0) {
        $mem = ($tm | Measure-Object WorkingSet64 -Sum).Sum
        Write-Log ((T 'diagnuvem.10') -f (Format-Bytes $mem), $tm.Count) 'DADO'
        if ($mem -gt 1GB) {
            Add-Achado 'ALERTA' ((T 'diagnuvem.11') -f (Format-Bytes $mem)) (T 'diagnuvem.12') (T 'diagnuvem.06') 'Alto'
        }
    } else {
        Write-Log (T 'diagnuvem.13') 'DADO'
    }

    $ant = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\previous'
    if (Test-Path -LiteralPath $ant) {
        $t = Get-TamanhoPasta -Caminho $ant -TimeoutSeg 45
        if ($t -gt 200MB) {
            Add-Achado 'ALERTA' ((T 'diagnuvem.14') -f (Format-Bytes $t)) (T 'diagnuvem.15') (T 'diagnuvem.06') 'Medio'
        }
    }
}

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

function Restore-Otimizacoes {
    Write-Titulo (T 'otimizacoes.01')
    $e = Read-Estado
    if ($e.Count -eq 0) { Write-Log (T 'otimizacoes.02') 'OK'; return }

    $r = Show-Escolha -Titulo (T 'otimizacoes.03') `
        -Mensagem (T 'otimizacoes.04') `
        -Opcoes @((T 'otimizacoes.05'), (T 'otimizacoes.06'))
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
                Write-Log ((T 'otimizacoes.07') -f $a.Nome) 'DADO'
            } catch { }
        }
    }

    if ($e.ContainsKey('inicializacao_desativada')) {
        foreach ($i in $e['inicializacao_desativada']) {
            if (Enable-ItemInicializacao -Nome $i.Nome -Tipo $i.Tipo) { Write-Log ((T 'otimizacoes.08') -f $i.Nome) 'DADO' }
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
        Write-Log ((T 'otimizacoes.09') -f $app) 'DADO'
    }

    if ($e.ContainsKey('zonas_citrix')) {
        foreach ($k in $e['zonas_citrix']) {
            try { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue; Write-Log ((T 'otimizacoes.10') -f $k) 'DADO' } catch { }
        }
    }

    Remove-Item -LiteralPath (Get-ArquivoEstado) -Force -ErrorAction SilentlyContinue
    Write-Log (T 'otimizacoes.11') 'OK'
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
                Write-Log ((T 'ajusteregistro.01') -f $a.Nome) 'ALERTA'
            } else {
                Write-Log ((T 'ajusteregistro.02') -f $a.Nome) 'ALERTA'
            }
        }
    }
    if ($bloqueados -gt 0) {
        Write-Log ((T 'ajusteregistro.03') -f $bloqueados, $Ajustes.Count) 'DADO'
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
        Write-Log (T 'explorer.01') 'OK'
    } catch {
        Write-Log (T 'explorer.02') 'DADO'
        Write-Log (T 'explorer.03') 'DADO'
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

function Get-VeredictoPasta {
    param([string]$Caminho)
    if (Test-DentroDaPasta -Caminho $Caminho -Pasta $script:PastaClinica)        { return @{ V = 'NAO APAGAR'; M = 'Pasta clinica. Nunca apague.' } }
    if (Test-Padrao $Caminho $script:RaizesClinicas) { return @{ V = 'NAO APAGAR'; M = 'Dado de sistema clinico.' } }
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

    $raizes = @($env:USERPROFILE, $env:LOCALAPPDATA, $env:APPDATA, $script:PastaClinica, ($env:SystemDrive + '\'))
    foreach ($r in $raizes) {
        if (-not (Test-Path -LiteralPath $r)) { continue }
        try {
            $lista += @(Get-ChildItem -LiteralPath $r -Directory -Force -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -notmatch '^\$|^System Volume Information$|^Recovery$|^Windows$|^Program Files|^ProgramData$|^Users$|^PerfLogs$|^Intel$|^Config\.Msi$|^AppData$' } |
                        Select-Object -ExpandProperty FullName)
        } catch { }
    }
    return @($lista | Where-Object { $_ -like ($env:SystemDrive + '\*') } | Select-Object -Unique)
}

function Invoke-ArquivosGrandes {
    Clear-SnapshotsOperacao
    Write-Titulo (T 'arquivosgrandes.01')
    Write-Log (T 'arquivosgrandes.02') 'DADO'
    Write-Log (T 'arquivosgrandes.03') 'DADO'
    Write-Log (T 'arquivosgrandes.04') 'DADO'

    try {
        foreach ($pf in (Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue)) {
            $letra = ($pf.Name.Substring(0, 2)).ToUpper()
            if ($letra -ne 'C:') {
                $d = Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='" + $letra + "'") -ErrorAction SilentlyContinue
                if ($d -and $d.Size) {
                    $pct = [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
                    Write-Log ((T 'arquivosgrandes.05') -f $letra, $pct) 'DADO'
                    if ($pct -lt 10) {
                        Add-Achado 'ALERTA' ((T 'arquivosgrandes.06') -f $letra, $pct) (T 'arquivosgrandes.07') (T 'diagespaco.03') 'Alto'
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

    Set-Status (T 'arquivosgrandes.08')
    $raizesArq = @($env:USERPROFILE, $script:PastaClinica, (Join-Path $env:SystemDrive 'Temp'), (Join-Path $env:SystemDrive 'Users\Public')) |
                 Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique
    $arqs = @()
    foreach ($r in $raizesArq) {
        if ($script:Cancelar) { break }
        Set-Status ((T 'arquivosgrandes.09') + $r)
        $arqs += Get-MaioresArquivos -Raiz $r -Top 30 -TimeoutSeg 90 -MinimoMB $script:Lim.ArquivoGrandeMB
    }
    $arqs = @($arqs | Sort-Object Length -Descending | Select-Object -First 10)

    Write-Log '' 'DADO'
    Write-Log (T 'arquivosgrandes.10') 'TITULO'
    if ($arqs.Count -eq 0) {
        Write-Log ((T 'arquivosgrandes.11') -f $script:Lim.ArquivoGrandeMB) 'DADO'
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
        Write-Log ((T 'arquivosgrandes.12') -f $n, $ver, (Format-Bytes $f.Length), $f.FullName) $nivel
        Write-Log ('    ' + $v.M) 'DADO'
        $script:GrandesItens += [pscustomobject]@{ Tipo = 'Arquivo'; Caminho = $f.FullName; Bytes = $f.Length; Veredicto = $ver; Motivo = $v.M }
        [void]$script:ListaGrandes.Items.Add(('[{0,-15}] {1,10}  {2}' -f $ver, (Format-Bytes $f.Length), $f.FullName))
        Add-ItemInv 'Arquivo grande' $f.Name (Format-Bytes $f.Length) ($ver + ' - ' + $f.FullName)
    }

    Write-Log '' 'DADO'
    Write-Log (T 'arquivosgrandes.13') 'TITULO'
    $cands = Get-CandidatosPastas
    $medidas = @()
    $i = 0
    foreach ($c in $cands) {
        if ($script:Cancelar) { break }
        $i++
        Set-Status ((T 'arquivosgrandes.14') -f $i, $cands.Count, (Split-Path $c -Leaf))
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
        Write-Log ((T 'arquivosgrandes.12') -f $n, $ver, (Format-Bytes $m.Bytes), $m.Caminho) $nivel
        Write-Log ('    ' + $v.M) 'DADO'
        $script:GrandesItens += [pscustomobject]@{ Tipo = 'Pasta'; Caminho = $m.Caminho; Bytes = $m.Bytes; Veredicto = $ver; Motivo = $v.M }
        [void]$script:ListaGrandes.Items.Add(('[{0,-15}] {1,10}  {2}' -f $ver, (Format-Bytes $m.Bytes), $m.Caminho))
        Add-ItemInv 'Pasta grande' (Split-Path $m.Caminho -Leaf) (Format-Bytes $m.Bytes) ($v.V + ' - ' + $m.Caminho)
    }

    $pode  = @($script:GrandesItens | Where-Object { $_.Veredicto -eq 'PODE APAGAR' })
    $aten  = @($script:GrandesItens | Where-Object { $_.Veredicto -eq 'ATENCAO-AVALIAR' })
    Write-Log '' 'DADO'
    Write-Log ((T 'arquivosgrandes.15') -f (Format-Bytes (($pode | Measure-Object Bytes -Sum).Sum)), $pode.Count) 'DADO'
    Write-Log ((T 'arquivosgrandes.16') -f (Format-Bytes (($aten | Measure-Object Bytes -Sum).Sum)), $aten.Count) 'DADO'
    Write-Log (T 'arquivosgrandes.17') 'DADO'
    Write-Log (T 'arquivosgrandes.18') 'DADO'
    Write-Log (T 'arquivosgrandes.19') 'ACAO'
    $script:PainelLimpeza.Visible = $false
    $script:PainelSessao.Visible = $false
    $script:PainelGrandes.Visible = $true
}

function Open-ItemGrande {
    $idx = $script:ListaGrandes.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:GrandesItens.Count) {
        [System.Windows.Forms.MessageBox]::Show((T 'openitemgrande.01'), (T 'openitemgrande.02'), 'OK', (T 'modulo3limpeza.03')) | Out-Null
        return
    }
    $item = $script:GrandesItens[$idx]
    if (-not (Test-Path -LiteralPath $item.Caminho)) {
        Write-Log ((T 'openitemgrande.03') -f $item.Caminho) 'ALERTA'
        return
    }
    try {
        if ($item.Tipo -eq 'Pasta') {
            Start-Process explorer.exe -ArgumentList ('"{0}"' -f $item.Caminho)
        } else {
            Start-Process explorer.exe -ArgumentList ('/select,"{0}"' -f $item.Caminho)
        }
        Write-Log ((T 'openitemgrande.04') -f $item.Caminho) 'DADO'
    } catch {
        Write-Log ((T 'openitemgrande.05') -f $_.Exception.Message) 'ALERTA'
    }
}

$script:SessaoAudio = 'RAVBg|RAVCpl|RtkNGUI|RtkAudUService|RtHDVBg|RtHDVCpl|RtlUpd|RealtekAudio' +
    '|WavesSvc|WavesSysSvc|WavesAudio|MaxxAudio' +
    '|Nahimic|A-Volute|AudioCenter' +
    '|DolbyDAX|DAX3API|DolbyAudio|DTSAPO|DTSAudio|SonicStudio|SonicSuite|SmartAudio|CxAudMsg|CxUIU' +
    '|LogiOptions|LogiLDA|Logitech|iCUE|SteelSeries|CorsairService' +
    '|WebcamService|CameraHelper|YourPhoneCamera' +
    '|fsquirt|BTTray|BTStackServer|BluetoothUserService|IntelBluetooth|btplayerctrl|BluetoothHeadset|BtwRSupportService'

$script:SessaoEnfeites = 'Widgets|WidgetService|SearchHost|SearchApp|SearchUI' +
    '|StartMenuExperienceHost|ShellExperienceHost|ShellHost|TaskViewHost|MultitaskingViewHost|TextInputHost' +
    '|FeedbackHub|PilotshubApp|GetHelp|Windows\.Feedback' +
    '|jusched|jucheck|JavaUpdate|JavaCheck|OneDriveStandaloneUpdater|GoogleUpdate|MicrosoftEdgeUpdate|EdgeUpdate' +
    '|SCNotification|UserOOBEBroker|CompPkgSrv|WindowsInternal|LockApp|SystemSettingsBroker'

$script:SessaoRemoto = 'TeamViewer|tv_w32|tv_x64|uvnc|winvnc|vncserver|tvnserver|ScreenConnect|AnyDesk|RustDesk' +
    '|CmRcService|CmRcViewer|RcAgent|DameWare|BeyondTrust|bomgar|LogMeIn|LMIGuardian|Splashtop|ZohoAssist' +
    '|quickassist|^msra$|^mstsc$|RdpClip|rdpinit'

$script:AppsDeTrabalho =
    'ARIA|Eclipse|Varian|MOSAIQ|IMPAC|Monaco|Focal|RayStation|RayCare|MIM|Velocity' +
    '|Pinnacle|Oncentra|XiO|iPlan|Brainlab|Precision|TomoTherapy|Limbus|AutoContour' +
    '|Accuray|Elekta|RaySearch' +
    '|Vitrea|Vsp|^VI\.|Mirada|Medis|4DM|Corridor|NeuroQ|Olea|TomTec|Onis|Digitalcore' +
    '|RadiAnt|Weasis|MicroDicom|Horos|dicom|PACS|Sectra|IDS7' +
    '|wfica32|wfcrun32|CDViewer' +
    '|tasy|Wheb|^Epic$|EpicSystems|Hyperspace|PowerChart|Cerner|MEDITECH|Soarian'

$script:SessaoPreservar = 'wfica32|wfcrun32|CDViewer|SelfService|Receiver|concentr|CtxWebHelper|AuthManSvr|redirector|HdxRtcEngine|CtxCFRUI|Citrix' +
    '|CentBrowser|msedge|chrome|firefox|iexplore|opera|vivaldi|brave' +
    '|ssonsvr|Microsoft\.AAD|AADBroker|TokenBroker' +
    '|WindowsTerminal|OpenConsole|FortiTray|FortiClient|FortiSSLVPN|Forti' +
    '|EXCEL|WINWORD|POWERPNT|OUTLOOK|MSACCESS|^olk$|onenote|StickyNot' +
    '|javaw|^java$|jp2launcher' +
    '|sqlservr|Tomcat|catalina|w3wp|inetinfo|MSMQ|postgres|mysqld|oracle|firebird' +
    '|^claude$|^claude-code$' +
    '|' + $script:AppsDeTrabalho +
    '|' + $script:SessaoRemoto

$script:SessaoNaoCompactar = '^claude$|^claude-code$|^python$|^pythonw$|concentr|sqlservr|Tomcat|w3wp' +
    '|' + $script:AppsDeTrabalho +
    '|' + $script:SessaoRemoto

$script:SnapProc = $null

function Get-SnapshotSvc {

    if ($script:SnapSvc -and $script:SnapSvc.Count -gt 0) { return $script:SnapSvc }

    $script:SnapSvc = @(Get-Service -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; DisplayName = $_.DisplayName; State = "$($_.Status)" }
    })

    if ($script:SnapSvc.Count -eq 0) {

        try {
            $script:SnapSvc = @(Get-CimInstance Win32_Service -Property Name,DisplayName,State -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ Name = $_.Name; DisplayName = $_.DisplayName; State = "$($_.State)" }
            })
        } catch { }
    }
    if ($script:SnapSvc.Count -eq 0) {

        Write-Log (T 'snapshotsvc.01') 'ALERTA'
    }
    return $script:SnapSvc
}

function Clear-SnapshotsOperacao {

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

        $r.Fresco = $true
    }
    $r.Protegidos = @($set.Keys)
    return $r
}

function Get-CaminhoAoVivo {

    param($Processo, $Mapa)
    $cam = try { [string]$Processo.Path } catch { '' }
    if (-not $cam -and $Mapa) { $cam = [string]$Mapa[[int]$Processo.Id] }
    return [string]$cam
}

function Get-SnapshotProc {

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

    if ($snap.Vivos.Count -gt 0) {
        $script:SnapProc   = $snap
        $script:SnapProcEm = Get-Date
    }
    return $snap
}

function Get-CaminhosProcesso {

    return (Get-SnapshotProc).Caminhos
}

function Get-SufixosDns {

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

function Test-Padrao {

    param([string]$Texto, [string]$Padrao)
    if (-not $Texto -or -not $Padrao) { return $false }
    try {
        return [regex]::IsMatch($Texto, $Padrao,
            [System.Text.RegularExpressions.RegexOptions]'IgnoreCase,CultureInvariant')
    } catch { return $false }
}

function Test-CaminhoEcossistema {

    param([string]$Caminho)
    if (-not $Caminho) { return $false }

    return ($Caminho.Split([char]92, [char]47) -contains $script:MarcaEcossistema)
}

$script:ArquivoEmUso = Join-Path $env:LOCALAPPDATA (Join-Path $script:MarcaEcossistema 'em_uso.json')
$script:EmUsoMaxMin  = 5

function Get-AppEmUso {

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

        $snap = Get-SnapshotProc
        $pai = $snap.Pais
        $vivos = $snap.Vivos
        $set = @{}

        foreach ($id in $r.Pids) { $set[$id] = $true }

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

    param($EmUso)
    if (-not $EmUso -or -not $EmUso.Existe) { return }
    if (-not $EmUso.Fresco) {
        Write-Log ((T 'logemuso.01') -f $EmUso.Idade.TotalMinutes) 'DADO'
        return
    }
    $quando = if ($EmUso.CarimboLegivel) { ('atualizada ha {0:N0} min' -f $EmUso.Idade.TotalMinutes) } else { 'com data ILEGIVEL, tratada como recente' }
    $oque   = if ($EmUso.Segmentando) { 'lote em andamento' } else { 'aplicativo aberto' }
    Write-Log ((T 'logemuso.02') -f $EmUso.App, $oque, $EmUso.Protegidos.Count, $quando) 'ALERTA'
    Write-Log ((T 'logemuso.03') -f $EmUso.Porta) 'DADO'
    if (-not $EmUso.CarimboLegivel) {
        Write-Log (T 'logemuso.04') 'ACAO'
        Write-Log ('   ' + $script:ArquivoEmUso) 'DADO'
    }
}

function Test-PodeEncerrarSessao {
    param([string]$Nome, [string]$Caminho = '', [int]$ProcId = 0, $EmUso = $null)

    if ($ProcId -gt 0 -and $EmUso -and $EmUso.Fresco -and ($EmUso.Protegidos -contains $ProcId)) { return $false }

    if (Test-CaminhoEcossistema $Caminho) { return $false }
    if (Test-Padrao $Nome $script:SessaoAudio) { return $true }
    if (Test-Padrao $Nome $script:SessaoEnfeites) { return $true }
    if (Test-Padrao $Nome $script:SessaoPreservar) { return $false }
    if ((Test-Padrao $Nome $script:Protegidos) -and -not (Test-Padrao $Nome 'msedge|chrome|firefox|CentBrowser')) { return $false }
    return $true
}

function Get-MomentoSessao {

    $r = [pscustomobject]@{
        IdadeMin    = $null
        Recem       = $false
        Ecossistema = 0
        Trabalho    = @()
        Veredicto   = 'INDEFINIDO'
        Texto       = ''
        Detalhe     = ''
    }

    try {
        $eu = (Get-Process -Id $PID).SessionId
        $ex = @(Get-Process explorer -ErrorAction SilentlyContinue |
                Where-Object { $_.SessionId -eq $eu } | Sort-Object StartTime)
        if ($ex.Count -gt 0) {
            $r.IdadeMin = [Math]::Round(((Get-Date) - $ex[0].StartTime).TotalMinutes)
            $r.Recem = ($r.IdadeMin -le $script:Lim.MomentoRecemMin)
        }
    } catch { }

    try {
        $nomes = @()
        foreach ($p in (Get-ProcessosSessao)) {
            if ($p.Classe -eq 'Ecossistema') { $r.Ecossistema++; continue }
            if ($p.TemJanela -and (Test-Padrao $p.Nome $script:AppsDeTrabalho)) { $nomes += $p.Nome }
        }
        $r.Trabalho = @($nomes | Sort-Object -Unique)
    } catch { }

    if ($null -eq $r.IdadeMin) {

        $r.Veredicto = 'INDEFINIDO'
        $r.Texto = 'Nao consegui medir ha quanto tempo esta sessao esta aberta.'
        $r.Detalhe = 'Rode o kit logo depois de logar. No meio do dia ele pode encerrar trabalho em curso.'
    } elseif ($r.Ecossistema -gt 0) {
        $r.Veredicto = 'EM_USO'
        $r.Texto = ('Ha {0} processo(s) de carga do ecossistema em andamento nesta sessao.' -f $r.Ecossistema)
        $r.Detalhe = 'Esses nao entram na lista e nao serao encerrados, mas a maquina esta trabalhando agora.'
    } elseif ($r.Trabalho.Count -gt 0) {
        $r.Veredicto = 'EM_USO'
        $r.Texto = ('O trabalho do dia parece ter comecado: {0} com janela aberta.' -f ($r.Trabalho -join ', '))
        $r.Detalhe = 'Esses programas estao protegidos e nao serao encerrados. Mas o kit e para rodar ANTES deles.'
    } elseif (-not $r.Recem) {
        $r.Veredicto = 'SESSAO_ANTIGA'
        $r.Texto = ('Esta sessao esta aberta ha {0} minuto(s).' -f $r.IdadeMin)
        $r.Detalhe = 'Nao vejo trabalho aberto, mas o kit e para rodar logo depois de logar.'
    } else {
        $r.Veredicto = 'RECEM'
        $r.Texto = ('Sessao aberta ha {0} minuto(s), sem trabalho aberto: e o momento certo.' -f $r.IdadeMin)
        $r.Detalhe = ''
    }
    return $r
}

function Write-MomentoSessao {

    param($Momento)
    if (-not $Momento) { return }
    Write-Log '' 'DADO'
    if ($Momento.Veredicto -eq 'RECEM') {
        Write-Log $Momento.Texto 'OK'
        return
    }
    $nivel = if ($Momento.Veredicto -eq 'EM_USO') { 'CRITICO' } else { 'ALERTA' }
    Write-Log $Momento.Texto $nivel
    if ($Momento.Detalhe) { Write-Log $Momento.Detalhe 'ALERTA' }
    Write-Log (T 'momentosessao.01') 'DADO'
}

function Get-ProcessosSessao {

    param([switch]$TodasAsSessoes)
    $chave = if ($TodasAsSessoes) { 'todas' } else { 'minha' }
    if ($script:SnapSessao -and $script:SnapSessao.ContainsKey($chave)) { return $script:SnapSessao[$chave] }

    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $caminhos = Get-CaminhosProcesso
    $emUso = Get-AppEmUso

    $protSet = @{}
    foreach ($id in $emUso.Protegidos) { $protSet[[int]$id] = $true }
    $saida = New-Object System.Collections.ArrayList
    try {
        foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
            if (-not $TodasAsSessoes -and $p.SessionId -ne $eu) { continue }
            if ($p.Id -eq $PID) { continue }
            $nome = $p.ProcessName

            $cam  = Get-CaminhoAoVivo -Processo $p -Mapa $caminhos
            if ($emUso.Fresco -and $protSet.ContainsKey([int]$p.Id)) {
                $classe = 'Ecossistema'
            } elseif (Test-CaminhoEcossistema $cam) {
                $classe = 'Ecossistema'
            } elseif ((Test-Padrao $nome $script:SessaoAudio) -or (Test-Padrao $nome $script:SessaoEnfeites)) {
                $classe = 'Dispensavel'
            } elseif (Test-Padrao $nome $script:SessaoPreservar) {
                $classe = 'Preservar'
            } elseif ((Test-Padrao $nome $script:Protegidos) -and -not (Test-Padrao $nome 'msedge|chrome|firefox|CentBrowser')) {
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

    if ($script:SnapGrupos) { return $script:SnapGrupos }

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

    Set-Status (T 'tarefasemexecucao.01')
    Pump
    $lista = @()
    try {
        foreach ($t in (Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.State -eq 'Running' })) {
            if ("$($t.TaskPath)" -like '\Microsoft\Windows\*') { continue }

            $exes = @()
            try { $exes = @($t.Actions | ForEach-Object { "$($_.Execute)" } | Where-Object { $_ }) } catch { }
            $doEcossistema = $false
            foreach ($x in $exes) {

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
    Write-Titulo (T 'diagsessaoativa.01')

    $todos = Get-ProcessosSessao
    if ($todos.Count -eq 0) { Write-Log (T 'diagsessaoativa.02') 'ALERTA'; return }

    $euId = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $maquina = @(Get-ProcessosSessao -TodasAsSessoes)
    $qtdSessoes = @($maquina | Select-Object -ExpandProperty Sessao -Unique).Count
    Write-Log ((T 'diagsessaoativa.03') -f $euId, $qtdSessoes) 'DADO'
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
        Write-Log ((T 'diagsessaoativa.04') -f $rot, $g.Count, (Format-Bytes $mem)) 'DADO'
        Add-ItemInv 'Sessao' $c ("$($g.Count) processos") ((Format-Bytes $mem) + " | sessao $euId")
    }

    $ecoFora = @($maquina | Where-Object { $_.Classe -eq 'Ecossistema' -and -not $_.Minha })
    if ($ecoFora.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ((T 'diagsessaoativa.05') -f $ecoFora.Count) 'ALERTA'
        foreach ($e in ($ecoFora | Sort-Object Memoria -Descending | Select-Object -First 6)) {
            Write-Log ((T 'diagsessaoativa.06') -f $e.Sessao, $e.Nome, $e.Id, (Format-Bytes $e.Memoria)) 'DADO'
        }
        Add-ItemInv 'Sessao' 'Carga do ecossistema fora desta sessao' ("$($ecoFora.Count) processos") (Format-Bytes (($ecoFora | Measure-Object Memoria -Sum).Sum))
    }

    Write-LogEmUso (Get-AppEmUso)
    $eco = @($todos | Where-Object { $_.Classe -eq 'Ecossistema' })
    $ecoMaquina = @($maquina | Where-Object { $_.Classe -eq 'Ecossistema' })
    if ($eco.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log (T 'diagsessaoativa.07') 'DADO'
        foreach ($e in ($eco | Sort-Object Memoria -Descending | Select-Object -First 6)) {
            Write-Log ((T 'diagsessaoativa.08') -f $e.Nome, $e.Id, (Format-Bytes $e.Memoria), $e.Caminho) 'DADO'
        }
    }
    $pesado = @($ecoMaquina | Where-Object { $_.Memoria -gt 1GB } | Sort-Object Memoria -Descending)
    if ($pesado.Count -gt 0) {
        $onde = if ($pesado[0].Minha) { 'nesta sessao' } else { ('na sessao ' + $pesado[0].Sessao) }
        Add-Achado 'ALERTA' ((T 'diagsessaoativa.09') -f $onde, $pesado[0].Nome, (Format-Bytes $pesado[0].Memoria)) (T 'diagsessaoativa.10') (T 'modulo3limpeza.09') 'Alto'
    }

    $grupos = Get-GruposSessao
    $memDisp = ($grupos | Measure-Object Memoria -Sum).Sum
    if ($grupos.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log (T 'diagsessaoativa.11') 'DADO'
        foreach ($g in ($grupos | Select-Object -First 12)) {
            Write-Log ((T 'diagsessaoativa.12') -f $(if ($g.TemJanela) { 'j' } else { ' ' }), $g.Nome, (Format-Bytes $g.Memoria), $g.Qtd) 'DADO'
        }
    }

    try {
        $bt = @(Get-SnapshotSvc | Where-Object { $_.Name -eq 'bthserv' -and $_.State -eq 'Running' })
        $cabo = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -notmatch 'Wi-?Fi|Wireless|802\.11' })
        if ($bt.Count -gt 0 -and $cabo.Count -gt 0) {
            Write-Log '' 'DADO'
            Add-Achado 'ALERTA' (T 'diagsessaoativa.13') (T 'diagsessaoativa.14') (T 'modulo3limpeza.09') 'Medio'
        }
    } catch { }

    $tarefas = Get-TarefasEmExecucao
    if ($tarefas.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ((T 'diagsessaoativa.15') -f $tarefas.Count) 'DADO'
        foreach ($t in $tarefas) { Write-Log ((T 'diagsessaoativa.16') -f $t.Caminho, $t.Nome) 'DADO' }
    }

    if ($memDisp -gt 400MB -or $grupos.Count -ge 6) {
        Add-Achado 'ALERTA' ((T 'diagsessaoativa.17') -f (Format-Bytes $memDisp), $grupos.Count) (T 'diagsessaoativa.18') (T 'modulo3limpeza.09') 'Alto'
    } else {
        Add-Achado 'OK' ((T 'diagsessaoativa.19') -f (Format-Bytes $memDisp)) '' (T 'modulo3limpeza.09')
    }
}

function Get-PlanoSessao {
    $plano = New-Object System.Collections.ArrayList

    foreach ($g in (Get-GruposSessao)) {

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

    foreach ($t in (Get-TarefasEmExecucao)) {

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

    [void]$plano.Add([pscustomobject]@{
        Grupo = 'PRIORIDADE'; Rotulo = (T 'item.42'); Valor = ''; Bytes = 0
        Marcar = $true; Dados = 'prioridade'
        Nota = (T 'nota.15')
    })

    [void]$plano.Add([pscustomobject]@{
        Grupo = 'MEMORIA'; Rotulo = (T 'item.43'); Valor = ''; Bytes = 0
        Marcar = $true; Dados = 'trim'
        Nota = (T 'nota.16')
    })

    return $plano
}

function Invoke-Modulo7Sessao {
    Clear-SnapshotsOperacao
    Write-Titulo (T 'modulo7sessao.01')
    Write-Log (T 'modulo7sessao.02') 'DADO'
    Write-Log (T 'modulo7sessao.03') 'DADO'
    Write-Log (T 'modulo7sessao.04') 'DADO'

    Write-MomentoSessao (Get-MomentoSessao)

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
            Write-Log ((T 'modulo7sessao.05') -f ($papeis -join ', ')) 'CRITICO'
            Write-Log (T 'modulo7sessao.06') 'ALERTA'
            Write-Log (T 'modulo7sessao.07') 'ALERTA'
        }
    } catch { }

    Set-Status (T 'modulo7sessao.08')
    $plano = Get-PlanoSessao
    Set-Status (T 'modulo7sessao.09')
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
    Write-Log (T 'modulo3analise.03') 'TITULO'
    Write-Log ((T 'modulo7sessao.10') -f $nProc) 'ACAO'
    Write-Log ((T 'modulo3limpeza.43') -f (Format-Bytes $mem)) 'ACAO'
    Write-Log ((T 'modulo7sessao.11') -f $nTar) 'ACAO'
    Write-Log (T 'modulo7sessao.12') 'ACAO'
    Write-Log (T 'modulo7sessao.13') 'ACAO'
    if ($naoRec.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ((T 'modulo7sessao.14') -f $naoRec.Count) 'DADO'
        foreach ($n in $naoRec) { Write-Log ('   ' + $n.Rotulo) 'DADO' }
        Write-Log (T 'modulo7sessao.15') 'DADO'
    }
    Write-Log '' 'DADO'
    Write-Log (T 'modulo3analise.11') 'ACAO'
    $script:PainelGrandes.Visible = $false
    $script:PainelLimpeza.Visible = $true
}

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

    if ($Estado -eq 'OTIMIZADA' -and $script:StatusSessao -and $script:StatusSessao.Encerrados.Count -gt 0) {
        Add-LinhaStatus ('ENCERRADOS NESTA SESSAO ({0})' -f $script:StatusSessao.Encerrados.Count) $script:Cor.Titulo
        foreach ($n in ($script:StatusSessao.Encerrados | Select-Object -Unique)) {
            Add-LinhaStatus ('  ' + $n) $script:Cor.Fraco
        }
        Add-LinhaStatus ''
    }

    if ($Estado -eq 'OTIMIZADA' -and $script:UltimaPrioridade) {
        Add-LinhaStatus 'PRIORIDADE DE CPU' $script:Cor.Titulo
        Add-LinhaStatus ('  {0} clinico(s) acima do normal, {1} abaixo' -f $script:UltimaPrioridade.Acima, $script:UltimaPrioridade.Abaixo)
        Add-LinhaStatus ''
    }

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

    $script:ListaLimpeza.Items.Clear()
    $script:PlanoAtual = @()
    $script:PainelLimpeza.Visible = $false
    $script:PainelGrandes.Visible = $false
    $script:PainelSessao.Visible = $true
    Pump
}

function Invoke-RestaurarSessao {
    Write-Titulo (T 'restaurarsessao.01')
    Write-Log (T 'restaurarsessao.02') 'DADO'

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
    Write-Log ((T 'restaurarsessao.03') -f $n) 'OK'

    $e = Read-Estado
    $fechados = @()
    if ($e.ContainsKey('sessao_encerrados')) { $fechados = @($e['sessao_encerrados']) }

    if ($fechados.Count -eq 0) {
        Write-Log (T 'restaurarsessao.04') 'DADO'
    } else {
        Write-Log ((T 'restaurarsessao.05') -f $fechados.Count) 'DADO'
    }

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
        [pscustomobject]@{ Padrao = '^GoogleDriveFS$'; Rotulo = (T 'item.44'); Caminhos = $driveFs; Args = '' }
        [pscustomobject]@{ Padrao = 'OneDrive'; Rotulo = (T 'item.16'); Caminhos = @(
            (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
            (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')); Args = '/background' }
        [pscustomobject]@{ Padrao = 'ms-teams|msteams|^Teams$'; Rotulo = (T 'item.02'); Caminhos = @(
            (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\MSTeams_8wekyb3d8bbwe\ms-teams.exe')); Args = '' }
        [pscustomobject]@{ Padrao = 'soffice'; Rotulo = (T 'item.45'); Caminhos = @(
            (Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.exe')); Args = '' }
    )

    $reabertos = 0
    foreach ($m in $mapa) {
        $foiFechado = @($fechados | Where-Object { "$_" -match $m.Padrao }).Count -gt 0
        if (-not $foiFechado) { continue }
        if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match $m.Padrao }).Count -gt 0) {
            Write-Log ((T 'restaurarsessao.06') -f $m.Rotulo) 'DADO'
            continue
        }
        $abriu = $false
        foreach ($c in $m.Caminhos) {
            if (-not $c -or -not (Test-Path -LiteralPath $c)) { continue }
            try {
                if ($m.Args) { Start-Process -FilePath $c -ArgumentList $m.Args | Out-Null }
                else { Start-Process -FilePath $c | Out-Null }
                Write-Log ((T 'restaurarsessao.07') -f $m.Rotulo) 'OK'
                $abriu = $true; $reabertos++
                break
            } catch { }
        }
        if (-not $abriu) { Write-Log ((T 'restaurarsessao.08') -f $m.Rotulo) 'ALERTA' }
    }

    $sozinhos = @($fechados | Where-Object { "$_" -match 'SearchApp|ShellExperience|StartMenu|TextInput|Notification|Update|jusched|jucheck|Broker|CompPkgSrv|Widget|Copilot|Audio|Waves|Nahimic|Realtek|RAVBg|Rtk|Dolby|DTS|Logi|iCUE' })
    if ($sozinhos.Count -gt 0) {
        Write-Log '' 'DADO'
        Write-Log ((T 'restaurarsessao.09') -f $sozinhos.Count) 'DADO'
        foreach ($x in ($sozinhos | Select-Object -First 12 -Unique)) { Write-Log ('   ' + $x) 'DADO' }
        Write-Log (T 'restaurarsessao.10') 'DADO'
    }

    try {
        $e2 = Read-Estado
        if ($e2.ContainsKey('sessao_encerrados')) {
            $e2.Remove('sessao_encerrados')
            Write-Estado $e2
        }
    } catch { }

    Write-Log '' 'DADO'
    Write-Log ((T 'restaurarsessao.11') -f $reabertos) 'OK'
    Write-Log (T 'restaurarsessao.12') 'DADO'
}

function Invoke-AcaoPrioridade {

    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $emUso = Get-EmUsoOperacao -Renovar
    $caminhos = Get-CaminhosProcesso
    $subiu = 0
    $baixou = 0
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.SessionId -ne $eu -or $p.Id -eq $PID) { continue }
        try {
            if ((Test-Padrao $p.ProcessName $script:SessaoPreservar) -and -not (Test-Padrao $p.ProcessName $script:SessaoAudio) -and -not (Test-Padrao $p.ProcessName $script:SessaoEnfeites)) {
                $p.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal
                $subiu++
            } elseif (Test-PodeEncerrarSessao -Nome $p.ProcessName -Caminho (Get-CaminhoAoVivo -Processo $p -Mapa $caminhos) -ProcId ([int]$p.Id) -EmUso $emUso) {
                $p.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::BelowNormal
                $baixou++
            }
        } catch { }
    }
    $script:UltimaPrioridade = [pscustomobject]@{ Acima = $subiu; Abaixo = $baixou }
    Write-Log ((T 'acaoprioridade.01') -f $subiu, $baixou) 'OK'
    Write-Log (T 'acaoprioridade.02') 'DADO'
}

function Invoke-AcaoMemoria {

    $eu = try { (Get-Process -Id $PID).SessionId } catch { 1 }
    $emUso = Get-EmUsoOperacao -Renovar
    $caminhos = Get-CaminhosProcesso
    $antes = 0.0
    $depois = 0.0
    $n = 0
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.SessionId -ne $eu -or $p.Id -eq $PID) { continue }

        if (Test-Padrao $p.ProcessName $script:SessaoNaoCompactar) { continue }

        if (Test-CaminhoEcossistema (Get-CaminhoAoVivo -Processo $p -Mapa $caminhos)) { continue }
        if ($emUso.Fresco -and ($emUso.Protegidos -contains [int]$p.Id)) { continue }
        try {
            $ws = $p.WorkingSet64

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
    if ($n -eq 0) { Write-Log (T 'acaomemoria.01') 'ALERTA'; return }
    Write-Log ((T 'acaomemoria.02') -f $n, (Format-Bytes $ganho)) 'OK'
    Write-Log (T 'acaomemoria.03') 'DADO'
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
$script:Form.Text          = (T 'topo.01')
$script:Form.Size          = New-Object System.Drawing.Size(1240, 820)
$script:Form.MinimumSize   = New-Object System.Drawing.Size(980, 620)
$script:Form.StartPosition = 'CenterScreen'
$script:Form.BackColor     = $script:Cor.Fundo
$script:Form.ForeColor     = $script:Cor.Texto
$script:Form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

$cab = New-Object System.Windows.Forms.Panel
$cab.Dock = 'Top'; $cab.Height = 58; $cab.BackColor = $script:Cor.Painel
$lblTit = New-Object System.Windows.Forms.Label
$lblTit.Text = (T 'topo.02')
$lblTit.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
$lblTit.ForeColor = $script:Cor.Texto
$lblTit.SetBounds(18, 9, 620, 24)
$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = ((T 'topo.03') -f $script:Versao, $env:COMPUTERNAME, $env:USERNAME)
$lblSub.ForeColor = $script:Cor.Fraco
$lblSub.SetBounds(20, 33, 700, 18)
$cab.Controls.AddRange(@($lblTit, $lblSub))

$rodape = New-Object System.Windows.Forms.Panel
$rodape.Dock = 'Bottom'; $rodape.Height = 46; $rodape.BackColor = $script:Cor.Painel

$script:LblStatus = New-Object System.Windows.Forms.Label
$script:LblStatus.Text = (T 'topo.04')
$script:LblStatus.ForeColor = $script:Cor.Fraco
$script:LblStatus.SetBounds(18, 15, 430, 18)

$script:Barra = New-Object System.Windows.Forms.ProgressBar
$script:Barra.SetBounds(452, 14, 150, 16)
$script:Barra.Style = 'Blocks'

$btnCopiar = New-Object System.Windows.Forms.Button
$btnCopiar.Text = (T 'topo.05'); $btnCopiar.SetBounds(620, 9, 100, 28); $btnCopiar.FlatStyle = 'Flat'
$btnCopiar.BackColor = $script:Cor.Botao; $btnCopiar.ForeColor = $script:Cor.Texto
$btnCopiar.FlatAppearance.BorderColor = $script:Cor.Borda

$btnSalvar = New-Object System.Windows.Forms.Button
$btnSalvar.Text = (T 'topo.06'); $btnSalvar.SetBounds(728, 9, 100, 28); $btnSalvar.FlatStyle = 'Flat'
$btnSalvar.BackColor = $script:Cor.Botao; $btnSalvar.ForeColor = $script:Cor.Texto
$btnSalvar.FlatAppearance.BorderColor = $script:Cor.Borda

$btnLimparLog = New-Object System.Windows.Forms.Button
$btnLimparLog.Text = (T 'topo.07'); $btnLimparLog.SetBounds(836, 9, 100, 28); $btnLimparLog.FlatStyle = 'Flat'
$btnLimparLog.BackColor = $script:Cor.Botao; $btnLimparLog.ForeColor = $script:Cor.Texto
$btnLimparLog.FlatAppearance.BorderColor = $script:Cor.Borda

$script:BtnCancelar = New-Object System.Windows.Forms.Button
$script:BtnCancelar.Text = (T 'topo.08'); $script:BtnCancelar.SetBounds(944, 9, 90, 28); $script:BtnCancelar.FlatStyle = 'Flat'
$script:BtnCancelar.BackColor = $script:Cor.Botao; $script:BtnCancelar.ForeColor = $script:Cor.Alerta
$script:BtnCancelar.FlatAppearance.BorderColor = $script:Cor.Borda
$script:BtnCancelar.Enabled = $false

$rodape.Controls.AddRange(@($script:LblStatus, $script:Barra, $btnCopiar, $btnSalvar, $btnLimparLog, $script:BtnCancelar))

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

$script:BtnMod2 = New-Botao -Texto (T 'topo.09')     -Y 16
$script:BtnMod3 = New-Botao -Texto (T 'modulo3limpeza.02')                      -Y 66
$script:BtnGrandes = New-Botao -Texto (T 'topo.10') -Y 116
$script:BtnSessao  = New-Botao -Texto (T 'topo.11') -Y 166
$script:BtnPersist = New-Botao -Texto (T 'topo.12') -Y 216

$sec3 = New-Secao -Texto (T 'topo.13') -Y 280
$script:BtnRestaurar = New-Botao -Texto (T 'topo.14') -Y 302 -Altura 32 -Secundario -Cor $script:Cor.Titulo
$script:BtnDesfazer  = New-Botao -Texto (T 'otimizacoes.01')    -Y 340 -Altura 32 -Secundario -Cor $script:Cor.Acao
$script:BtnReiniciar = New-Botao -Texto (T 'restartcomputador.01')  -Y 378 -Altura 32 -Secundario -Cor $script:Cor.Alerta

$lblRodape = New-Object System.Windows.Forms.Label
$lblRodape.Text = (T 'topo.15')
$lblRodape.SetBounds(18, 462, 244, 68); $lblRodape.ForeColor = $script:Cor.Fraco
$lblRodape.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$script:BtnIdioma = @{}
$xIdioma = 14
foreach ($idi in $script:Idiomas) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $idi.Cod.ToUpper()
    $b.SetBounds($xIdioma, 424, 78, 26)
    $b.FlatStyle = 'Flat'
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize = 1
    $b.Tag = $idi.Cod
    $script:BtnIdioma[$idi.Cod] = $b
    $xIdioma += 82
}

function Update-BotoesIdioma {
    foreach ($cod in $script:BtnIdioma.Keys) {
        $b = $script:BtnIdioma[$cod]
        if ($cod -eq $script:Idioma) {
            $b.BackColor = $script:Cor.Botao
            $b.ForeColor = $script:Cor.Titulo
            $b.FlatAppearance.BorderColor = $script:Cor.Titulo
        } else {
            $b.BackColor = $script:Cor.Painel
            $b.ForeColor = $script:Cor.Fraco
            $b.FlatAppearance.BorderColor = $script:Cor.Borda
        }
    }
}

function Set-IdiomaInterface {
    param([string]$Cod)
    if (-not $Cod -or $Cod -eq $script:Idioma) { return }
    $script:Idioma = $Cod
    Save-Idioma $Cod
    try {
        $lblTit.Text    = (T 'topo.02')
        $lblSub.Text    = ((T 'topo.03') -f $script:Versao, $env:COMPUTERNAME, $env:USERNAME)
        $lblRodape.Text = (T 'topo.15')
        $script:BtnCancelar.Text  = (T 'topo.08')
        $script:BtnMod2.Text      = (T 'topo.09')
        $script:BtnMod3.Text      = (T 'modulo3limpeza.02')
        $script:BtnGrandes.Text   = (T 'topo.10')
        $script:BtnSessao.Text    = (T 'topo.11')
        $script:BtnPersist.Text   = (T 'topo.12')
        $sec3.Text                = (T 'topo.13')
        $script:BtnRestaurar.Text = (T 'topo.14')
        $script:BtnDesfazer.Text  = (T 'otimizacoes.01')
        $script:BtnReiniciar.Text = (T 'restartcomputador.01')
        $script:BtnLimparSel.Text = (T 'topo.21')
        $script:BtnReverterSes.Text = (T 'topo.26')
    } catch { }
    Update-BotoesIdioma
    Write-Log ('-- ' + (T 'idioma.01') + ': ' + $Cod.ToUpper()) 'DADO'
}

foreach ($cod in @($script:BtnIdioma.Keys)) {
    $script:BtnIdioma[$cod].Add_Click({ Set-IdiomaInterface ([string]$this.Tag) }.GetNewClosure())
}

$lateral.Controls.AddRange(@($script:BtnIdioma.Values))
Update-BotoesIdioma
$lateral.Controls.AddRange(@($script:BtnMod2, $script:BtnMod3,
                             $script:BtnGrandes, $script:BtnSessao, $script:BtnPersist,
                             $sec3, $script:BtnRestaurar, $script:BtnDesfazer, $script:BtnReiniciar, $lblRodape))

function New-BotaoPainel {
    param([string]$Texto, [int]$X, [int]$Y, [int]$L, [int]$A, $Cor = $null, [switch]$Fraco)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Texto; $b.SetBounds($X, $Y, $L, $A); $b.FlatStyle = 'Flat'
    $b.BackColor = $(if ($Fraco) { $script:Cor.Painel } else { $script:Cor.Botao })
    $b.ForeColor = $(if ($Cor) { $Cor } else { $script:Cor.Texto })
    $b.FlatAppearance.BorderColor = $script:Cor.Borda
    return $b
}

$script:PainelLimpeza = New-Object System.Windows.Forms.Panel
$script:PainelLimpeza.Dock = 'Right'; $script:PainelLimpeza.Width = 520
$script:PainelLimpeza.BackColor = $script:Cor.Painel
$script:PainelLimpeza.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)
$script:PainelLimpeza.Visible = $false

$topoLimp = New-Object System.Windows.Forms.Panel
$topoLimp.Dock = 'Top'; $topoLimp.Height = 146; $topoLimp.BackColor = $script:Cor.Painel

$lblLimp = New-Object System.Windows.Forms.Label
$lblLimp.Text = (T 'topo.16')
$lblLimp.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblLimp.SetBounds(0, 12, 300, 20); $lblLimp.ForeColor = $script:Cor.Texto

$lblLimp2 = New-Object System.Windows.Forms.Label
$lblLimp2.Text = (T 'topo.17')
$lblLimp2.SetBounds(0, 34, 480, 32); $lblLimp2.ForeColor = $script:Cor.Fraco
$lblLimp2.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$lblDown = New-Object System.Windows.Forms.Label
$lblDown.Text = (T 'topo.18')
$lblDown.SetBounds(0, 76, 180, 20); $lblDown.ForeColor = $script:Cor.Texto

$script:ComboPeriodo = New-Object System.Windows.Forms.ComboBox
$script:ComboPeriodo.SetBounds(184, 72, 290, 24)
$script:ComboPeriodo.DropDownStyle = 'DropDownList'
$script:ComboPeriodo.BackColor = $script:Cor.Fundo
$script:ComboPeriodo.ForeColor = $script:Cor.Texto
$script:ComboPeriodo.FlatStyle = 'Flat'
foreach ($op in @('3 dias', '7 dias', '15 dias', '30 dias', 'tudo, sem limite de idade')) { [void]$script:ComboPeriodo.Items.Add($op) }
$script:ComboPeriodo.SelectedIndex = 1

$btnMarcar    = New-BotaoPainel -Texto (T 'topo.19')    -X 0   -Y 106 -L 150 -A 28
$btnDesmarcar = New-BotaoPainel -Texto (T 'topo.20') -X 158 -Y 106 -L 150 -A 28
$topoLimp.Controls.AddRange(@($lblLimp, $lblLimp2, $lblDown, $script:ComboPeriodo, $btnMarcar, $btnDesmarcar))

$rodapeLimp = New-Object System.Windows.Forms.Panel
$rodapeLimp.Dock = 'Bottom'; $rodapeLimp.Height = 66; $rodapeLimp.BackColor = $script:Cor.Painel
$script:BtnLimparSel = New-BotaoPainel -Texto (T 'topo.21') -X 0 -Y 10 -L 300 -A 44 -Cor $script:Cor.Ok
$script:BtnLimparSel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnFecharPainel = New-BotaoPainel -Texto (T 'topo.22') -X 312 -Y 10 -L 162 -A 44 -Fraco
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

$script:PainelGrandes = New-Object System.Windows.Forms.Panel
$script:PainelGrandes.Dock = 'Right'; $script:PainelGrandes.Width = 520
$script:PainelGrandes.BackColor = $script:Cor.Painel
$script:PainelGrandes.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)
$script:PainelGrandes.Visible = $false

$topoGr = New-Object System.Windows.Forms.Panel
$topoGr.Dock = 'Top'; $topoGr.Height = 78; $topoGr.BackColor = $script:Cor.Painel
$lblGr = New-Object System.Windows.Forms.Label
$lblGr.Text = (T 'topo.10')
$lblGr.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblGr.SetBounds(0, 12, 320, 20); $lblGr.ForeColor = $script:Cor.Texto
$lblGr2 = New-Object System.Windows.Forms.Label
$lblGr2.Text = (T 'topo.23')
$lblGr2.SetBounds(0, 34, 480, 34); $lblGr2.ForeColor = $script:Cor.Fraco
$lblGr2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$topoGr.Controls.AddRange(@($lblGr, $lblGr2))

$rodapeGr = New-Object System.Windows.Forms.Panel
$rodapeGr.Dock = 'Bottom'; $rodapeGr.Height = 66; $rodapeGr.BackColor = $script:Cor.Painel
$btnAbrir = New-BotaoPainel -Texto (T 'openitemgrande.02') -X 0 -Y 10 -L 300 -A 44 -Cor $script:Cor.Titulo
$btnAbrir.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnFecharGr = New-BotaoPainel -Texto (T 'topo.22') -X 312 -Y 10 -L 162 -A 44 -Fraco
$rodapeGr.Controls.AddRange(@($btnAbrir, $btnFecharGr))

$script:ListaGrandes = New-Object System.Windows.Forms.ListBox
$script:ListaGrandes.Dock = 'Fill'
$script:ListaGrandes.BackColor = $script:Cor.Fundo
$script:ListaGrandes.ForeColor = $script:Cor.Texto
$script:ListaGrandes.BorderStyle = 'FixedSingle'
$script:ListaGrandes.HorizontalScrollbar = $true
$script:ListaGrandes.Font = New-Object System.Drawing.Font('Consolas', 9)

$script:PainelGrandes.Controls.AddRange(@($script:ListaGrandes, $rodapeGr, $topoGr))

$script:PainelSessao = New-Object System.Windows.Forms.Panel
$script:PainelSessao.Dock = 'Right'; $script:PainelSessao.Width = 520
$script:PainelSessao.BackColor = $script:Cor.Painel
$script:PainelSessao.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 0)
$script:PainelSessao.Visible = $false

$topoSes = New-Object System.Windows.Forms.Panel
$topoSes.Dock = 'Top'; $topoSes.Height = 72; $topoSes.BackColor = $script:Cor.Painel
$lblSes = New-Object System.Windows.Forms.Label
$lblSes.Text = (T 'topo.24')
$lblSes.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblSes.SetBounds(0, 12, 320, 20); $lblSes.ForeColor = $script:Cor.Texto
$lblSes2 = New-Object System.Windows.Forms.Label
$lblSes2.Text = (T 'topo.25')
$lblSes2.SetBounds(0, 34, 480, 32); $lblSes2.ForeColor = $script:Cor.Fraco
$lblSes2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$topoSes.Controls.AddRange(@($lblSes, $lblSes2))

$rodapeSes = New-Object System.Windows.Forms.Panel
$rodapeSes.Dock = 'Bottom'; $rodapeSes.Height = 66; $rodapeSes.BackColor = $script:Cor.Painel
$script:BtnReverterSes = New-BotaoPainel -Texto (T 'topo.26') -X 0 -Y 10 -L 240 -A 44 -Cor $script:Cor.Alerta
$script:BtnReverterSes.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$btnAtualizarSes = New-BotaoPainel -Texto (T 'topo.27') -X 252 -Y 10 -L 110 -A 44
$btnFecharSes = New-BotaoPainel -Texto (T 'topo.22') -X 374 -Y 10 -L 118 -A 44 -Fraco
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

function Invoke-ComProtecao {
    param([scriptblock]$Bloco, [string]$Nome)
    if ($script:Ocupado) { return }
    $script:Cancelar = $false
    Set-Ocupado $true
    try { & $Bloco }
    catch { Write-Log ((T 'comprotecao.01') -f $Nome, $_.Exception.Message) 'CRITICO' }
    finally {
        Set-Ocupado $false
        Set-Status (T 'topo.04')
        $script:Cancelar = $false
    }
}

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
        Write-Log ((T 'topo.28') -f $script:ComboPeriodo.SelectedItem) 'DADO'
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
    Set-Status (T 'topo.29')
})

$btnCopiar.Add_Click({
    try {
        [System.Windows.Forms.Clipboard]::SetText($script:Log.Text)
        Set-Status (T 'topo.30')
    } catch { Set-Status (T 'topo.31') }
})

$btnSalvar.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Arquivo de texto (*.txt)|*.txt'
    $dlg.FileName = ('Log_{0}_{1:yyyyMMdd_HHmm}.txt' -f $env:COMPUTERNAME, (Get-Date))
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:Log.Text | Out-File -FilePath $dlg.FileName -Encoding UTF8
            Set-Status ((T 'topo.32') + $dlg.FileName)
        } catch { Set-Status (T 'topo.33') }
    }
})

$btnLimparLog.Add_Click({ $script:Log.Clear(); Set-Status (T 'topo.34') })

$script:Form.Add_Shown({
    Write-Log (T 'topo.35') 'OK'
    Write-Log ((T 'topo.36') -f $env:COMPUTERNAME, $env:USERNAME, (Get-Date)) 'DADO'
    Write-Log ''
        Write-Log (T 'topo.37') 'ACAO'
    Write-Log (T 'topo.38') 'ACAO'
    Write-Log (T 'topo.39') 'ACAO'
    Write-Log (T 'topo.40') 'ACAO'
    Write-Log (T 'topo.41') 'DADO'
    Write-Log (T 'topo.42') 'DADO'
    Write-Log (T 'topo.43') 'DADO'
})

[void]$script:Form.ShowDialog()
