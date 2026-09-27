#requires -Version 5.1
<#
    Verificar-Vazamento.ps1 - nada de infraestrutura de terceiro neste repositorio.

    POR QUE POR CATEGORIA, E NAO POR LISTA DE NOMES

    Este projeto nasceu de um fork cujo arquivo de origem tinha 47 ocorrencias de
    identificador de infraestrutura de um hospital real: dominio, fileserver,
    servidor de scanner, dois IPs internos, cinco hostnames e o caminho de um
    compartilhamento de departamento.

    A primeira versao desta conferencia era um `git grep` com esses nomes escritos
    dentro dele. Isso se anula: para conferir que os nomes nao estao no
    repositorio, os nomes passavam a estar no repositorio - dentro do proprio
    comando de conferencia, em documento versionado.

    Entao a conferencia e por CATEGORIA. Procura a FORMA de um identificador de
    rede, nao o nome de ninguem:

        - endereco IPv4 privado (10.x, 172.16-31.x, 192.168.x)
        - caminho UNC apontando para host nomeado
        - nome de dominio com TLD, fora de uma lista curta de dominios publicos
        - endereco de e-mail

    Isso e melhor que a lista por dois motivos. Nao nomeia ninguem, e pega
    vazamento que ninguem previu - inclusive de outra clinica, de outro pais, de
    um dado que ainda nao existe.

    O QUE ELE NAO PEGA - LIMITES DECLARADOS

    Um varredor que promete o que nao cumpre e pior que varredor nenhum, porque
    produz a sensacao de ter conferido. Estes sao os furos conhecidos:

      - NOME CORPORATIVO SEM TLD. Um dominio escrito sem o sufixo - so o nome da
        instituicao - nao e pego por nenhuma categoria, porque tem a forma de
        qualquer palavra. Nao ha como pegar por forma.
        O que fazer: nao escreva o nome. A documentacao deste projeto descreve o
        que saiu SEM nomear ninguem, exatamente por isso.

      - UNC PARA HOST DE NOME CURTO E SEM DIGITO. '\\ARQUIVOS\publico' escapa,
        porque exigir menos que isso acendia em '\\Citrix\\' e '\\Patients\\' -
        regex com barra escapada, nao caminho de rede.
        O que fazer: revisar toda string que comece com duas barras.

      - DADO DENTRO DE ARQUIVO BINARIO. So varre extensao de texto.

      - NOME DE PESSOA, de paciente ou de funcionario. Nao tem forma nenhuma.
        Nenhum varredor pega isso, e e por isso que a regra do ecossistema e
        separacao FISICA: o que nao pode ir para a nuvem nao existe na arvore.

    FALSO POSITIVO E ESPERADO, E TEM TRATAMENTO

    Codigo legitimo cita dominio publico (microsoft.com), IP de documentacao e
    endereco de loopback. O arquivo tools\vazamento-permitido.txt lista o que foi
    olhado e liberado, um padrao por linha, com o motivo em comentario. Liberar
    exige escrever o motivo - e o que impede a lista de virar depositario de
    tudo o que incomoda.

    COMO RODAR

        powershell -NoProfile -ExecutionPolicy Bypass -File tools\Verificar-Vazamento.ps1

    Sai com 0 se limpo, ou com o numero de achados. O empacotador chama isto antes
    de gerar o .zip: e o que faz a promessa valer no tempo, em vez de valer so no
    dia em que o fork foi criado.
#>

[CmdletBinding()]
param(
    # Pasta a varrer. Vazio = a raiz do projeto, resolvida no corpo.
    #
    # NAO resolve aqui: $PSScriptRoot vem VAZIO em valor padrao de param quando o
    # script e chamado com -File, e Split-Path de string vazia lanca. Resolver no
    # corpo funciona nos dois modos de chamada.
    [string]$Raiz = '',
    # Mostra tambem o que foi liberado pela lista de excecoes.
    [switch]$Detalhado
)

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$meuDir = $PSScriptRoot
if (-not $meuDir) { $meuDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Raiz) { $Raiz = Split-Path -Parent $meuDir }
if (-not $Raiz) { Write-Host 'Nao consegui descobrir a raiz do projeto.'; exit 1 }

# ---------------------------------------------------------------------
# As categorias. Forma de identificador, nunca nome de ninguem.
# ---------------------------------------------------------------------
$CATEGORIAS = @(
    [pscustomobject]@{
        Nome    = 'IP privado'
        Padrao  = '\b(?:10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3})\b'
        Porque  = 'endereco de rede interna de alguem'
    }
    [pscustomobject]@{
        Nome    = 'UNC para host nomeado'
        # So dispara quando o host PARECE host. Tres formas, em alternancia:
        #   1. com ponto        \\fs.example.com\  ou  \\10.0.0.1\      EXEMPLO-FICTICIO
        #   2. com sublinhado   \\file_server\                          EXEMPLO-FICTICIO
        #   3. comprido, com digito   \\hostAB0000\                     EXEMPLO-FICTICIO
        #
        # Exigir isso e deliberado. A versao anterior era '\\\\[A-Za-z][\w.-]{2,}\\'
        # e acendia em '\\Citrix\\', '\\Patients\\', '\\Windows\\' - que em codigo
        # PowerShell sao regex com barra ESCAPADA, nao caminho de rede. Dezessete
        # falsos positivos num arquivo, o que treina qualquer pessoa a ignorar o
        # varredor. Ver o LIMITE no cabecalho: '\\NOMECURTO\share' escapa.
        Padrao  = '\\\\(?![.?])(?!\$)(?:[A-Za-z0-9][A-Za-z0-9-]*\.[A-Za-z0-9][A-Za-z0-9.-]*|[A-Za-z0-9]+_[A-Za-z0-9_-]+|[A-Za-z][A-Za-z0-9-]*[0-9][A-Za-z0-9-]*)\\'
        Porque  = 'caminho para servidor de arquivos de alguem'
    }
    [pscustomobject]@{
        Nome    = 'dominio com TLD'
        Padrao  = '\b[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9][A-Za-z0-9-]*)+\.(?:com|net|org|gov|edu|br|mil|int|info|io|ai|co)\b'
        Porque  = 'nome de dominio, possivelmente corporativo'
    }
    [pscustomobject]@{
        Nome    = 'endereco de e-mail'
        Padrao  = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'
        Porque  = 'endereco de pessoa'
    }
)

# ---------------------------------------------------------------------
$permitidos = @()
$arqPerm = Join-Path $meuDir 'vazamento-permitido.txt'
if (Test-Path -LiteralPath $arqPerm) {
    foreach ($l in (Get-Content -LiteralPath $arqPerm -Encoding UTF8)) {
        $l = "$l".Trim()
        if (-not $l -or $l.StartsWith('#')) { continue }
        $permitidos += $l
    }
}

function Test-Permitido {
    param([string]$Texto)
    foreach ($p in $permitidos) {
        try { if ($Texto -match $p) { return $true } } catch { }
    }
    return $false
}

# ---------------------------------------------------------------------
$extensoes = '.ps1', '.psm1', '.psd1', '.cmd', '.bat', '.md', '.txt', '.json',
             '.xml', '.csv', '.yml', '.yaml', '.py', '.js', '.reg', '.ini'

Write-Host ''
Write-Host "== varrendo $Raiz" -ForegroundColor Cyan
Write-Host "   $($permitidos.Count) padrao(oes) liberado(s) em tools\vazamento-permitido.txt"

$arquivos = @(Get-ChildItem -LiteralPath $Raiz -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object {
        $_.FullName -notmatch '\\\.git\\' -and
        $_.FullName -notmatch '\\_montagem\\' -and
        $extensoes -contains $_.Extension.ToLower()
    })
Write-Host "   $($arquivos.Count) arquivo(s) de texto"
Write-Host ''

$achados = @()
$liberados = 0

foreach ($f in $arquivos) {
    $linhas = $null
    try { $linhas = Get-Content -LiteralPath $f.FullName -Encoding UTF8 -ErrorAction Stop } catch { continue }
    $rel = $f.FullName.Substring($Raiz.Length).TrimStart([char]92)

    for ($i = 0; $i -lt $linhas.Count; $i++) {
        $linha = "$($linhas[$i])"
        if (-not $linha.Trim()) { continue }

        # Linha marcada como exemplo ficticio. Existe porque a documentacao PRECISA
        # mostrar a forma de um identificador para explicar o que o varredor procura,
        # e mostrar a forma acende o varredor. O marcador e explicito e fica na
        # propria linha: quem le sabe na hora por que aquilo foi pulado, sem ter de
        # cruzar com uma lista de excecoes em outro arquivo.
        #
        # So vale para exemplo INVENTADO. Nome de host, IP ou dominio de verdade
        # marcado assim e uma mentira escrita a mao, e o varredor nao tem como
        # saber - por isso o marcador e longo e feio, para nao ser usado de leve.
        if ($linha -match 'EXEMPLO-FICTICIO') { $liberados++; continue }

        foreach ($c in $CATEGORIAS) {
            $ms = $null
            try { $ms = [regex]::Matches($linha, $c.Padrao) } catch { continue }
            foreach ($m in $ms) {
                if (Test-Permitido $m.Value) { $liberados++; continue }
                $achados += [pscustomobject]@{
                    Arquivo   = $rel
                    Linha     = $i + 1
                    Categoria = $c.Nome
                    Achado    = $m.Value
                    Porque    = $c.Porque
                    Contexto  = $linha.Trim()
                }
            }
        }
    }
}

if ($Detalhado) { Write-Host "   $liberados ocorrencia(s) liberada(s) pela lista de excecoes"; Write-Host '' }

if ($achados.Count -eq 0) {
    Write-Host 'LIMPO: nenhum identificador de infraestrutura de terceiro.' -ForegroundColor Green
    Write-Host ''
    exit 0
}

Write-Host "$($achados.Count) ACHADO(S):" -ForegroundColor Red
Write-Host ''
foreach ($g in ($achados | Group-Object Categoria)) {
    Write-Host ("  $($g.Name)  -  $($g.Group[0].Porque)") -ForegroundColor Yellow
    foreach ($a in $g.Group) {
        Write-Host ("     $($a.Arquivo):$($a.Linha)  ->  $($a.Achado)")
        $ctx = $a.Contexto
        if ($ctx.Length -gt 110) { $ctx = $ctx.Substring(0, 110) + '...' }
        Write-Host ("        $ctx") -ForegroundColor DarkGray
    }
    Write-Host ''
}
Write-Host 'Cada achado tem dois destinos possiveis, e nenhum deles e ignorar:' -ForegroundColor Yellow
Write-Host '  1. sai do codigo - vira parametro, ou e descoberto em tempo de execucao'
Write-Host '  2. e legitimo, e entra em tools\vazamento-permitido.txt COM O MOTIVO escrito'
Write-Host ''
exit $achados.Count
