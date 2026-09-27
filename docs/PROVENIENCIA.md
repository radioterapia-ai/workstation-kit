# Proveniência

Este projeto é um fork do `WORKSTATION_RT`, aberto em **27/09/2026**.

O original continua existindo, continua sendo mantido, e continua sendo a
ferramenta da clínica do autor. Não foi substituído: ganhou um irmão.

| | |
|---|---|
| **Origem** | `C:\AI_PROJETOS\REDE_AMERICAS\WORKSTATION_RT` |
| **Commit de origem** | `ad13e33` — *Entrega v3.1: empacotador que confere o próprio pacote* |
| **Versão de origem** | 3.1 |
| **Versão deste projeto** | 1.0 |

---

## Por que SEM histórico compartilhado

Um `git clone` teria preservado `git log`, que é como se descobre por que uma linha
é do jeito que é. Foi tentador, e não foi feito.

A varredura do `.ps1` de origem encontrou **47 ocorrências** de identificador de
infraestrutura de um hospital real:

- o domínio corporativo
- o nome do fileserver, e o caminho do compartilhamento do departamento
- o nome do servidor de scanner
- **dois endereços IP internos**
- cinco hostnames de servidor, por nome completo
- três endereços de portal de prontuário

> **Um clone gravaria a topologia de rede interna de um hospital na história deste
> projeto, para sempre. E história não se desfaz depois.**

Apagar o arquivo num commit seguinte não resolve: `git log -p` continua mostrando.
E este projeto existe para servir «qualquer computador de qualquer país», o que
significa que ele plausivelmente sai da máquina do autor um dia.

A escolha, então, foi a que erra na direção reversível: **repositório novo**,
proveniência escrita aqui. Perde-se o `git log`, que continua a um comando de
distância no projeto de origem. Não se perde a chance de nunca ter publicado o
que não devia.

> É a mesma família da regra do ecossistema sobre a pasta sincronizada: **não
> existe exclusão confiável por padrão, então a proteção é separação física.**

---

## O que saiu, e o que ficou no lugar

### Saiu inteira: a rede daquele hospital

| O que era | O que é agora |
|---|---|
| `$script:PastaClinicaRede` — UNC do fileserver | `''`, e é opcional |
| `$script:PastaRelatorios` — UNC de uma pasta de logs | `[Environment]::GetFolderPath('MyDocuments')` |
| `$script:CitrixStores` — 1 IP interno e 2 hostnames | `@()` |
| `$script:TasyUrls` — 3 endereços de prontuário | `@()` |
| o domínio corporativo, chumbado em dois lugares como sufixo de DNS | `Get-SufixosDns`, que descobre o domínio real da máquina |

O sufixo de DNS merece nota: o código **já descobria** o domínio da máquina por
duas vias (`IPGlobalProperties().DomainName` e `USERDNSDOMAIN`), e o domínio
chumbado era um terceiro palpite que só servia naquela rede. Em qualquer outra, é
uma consulta de DNS que sempre falha — e, pior, **o nome de um servidor de terceiro
sendo consultado por uma máquina que não é dele**.

Agora há também `$script:SufixosDnsExtra`, vazio por padrão: quem tem um sufixo que
não aparece nem no adaptador nem em `USERDNSDOMAIN` preenche.

### Saiu inteiro: o batch embutido do Módulo 1

**359 linhas** de batch viviam dentro do `.ps1`, como here-string, montando a
árvore de trabalho de um hospital: portais em Excel copiados de um
compartilhamento, atalhos na Área de Trabalho, fonte de código de barras,
mapeamento de unidade de rede, e um passo 10 que copiava a pasta da rede para o
`C:`. Era o Módulo 1.

Nada disso é universal, e sair levou embora o domínio, o fileserver e os dois IPs.

**O que ficou no lugar** é o caminho que o próprio código já tinha: sem rotina
embutida, o Módulo 1 pede um `.bat` de preparação ao usuário. Na versão universal é
isso que ele é — *«rode o seu script de preparação»* —, e é uma função legítima em
qualquer máquina.

#### Um defeito que a remoção quase criou

`Get-ConteudoPreparador` passando a devolver string vazia fazia `Save-Preparador`
**gravar um `.bat` de zero byte e devolver sucesso**. Um `.bat` vazio roda, sai com
código 0 e não imprime nada: o Módulo 1 diria *«Executando: ...»* e mais nada, e
quem lesse o log concluiria que a preparação rodou.

Era o defeito fundador deste projeto se repetindo dentro da própria remoção. A
guarda em `Save-Preparador` — conteúdo vazio devolve `$null` — existe por isso.

### Virou parâmetro em vez de sair

`$script:MarcaEcossistema` continua valendo `'RADIOTERAPIA_AI'` por padrão, mas o
conceito é geral e vale para qualquer máquina:

> **existe uma árvore de trabalho cuja carga nunca deve ser encerrada, e ela se
> reconhece pelo CAMINHO do processo, não pelo nome.**

Um consultório que roda o próprio lote de processamento tem a mesma necessidade.
Vazio, desliga a proteção por marca sem quebrar nada: `Test-CaminhoEcossistema`
passa a responder «não sei» para todo caminho, que é o padrão seguro dela.

A seção 4G — o acordo `em_uso.json` com o Local Suite — ficou. O conceito também é
geral: *«um programa está trabalhando agora, não o encerre»*. O nome do arquivo e
da pasta seguem a marca, então mudam junto com o parâmetro.

---

## O que valeu mais a pena herdar

Não foi o código. Foi `tools/Testar-Sessao.ps1`.

A suíte tem 36 casos, e cada bloco nasceu de um defeito real. Todos são de
**arquitetura**, não de clínica, e por isso valem igual aqui:

1. caminho ao vivo vence mapa congelado
2. a sinalização só sabe acrescentar proteção
3. união das duas leituras, porque cada uma sozinha deixa buraco
4. o instantâneo tem prazo, e consulta que falhou não se guarda
5. a lista de serviços não pode emudecer
6. o que nunca pode mudar

No fork ela ganhou três blocos novos — fabricantes de radioterapia, navegadores do
mundo, e o que **ainda não** está coberto — e duas correções de si mesma, ambas
encontradas ao portá-la:

- **Ela se recusava a rodar** quando havia um `em_uso.json` de verdade, para não
  mexer nele. Intenção certa, efeito ruim: suíte que não roda quando o produto está
  rodando é suíte que não roda. E foi o que aconteceu na prática. Agora trabalha
  numa cópia isolada no `TEMP`.
- **Ela carregava as listas na ordem dos nomes**, não na ordem do arquivo. Como
  `SessaoPreservar` e `SessaoNaoCompactar` terminam concatenando `SessaoRemoto`, a
  ordem certa estava lá **por sorte**. Agora carrega na ordem do arquivo, que é a
  ordem em que o aplicativo executa, e a dependência resolve por construção.

E ganhou a guarda que faltava: **alternativa vazia num regex de alternância casa com
qualquer coisa.** `'Spotify' -match 'a||b'` devolve `True`. Numa lista de proteção
isso vira «protege tudo», e a suíte inteira passaria — pelo motivo errado, porque
nada poderia ser encerrado. É o mesmo modo de falha do `[\\/]`, com o agravante de
falhar na direção **segura**: nada quebra, e os testes ficam verdes.

---

## O que NÃO veio, e por quê

| Não veio | Por quê |
|---|---|
| `docs/PARQUE-E-ACHADOS.md` | inventário de 11 máquinas reais, com hostnames |
| `docs/AMBIENTE-CLINICO.md` | servidores, Citrix, prontuário e agentes daquele hospital |
| `docs/TRANSPLANTE.md` | caminhos e rito daquele parque |
| `docs/JORNADA.md` | a história é valiosa e está cheia de detalhe daquela clínica |
| `referencia/` | scripts originais do usuário e do TI daquele hospital |
| `_arquivo/` | decisões de orquestração de outro momento do ecossistema |

As lições desses documentos que valem aqui foram reescritas, não copiadas — em
`REGRAS-DE-SEGURANCA.md` e no `CLAUDE.md` deste projeto. Reescrever custou mais e é
a única forma de garantir que nada atravessou junto.

---

## Como conferir que nada atravessou

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Sessao.ps1
```

E, para a varredura de identificadores:

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Verificar-Vazamento.ps1
```

Ele confere por **categoria**, não por lista de nomes — procura a *forma* de um
identificador de rede (IP privado, UNC para host nomeado, domínio com TLD,
endereço de e-mail), não o nome de ninguém.

> A primeira versão desta conferência era um `git grep` com os nomes do hospital
> escritos dentro dele. **Isso se anula:** para conferir que os nomes não estão no
> repositório, os nomes passavam a estar no repositório — dentro do próprio comando
> de conferência, em documento versionado. O varredor por categoria não tem esse
> defeito, e ainda pega vazamento que ninguém previu.

O varredor tem um limite declarado no próprio cabeçalho: **nome corporativo sem TLD
não é pego por nenhuma categoria**, porque o nome de uma instituição, escrito sem o
sufixo de domínio, tem a forma de qualquer palavra. Não há como pegar por forma.

Esse resíduo é responsabilidade de quem edita, e é o motivo de este documento
descrever o que saiu **sem escrever os nomes**.

> A primeira versão deste próprio parágrafo escrevia o domínio do hospital como
> exemplo do que o varredor não pega. Explicar a regra violando-a — e o `grep` de
> conferência achou, no mesmo dia. O limite do varredor é real: quem edita é a
> última barreira, e a documentação é o lugar mais fácil de esquecer disso.
