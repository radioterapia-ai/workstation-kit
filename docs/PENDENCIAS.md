# Pendências da versão universal

Estado em **27/09/2026**, dia da abertura do fork.

Este arquivo distingue duas coisas, e a distinção importa:

- **VERIFICADO** — medido ou testado nesta máquina, com o número ao lado
- **A LEVANTAR** — ainda não conferido; está aqui para não ser esquecido

Nada aqui é estimativa apresentada como fato. Quando um item passar de «a levantar»
para «verificado», ele ganha o número e perde a marca.

---

## 1. As listas de proteção — VERIFICADO

`tools/Testar-Sessao.ps1`, bloco 11, mede isso a cada execução e imprime a lista.
**23 entradas por cobrir**, na última execução.

O bloco não conta como falha de propósito: conta como trabalho. Uma suíte que
falhasse por isso seria ignorada em uma semana.

| Grupo | Faltando hoje |
|---|---|
| acesso remoto | `Parsec`, `ToDesk` |
| navegador | `UCBrowser`, `QQBrowser`, `Whale`, `Maxthon`, `Sleipnir`, e o genérico `browser` |
| planejamento | `Pinnacle`, `Oncentra`, `iPlan`, `Precision`, `RayCare`, `XiO` |
| prontuário | `Epic`, `PowerChart`, `MEDITECH`, `Sectra` |
| escritório | `soffice`, `wps`, `et`, `wpp`, `Hwp` |

Já cobertos, e a suíte confirma: `AnyDesk`, `RustDesk`, `ScreenConnect`,
`Splashtop` (via `$script:SessaoRemoto`), e os oito fabricantes de radioterapia da
clínica de origem.

### O problema de âncora — VERIFICADO

`Soarian` (prontuário da Siemens) casa com a entrada `ARIA` (sistema oncológico da
Varian), porque `ARIA` não tem âncora. Medido: casa em `Protegidos` **e** em
`SessaoPreservar`.

Aqui protege por acidente, que é a direção segura. Mas o mecanismo é o mesmo que
faria uma entrada curta casar com algo que **deveria** poder ser encerrado.

> **Nome de três ou quatro letras numa lista de proteção casa demais.** `et`, o
> processo de planilha do WPS Office, é o caso extremo: acrescentá-lo sem âncora
> casaria com dezenas de processos. Provavelmente precisa ser `^et$`.

**A fazer:** revisar as entradas curtas de todas as sete listas e decidir âncora
caso a caso. Escrever o teste antes de mexer na lista, sempre.

### O que a pesquisa deve acrescentar — A LEVANTAR

Um levantamento em curso cobre fabricantes de radioterapia no mundo (Varian,
Elekta, RaySearch, Accuray, Brainlab, Philips, Canon, Sun Nuclear, IBA, ViewRay,
Mirada, Limbus, MVision, Radformation e outros), navegadores regionais, acesso
remoto, suítes de escritório e segurança corporativa.

Cada nome que vier entra **primeiro** como caso na suíte, com a fonte registrada, e
só depois na lista. E cada um vem com a confiança declarada: confirmado em fonte,
provável, ou palpite.

---

## 2. A camada de idioma — DECIDIDO, NÃO COMEÇADO

Decisão de 27/09/2026: **construir a camada agora, com inglês, português e
espanhol completos.** Os demais idiomas entram depois acrescentando tabela, sem
mexer no código.

O motivo de não deixar para depois: retrofit de i18n em ~800 strings é edição em
massa num arquivo de 5.470 linhas, e edição em massa é exatamente o que já
introduziu defeito neste código.

| A fazer | Estado |
|---|---|
| contar as strings visíveis de verdade | A LEVANTAR |
| escolher a forma da tabela em arquivo único | A LEVANTAR |
| resolver ordem de argumento no `-f` quando a frase muda de ordem | A LEVANTAR |
| decidir onde guardar a escolha do usuário (sem admin, sem HKLM, máquina que reseta) | A LEVANTAR |
| decidir o palpite inicial de idioma na primeira execução | A LEVANTAR |
| decidir o idioma do **texto de chamado** — o que o usuário leva ao TI | A LEVANTAR |

A última é a menos óbvia e talvez a mais importante: o texto de achado do
diagnóstico existe para ser colado num chamado. O idioma da interface é o do
usuário; o idioma do TI local pode ser outro. Pode ser que esse texto precise sair
nos dois.

### O que NÃO é camada de idioma — VERIFICADO no projeto de origem

Há uma classe de defeito que nenhuma tabela de tradução resolve: o aplicativo
**compara texto que o Windows devolve**.

Caso real: uma versão filtrava `sc query` pela palavra `STATE`, e o Windows em
português devolve `ESTADO`. O dossiê relatou **zero agentes de segurança** em toda
estação pt-BR, sem erro, sem nada no log.

**A fazer:** varrer o `.ps1` inteiro procurando toda comparação com texto que o
sistema devolve, e trocar por objeto ou por chave não-localizada. Essa varredura
está no levantamento em curso.

---

## 3. O Módulo 1 — FUNCIONA, MAS É UM COTO

Hoje: sem rotina embutida, pede um `.bat` ao usuário e o executa, com log honesto
linha a linha. É legítimo e não engana.

**A decidir:** a versão universal merece uma rotina de preparação genérica própria?
Candidatos plausíveis — verificar pastas de trabalho, conferir espaço em disco,
testar alcance de servidores configurados, criar atalhos que o usuário pedir. Ou o
Módulo 1 deixa de existir e os módulos são renumerados.

Renumerar tem custo: o usuário da clínica de origem chama as coisas de «Módulo 5»
em conversa, e o log usa esses números.

---

## 4. Os módulos e diagnósticos — A LEVANTAR

O diagnóstico de Citrix e de prontuário roda dentro do Módulo 2 e foi escrito para
um ambiente. Quanto dele é universal está no levantamento em curso.

A seção 4F (relato de registro) e a 4G (o acordo `em_uso.json`) integram o kit com
outro projeto do autor. O **conceito** da 4G é geral — «um programa está trabalhando
agora, não o encerre» — e ficou. Se vale generalizar o contrato para qualquer
aplicativo, e como, está por decidir.

---

## 5. O empacotador — NÃO PORTADO

`tools/Empacotar.ps1` existe no projeto de origem e monta a entrega conferindo o
próprio pacote: reabre o `.zip`, valida o parse do `.ps1` de dentro dele, confere o
BOM e compara o SHA256.

**A fazer:** portar, com três mudanças —

1. nomes e destino próprios (`C:\AI_DEPLOY\WORKSTATION_KIT\v<versão>\`)
2. o `LEIA-ME.txt` de dentro do `.zip` em **inglês**, porque é texto de produto
3. **a varredura de identificadores como etapa de barreira**, a mesma que a criação
   do fork usou. Se algum identificador de infraestrutura voltar por edição manual,
   o empacotamento para antes de gerar o `.zip`.

O item 3 é o que faz a promessa de `PROVENIENCIA.md` se sustentar no tempo, em vez
de valer só no dia da criação.

---

## 6. A pergunta que ainda não tem resposta

O projeto de origem promete «nada clínico é tocado», e essa promessa se sustenta em
listas conferidas contra **11 máquinas reais de um parque conhecido**.

Sem esse parque, em que a promessa passa a se sustentar?

Não é retórica. As opções parecem ser: (a) a promessa muda de forma — de «nada
clínico é tocado» para «nada que o kit não reconheça é tocado sem você marcar»; (b)
o kit passa a exigir uma etapa de reconhecimento na primeira execução, em que o
usuário confirma o que é crítico naquela máquina; ou (c) as duas.

O código já tem material para (b): o Módulo 5 desmarca por segurança o que não
reconhece, e o Módulo 3 mostra veredicto antes de agir. Falta decidir se isso basta
como promessa, e reescrever a promessa para o que ela de fato é.

Esta pergunta está no levantamento em curso, com instrução explícita de não
suavizar.

---

# Atualização de 27/09/2026 — o levantamento adversarial

6 inventários, 6 críticos encarregados de achar os erros deles, 3 pesquisas.
**139 itens, 74 erros encontrados pelos críticos.** Onde o crítico contradisse o
inventário, ficou o crítico. O que eu próprio medi está marcado **VERIFICADO**.

## O achado que muda a estratégia

> **A versão universal *é* a versão atual.** O levantamento foi procurar
> amarramento a um hospital e encontrou o contrário: das **334 entradas** nas seis
> listas, apenas **11 são locais ou regionais** — e o crítico derrubou 3 dessas 11.

Arquitetura, resolvedores de caminho e motor de proteção são **100% comuns**. O que
sai cabe numa página; o que fica é o projeto inteiro.

## Cinco coisas que pareciam peculiaridade e NÃO saem

O crítico corrigiu o inventário nestes cinco. É o erro que mata se for errado:

1. **`\radioterapia\TC_DATA`** — "radioterapia" e "TC" são vocabulário ordinário em
   português, espanhol e italiano. Uma clínica em Lisboa, Madri ou Buenos Aires que
   instala e não configura nada **perde a proteção da árvore de TC**. Fica, e o fork
   **acrescenta** `radiotherapy`, `radiothérapie`, `Strahlentherapie`, `CT_DATA`.
2. **`Quest`, `OnDemand`, `ODMActiveDirectory`** — produto comercial global. Fusão de
   clínica e rollup de consultório é exatamente quem instala.
3. **`^claude$`, `^claude-code$`** — não apagar, **generalizar**: viram
   `^python$|^pythonw$|^node$|^Rscript$|^MATLAB$|^julia$`, e em **`SessaoPreservar`**,
   não só em `SessaoNaoCompactar` — este último só evita engasgo de tela, não protege
   de encerramento.
4. **Citrix e CentBrowser** — Citrix não é peculiaridade: é a camada pela qual Elekta
   e Varian entregam aplicativo em qualquer país.
5. **Vitrea, Mirada, Medis, INVIA, Digitalcore** — Canon/Vital, Reino Unido, Holanda,
   Michigan, Tóquio. Nenhum é daquele hospital, e custam dez palavras.

## A lista por nome nunca ficará completa

A pesquisa confirmou: **nome de processo de software de radioterapia praticamente
não existe em fonte pública.** Os fabricantes publicam folheto e DICOM Conformance
Statement — não publicam executável.

Dão **MATA hoje**: Pinnacle3, TomoTherapy, Accuray, Brainlab Elements, SNC Patient,
VeriSoft, myQA, Delta4, MobiusFX, ClearCheck, Limbus, RadiAnt, Orthanc, dcm4chee.

> **Lista incompleta é o modo de falha grave do enunciado.** A resposta são quatro
> mecanismos, e três não dependem de conhecer o fabricante:

| | Mecanismo | Custo |
|---|---|---|
| **A** | nome de processo, ancorado onde é curto | o que já existe |
| **B** | **pasta do fabricante** — `\Varian\`, `\CMS\`, `\Elekta\`, `\RaySearch\`, `\Accuray\`, `\Brainlab\` | é o mesmo `Test-CaminhoEcossistema` que **já existe e já funciona** |
| **C** | **publisher do executável** — `CompanyName` ou `Get-AuthenticodeSignature` | leitura de arquivo, **não exige admin** |
| **D** | **gerenciador de licença** — `hasplms`, `CodeMeter`, `lmgrd`, `FNPLicensingService` | o achado mais transversal: matar o gerenciador **não derruba o app na hora**; derruba no próximo checkout, com mensagem que fala de licença e nunca de processo encerrado |

Dois casos que a pesquisa marcou como os mais perigosos:

- **Limbus Contour / Radformation AutoContour** rodam autocontorno por rede neural
  **dentro da estação, em CPU, por 1 a 3 minutos**. Para um otimizador isso tem
  exatamente a assinatura de "processo travando a máquina". É o candidato número um
  a ser morto por engano, e o que ele interrompe é contorno de plano em andamento.
- **Sectra IDS7 é ClickOnce**: mora em `%LOCALAPPDATA%\Apps\2.0`. **Um módulo que
  limpe cache do próprio perfil pode apagar a instalação da estação de laudo.** Isso
  é risco do Módulo 3/4, não de lista de processo.

## Seis categorias universais que nenhuma lista cobre, todas MATA hoje

Todas rodam **na sessão do usuário** — o Módulo 5 as alcança de verdade:

1. **Ditado médico** — `natspeak` (Dragon Medical One), `DragonBar`. É o teclado do
   rádio-oncologista no mundo inteiro; encerrar no meio do laudo perde o texto. Hoje
   `nssystem` sobrevive **por acidente**, casando o token largo `System`.
2. **Assinatura digital e smartcard** — `SafeSignIC`, `SafeNetAuthentication`. Sem o
   middleware não se assina aprovação de plano nem prescrição.
3. **Acessibilidade** — `nvda`, `jfw`, `Narrator`. Leitor de tela roda sem janela
   principal, logo vem **pré-marcado**. Encerrar deixa o usuário cego sem interface.
4. **IME de terceiro** — `SogouInput`, `Baidu`, Google Japanese Input. Num app com
   seletor de idiomas, matar o IME é um app que não funciona em três deles.
5. **Backup** — `VeeamAgent`, `cvd`, **`mms` (Acronis — três letras)**. Prova de que a
   cobertura atual é acidental: **`AcronisAgent` sobrevive hoje porque casa o token
   `Onis`, do visualizador DICOM.**
6. **VPN além das quatro listadas** — `openvpn`, `wireguard`, `NetExtender`. Derrubar
   corta o acesso ao PACS **e a própria sessão pela qual o kit está sendo rodado**.

## Três achados que mudam decisão

- **Splashtop não se chama Splashtop**: `SRService`, `SRManager`, `strwinclt`. Quem
  procurar "Splashtop" conclui, errado, que a máquina não tem acesso remoto.
- **Cybereason não tem "Cybereason" no processo principal**: é `minionhost.exe`.
- **`^msedge` sem âncora final** cobre `msedgewebview2`, que hospeda o Outlook novo e
  o Teams novo. Sem isso, encerrar `msedgewebview2` fecha a janela do aplicativo
  corporativo.

## O turco, que não é sobre oferecer turco

**Medido em PowerShell 5.1 sob `tr-TR`: `'CITRIX' -match 'citrix'` devolve `False`.**

`-match`, `-like`, `-replace`, `-split` e `switch -Regex` herdam `IgnoreCase` **sem**
`CultureInvariant` e dobram a caixa pela cultura da thread. São **140 pontos** no
arquivo, e a falha é **fail-open e silenciosa**: a linha é
`if ($p.ProcessName -match $script:Protegidos) { continue }` — match falho significa
**não pula, encerra**.

> **Isto não depende de o fork oferecer turco.** Uma estação turca quebra o app de
> hoje, em português, sem ninguém ter escolhido nada. Incluir turco é o que
> **obriga** a corrigir, e a correção protege todos os outros idiomas.

## Riscos confirmados no código

**R4 — o caminho que APAGA não consulta veredicto nenhum.** VERIFICADO:
`$script:RaizesClinicas` só é usado no Módulo 4, que é **somente leitura**.
`Clear-Alvo` tem três guardas, e **nenhuma é `RaizesClinicas`**. O fork vai ampliar a
lista de alvos — se ampliar acreditando que "o motor de veredicto pega dado clínico",
não pega.

**R5 — o grupo TAREFA não consulta lista nenhuma.** `Get-TarefasEmExecucao` filtra só
`\Microsoft\Windows\*` e a marca. O que segura o estrago hoje é a regra número um
(tarefa como SISTEMA falha por permissão) — mas tarefa no contexto do usuário, que é
como muito export DICOM é instalado, **para sem resistência**.

**R6 — `Conhecido` vence janela aberta.** VERIFICADO e **parcialmente corrigido**.
`$marcar = ((-not TemJanela) -or Conhecido) -and (-not pesadoDesconhecido)`, e
`pesadoDesconhecido` exige `-not Conhecido`. Logo `Conhecido = $true` **força** a
marcação mesmo com janela aberta e acima de 300 MB.

Medido com janela aberta e 800 MB, **antes** da correção:

```
Epic                 conhecido=True  MARCADO=True   casa 'Epic'
ElektaNotification   conhecido=True  MARCADO=True   casa 'Notification'
MyVendorUpdater      conhecido=True  MARCADO=True   casa 'Update'
```

E a nota que o usuário lê ao lado: *"Tem janela aberta, mas é programa conhecido e
não guarda documento."* Falsa nas duas afirmações para um prontuário.

**Corrigido em 27/09/2026:** `Epic` ancorado em `EpicGamesLauncher|EpicWebHelper` — o
token existia para o lançador de jogo e casava a Epic Systems. E Epic, Cerner
(`PowerChart`), MEDITECH, Soarian e Sectra entraram em `SessaoPreservar`.

**NÃO corrigido, e é decisão sua:** `ElektaNotification` e atualizadores de
fabricante seguem pré-marcados. "Conhecido vence janela aberta" é deliberado para
Zoom, Slack e Teams — o comentário do autor diz isso com todas as letras. Trocar é
**desenho, não defeito**.

A proposta do levantamento é **inverter só na marcação**: desconhecido vem
desmarcado e apresentado. E **não** inverter o `return $true` final de
`Test-PodeEncerrarSessao`, que responde "há impedimento?" e tem três consumidores com
significados diferentes — inverter ali faria o executor imprimir "não foi possível
encerrar" para item que o usuário marcou conscientemente, e o grupo PRIORIDADE
pararia de funcionar **continuando a imprimir "Prioridade ajustada"**.

## A pergunta honesta, e a resposta que o levantamento defende

A promessa de hoje tem **três pernas**, não uma:

1. as listas conferidas contra 11 máquinas reais — **é a que cai sem o parque**
2. **a regra número um**: sem admin, e o Módulo 5 só enxerga a sessão do usuário.
   `sqlservr`, `Orthanc`, `dcm4chee`, `dbsrv17` são serviços em sessão 0 — **o Módulo
   5 nunca os alcança**. Boa parte da promessa é cumprida pelo **escopo**, não pela
   lista
3. os mecanismos independentes de nome, que ficam **mais fortes** no fork

> Sem o parque, a frase "nada clínico é tocado" passa de *"conferimos"* para
> *"achamos"* — e é exatamente o tipo de mensagem de sucesso que não verifica o
> resultado. Um app distribuído no mundo que diga isso apoiado numa lista que não
> conhece Pinnacle, TomoTherapy, Brainlab, Orthanc nem Dragon **está mentindo com boa
> intenção**, e a primeira clínica que descobrir isso descobre da pior maneira.

A frase proposta para o fork, que deixa de ser afirmação sobre o resultado e passa a
ser **propriedade verificável do código**:

> **Este kit nunca encerra, apaga ou rebaixa nada que ele não tenha mostrado a você
> antes — marcado ou desmarcado, com o motivo escrito ao lado. O que ele reconhece
> como clínico, ele protege sem perguntar, por quatro caminhos independentes: o nome
> do programa, a pasta do fabricante, a assinatura do arquivo e um sinal que o
> próprio programa pode deixar. O que ele não reconhece, ele não toca por conta
> própria. A lista de fabricantes que ele conhece está aberta e acrescentar o seu leva
> uma linha. Ele não pede administrador, e por isso não alcança serviço do sistema —
> o que protege o seu banco de dados não é a nossa lista, é o escopo.**

**E a versão local NÃO deve adotar essa promessa.** Ela tem parque, tem 11 máquinas,
tem log, tem o método de rodar-trazer-corrigir. Pode continuar prometendo mais — e o
momento em que as duas passam a prometer coisas diferentes é o momento em que a
separação vale a pena, não uma perda.

Existe caminho de volta: o Módulo 2 já produz inventário. Um dossiê **anônimo e
opt-in** que o usuário escolha enviar é o que reconstrói o parque, agora mundial. Aí,
e só aí, a versão universal pode voltar a prometer o que a local promete hoje.

## A ordem proposta

| Fase | O quê |
|---|---|
| **0** | corrigir **no local** o que é defeito universal — todo defeito corrigido antes da bifurcação é corrigido uma vez, depois duas |
| **1** | a bifurcação mecânica — **feita em 27/09/2026** |
| **2** | **primeiro entregável que serve**: universal em inglês, sem Módulo 1, listas ampliadas, mecanismos B/C/D |
| **3** | Módulo 1 universal, em PowerShell nativo |
| **4** | camada de idioma com **dois** idiomas (`en` e `pt`) — o `pt` valida a tabela contra o app real e custa zero, porque já está escrito |
| **5** | os idiomas restantes, dois por vez |

**Fora de escopo por decisão registrada: árabe e hebraico.** Não por mercado — por
RTL. O conteúdo deste app é quase todo latino e não traduzível (caminho do Windows,
nome de processo, versão), e o algoritmo bidirecional **reordena visualmente** esses
trechos dentro de parágrafo RTL. Caminho com barra invertida e dígito no meio de
texto árabe fica ilegível, e este é um app de caminhos, onde o usuário lê o caminho e
decide apagar. Mesma lógica do antivírus: não rodar naquela máquina foi melhor que
disfarçar o script.

## Sete blocos que a suíte ainda precisa ganhar

| Bloco | O que prova |
|---|---|
| cruzamento protegido × alvo | nenhum `-Caminhos` de alvo de limpeza devolve MANTER no veredicto |
| ~~configuração vazia~~ | **feito** — bloco 12 |
| catálogo de ~150 nomes esperados | transforma "esqueci a Elekta" em falha vermelha |
| cultura | rodar o bloco 7 inteiro sob `tr-TR`, `de-DE` e `ja-JP` |
| pré-marcação | desconhecido com janela, abaixo de 300 MB → exigir desmarcado |
| tarefas | tarefa em `\Vendor\` apontando para pasta de fabricante não entra |
| tradução | toda chave existe nos outros idiomas; largura medida abaixo do teto |

> E a regra de método que vale mais que qualquer bloco: **toda etapa que não roda por
> falta de configuração IMPRIME que não rodou e por quê.** Lista que encolheu em
> silêncio é o modo de falha desta casa.
