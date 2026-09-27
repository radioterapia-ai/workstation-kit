# Workstation Kit — a versão universal

Aplicativo PowerShell + WinForms que prepara, diagnostica e otimiza estações de
trabalho. Roda **sem privilégio de administrador**. Serve qualquer computador — de
clínica, de consultório, de recepção ou pessoal —, em qualquer país, preservando
software de radioterapia de qualquer fabricante.

Fork do `WORKSTATION_RT`, aberto em 27/09/2026. Aquele projeto continua existindo e
continua sendo a ferramenta da clínica do autor. Este é o irmão que não conhece
clínica nenhuma.

Autor: médico radio-oncologista e especialista em qualidade, não é da TI.

---

## O que este projeto herdou, e o que herdar significa

Herdou o **código** e, mais importante, herdou **seis defeitos já pagos**. O
`WORKSTATION_RT` levou meses de ciclo «rodar numa máquina real → trazer o log →
corrigir o que o log expôs», e quase todo defeito sério saiu daí, não de revisão.

> **A regra que atravessou o fork inteira: desconfie de toda mensagem de sucesso
> que não verifica o resultado.**

O caso fundador: durante semanas o kit imprimiu *«Atalhos copiados para a Área de
Trabalho!»* em todas as máquinas, sem ter copiado nada — o caminho estava errado e
o erro do `xcopy` estava silenciado. Só apareceu quando alguém acrescentou um
contador honesto.

`docs/REGRAS-DE-SEGURANCA.md` tem o resto. Leia antes de mexer nas listas.

---

## Regra número um

**Nada neste aplicativo pode exigir administrador.** Não é preferência, é o escopo.
Funcionalidade que só funciona elevada não entra no kit — vira achado no
diagnóstico, com o texto que o usuário leva ao chamado.

## Regra número dois

**Nada que alguém esteja usando pode ser encerrado.** No projeto de origem isso se
dizia «nada clínico», porque a máquina podia estar com uma sessão de planejamento
de tratamento aberta. Aqui vale mais largo e pelo mesmo motivo: a máquina pode
estar tratando paciente, atendendo alguém no balcão, ou rodando o lote de
processamento de um consultório.

## O momento é o mecanismo principal, não as listas

**Este é um app de Windows, não um app de radioterapia.** Ele deixa a máquina leve
para o trabalho do dia começar — e por isso é para rodar **logo após logar**, antes
de abrir o ARIA, o MOSAIQ, o Monaco, o Eclipse.

Rodar no **meio do dia** pode encerrar trabalho em curso.

> As listas de proteção existem para esse caso: elas são a **rede**, para quem
> rodar fora de hora. O mecanismo principal é `Get-MomentoSessao`, que diz que o
> momento está errado, com o motivo na frente, e deixa a pessoa decidir.

Isso reordena a arquitetura de segurança, e vale registrar por que: eu vinha
tratando as listas como o mecanismo principal, e concluí que a versão universal
precisaria conhecer o software de **todo fabricante do mundo** para a promessa se
sustentar. Estava errado pela raiz. Rodando na hora certa, o trabalho ainda não
está aberto, e não há o que proteger.

### Os três sinais, e por que exigem janela

| Sinal | O que é |
|---|---|
| **idade da sessão** | o `explorer` **desta** sessão. Não é uptime da máquina — uptime não distingue "acabei de logar" de "estou aqui desde as 7h" |
| **carga em andamento** | processo classificado `Ecossistema`: há lote rodando |
| **app de trabalho aberto** | e **com janela** |

**Janela é exigida de propósito.** Agente, bandeja e serviço sobem no logon sem
ninguém pedir; contar isso faria o aviso disparar em **toda** execução, e aviso que
dispara sempre é aviso que ninguém lê. Este mesmo arquivo já tem um caso assim — o
alerta de Wi-Fi, que num notebook dispara toda vez.

**E navegador e Office não contam.** Eles voltam sozinhos no logon por restauração
de sessão; estarem abertos não quer dizer que o trabalho começou. Continuam
protegidos de encerramento — são duas perguntas diferentes, e `$script:AppsDeTrabalho`
responde a segunda sem interferir na primeira.


Quatro listas governam isso, no topo do `.ps1`:

| Lista | O que faz |
|---|---|
| `$script:Protegidos` | nunca encerrado, em nenhuma circunstância |
| `$script:SessaoPreservar` | preservado na otimização de sessão |
| `$script:SessaoRemoto` | acesso remoto — encerrar derruba o próprio atendimento do TI |
| `$script:SessaoNaoCompactar` | não tem a memória compactada (causaria engasgo na tela) |

`SessaoPreservar` e `SessaoNaoCompactar` **terminam concatenando** `SessaoRemoto`.
Quem mexe numa mexe nas três.

---

## Idioma: a distinção que não se deve perder

Duas camadas, duas respostas:

| Camada | Idioma |
|---|---|
| **Texto que o usuário lê** na tela, no log e no `.zip` | inglês, português e espanhol |
| **Código, comentário e documentação** deste repositório | português do Brasil |

O segundo é a regra do ecossistema (`C:\AI_PROJETOS\CLAUDE.md`) e não muda por o
produto ser internacional: quem mantém este código é brasileiro, e a documentação
existe para ele e para o agente que trabalha com ele.

> **Não «traduza o projeto». Traduza a interface.** São coisas diferentes, e
> confundi-las produziria um repositório em inglês ruim e uma documentação que o
> autor lê mais devagar.

### O defeito de idioma que já custou caro

O aplicativo **compara texto que o Windows devolve**. Uma versão anterior filtrava
a saída de `sc query` pela palavra `STATE` — e o Windows em português devolve
`ESTADO`. Resultado: o dossiê relatou **zero agentes de segurança** em toda estação
pt-BR, sem erro nenhum, sem nada no log.

> **Cmdlet que devolve OBJETO é seguro. Comando que devolve TEXTO é localizado, e
> comparar esse texto é um defeito que só aparece na máquina de outra pessoa.**

Prefira `Get-Service` a `sc query`, `Get-CimInstance` a `wmic`,
`[Environment]::GetFolderPath()` a nome de pasta escrito à mão. Quando não houver
alternativa, compare pela parte que **não** se traduz — um código numérico, uma
palavra-chave em inglês que o Windows nunca localiza.

---

## Estrutura

```
WorkstationKit.ps1              o aplicativo inteiro (~5.470 linhas)
Start_WorkstationKit.cmd        o atalho que abre o app
docs/                           documentação
tools/                          suíte de teste e empacotador
```

O `.ps1` é **um arquivo só, e isso não se negocia.** A distribuição é cópia de
arquivo: no projeto de origem, atualizar a cópia da rede era o que distribuía a
correção para 11 máquinas sem ninguém passar de mesa em mesa. Dividir em módulos
quebraria o rito que faz a correção chegar.

> **Não proponha separar o `.ps1` em módulos.** Já foi decidido, e a razão é de
> distribuição, não de gosto.

---

## Como trabalhar neste código

### Sempre rode a suíte antes de entregar

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Sessao.ps1
```

Ela valida o parse, as duplicatas, o BOM, **confere que as sete listas de proteção
carregaram íntegras**, e exercita as proteções contra processos de verdade.

Os blocos são **testes negativos**: cada um monta o caso em que a proteção precisa
falhar se estiver quebrada. Foi assim que se achou o `[\\/]` que virou `[/]` e
fazia `Test-CaminhoEcossistema` devolver falso para **todo** caminho do Windows — o
parse passava, a regex compilava, e a proteção simplesmente não acontecia.

> **Teste positivo não vê proteção que parou de proteger.** Se um bloco passar a
> falhar, leia o nome dele antes de mexer no teste.

O bloco 11 lista o que **ainda não** está coberto, e isso é de propósito: é a lista
de trabalho da versão universal, registrada em vez de escondida.

### Ao mexer nas listas de processos

Escreva o caso de teste **antes** de acrescentar a entrada na lista. É o caso que
faz aparecer a entrada que não casa — nome curto demais, âncora faltando, grafia
diferente. Lista primeiro e teste depois só confirma o que já se acreditava.

**Nome de regex casa demais, e isso é o erro mais comum aqui.** Exemplo real deste
código: `ARIA` (o sistema oncológico da Varian) também casa com `Soarian` (o
prontuário da Siemens), porque não tem âncora. Nesse caso protege por acidente — a
direção segura —, mas o mesmo mecanismo pode casar com algo que **deveria** poder
ser encerrado. Nome de três ou quatro letras numa lista de proteção quase sempre
precisa de `^` e `$`.

### Ao otimizar qualquer coisa

O Módulo 5 ficou 10× mais rápido guardando o que antes consultava por processo. A
conta fechou e vieram seis defeitos juntos, todos iguais: **dado congelado usado
para decidir**.

> **Otimização não muda o que o código faz; muda *quando* ele sabe. E toda proteção
> que dependia de saber passa a depender de quando soube.**

Instantâneo tem prazo, quem decide lê ao vivo, e consulta que falhou não se guarda.
A regra inteira em `docs/REGRAS-DE-SEGURANCA.md`.

**E meça antes de reescrever.** No projeto de origem, duas hipóteses sobre onde
estava o tempo estavam erradas: a acumulação `+=` custava 29 ms e não segundos, e
os regex custavam 75 ms. O gargalo era outro.

### Ao editar o PowerShell

PowerShell **5.1** apenas: sem operador ternário, sem `??`, sem `-AsHashtable`.

Três regras de ambiente do ecossistema, e as três já morderam projetos daqui:

1. **`.cmd` e `.bat` em CRLF, sempre.** O `cmd.exe` lê o lote por offset de byte.
   Com LF ele perde o rumo e executa como comando do prompt o que vier depois. O
   sintoma não parece erro de fim de linha: parece código errado.
2. **Nada de `chcp` dentro de `.cmd`.** Cabeçalho em ASCII puro.
3. **`.ps1` com acento precisa de BOM.** Sem BOM, o PowerShell 5.1 lê como cp1252;
   em comentário dá mojibake inofensivo, **dentro de string é fatal**.

O `.gitattributes` protege o git. Não protege quem copia o arquivo à mão — e a
distribuição aqui é cópia de arquivo.

---

## O que este projeto não faz

- Não contorna antivírus. No projeto de origem o Cortex XDR bloqueou o app numa
  estação; a decisão foi **não rodar naquela máquina**, não disfarçar o script.
- Não compila código em tempo de execução. Havia um caminho que fazia isso; foi
  removido porque era desnecessário e porque compilar em runtime é comportamento
  que antivírus tratam como suspeito, com razão.
- Não apaga arquivo fora de pasta de cache do próprio perfil. O Módulo 4 lista e dá
  veredicto; quem apaga é o usuário, pelo Explorer.
- Não desinstala nada, não mexe em serviço, não altera HKLM.

---

## Documentação

| Arquivo | Para quê |
|---|---|
| `docs/PROVENIENCIA.md` | de onde veio, o que saiu, e por que sem histórico compartilhado |
| `docs/REGRAS-DE-SEGURANCA.md` | o que nunca se toca e por quê |
| `docs/PENDENCIAS.md` | o trabalho da versão universal, em ordem |

| Ferramenta | Para quê |
|---|---|
| `tools/Testar-Sessao.ps1` | a suíte: parse, BOM, integridade das listas, proteções |
| `tools/Verificar-Vazamento.ps1` | nenhum identificador de infraestrutura de terceiro |

---

## Dado de paciente, credencial e identificador

**Nada disso existe neste repositório, e não pode passar a existir.** A pasta
sincroniza com o Google Drive, e `.gitignore` protege o git — não protege o Drive.

Este projeto nasceu com uma varredura: o `.ps1` de origem tinha **47 ocorrências**
de identificador de infraestrutura de um hospital — domínio, fileserver, servidor
de scanner, dois IPs internos, cinco hostnames e o caminho de um compartilhamento
de departamento. Nenhuma atravessou, e o repositório foi criado **sem histórico
compartilhado** justamente porque histórico não se desfaz depois.

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Verificar-Vazamento.ps1
```

Confere por **categoria** — a forma de um identificador de rede: IP privado, UNC
para host nomeado, domínio com TLD, endereço de e-mail. **Nunca por lista de
nomes:** uma lista de nomes num arquivo versionado põe no repositório exatamente o
que ela existe para manter fora, e foi assim que a primeira versão desta
conferência nasceu errada.

O cabeçalho do varredor declara os furos que ele tem — nome corporativo sem TLD,
UNC de host curto, arquivo binário, nome de pessoa. **Varredor que promete o que
não cumpre é pior que varredor nenhum**, porque produz a sensação de ter conferido.
O nome sem TLD é o furo que já morreu na prática: um parágrafo da própria
`PROVENIENCIA.md` escreveu o domínio do hospital como exemplo do que o varredor não
pega, e só um `grep` à mão achou.

Ver `docs/PROVENIENCIA.md`.
