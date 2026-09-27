# Regras de segurança

O que nunca se toca, e por quê. Herdado do `WORKSTATION_RT` e reescrito para o
contexto universal — as regras não mudaram, o alcance mudou.

Todas nasceram de defeito real. Nenhuma é precaução teórica.

---

## 0. A regra da qual as outras derivam

> **Desconfie de toda mensagem de sucesso que não verifica o resultado.**

Caso fundador: durante semanas o kit imprimiu *«Atalhos copiados para a Área de
Trabalho!»* em **todas** as máquinas, sem ter copiado nada. O caminho estava errado
e o erro do `xcopy` estava silenciado. Só apareceu quando alguém acrescentou um
contador honesto.

O padrão se repetiu tantas vezes que virou o método de trabalho do projeto:

> **rodar numa máquina real → trazer o log → corrigir o que o log expôs**

Quase todo defeito sério foi descoberto assim, não por revisão de código.

E ele quase se repetiu na criação deste fork: ao remover o batch embutido,
`Save-Preparador` passou a gravar um `.bat` de **zero byte** e devolver sucesso. Um
`.bat` vazio roda, sai com código 0 e não imprime nada — o log diria *«Executando:»*
e mais nada. A guarda em `Save-Preparador` existe por isso.

---

## 1. Nada exige administrador

Não é preferência, é o escopo. Funcionalidade que só funciona elevada **não entra** —
vira achado no diagnóstico, com o texto que o usuário leva ao chamado.

No projeto de origem, 28 funções que dependiam de admin foram **apagadas**, não
escondidas atrás de uma opção.

---

## 2. Nada que alguém esteja usando é encerrado

No original: «nada clínico», porque a máquina podia estar com uma sessão de
planejamento de tratamento aberta. Aqui vale mais largo, pelo mesmo motivo: a
máquina pode estar tratando paciente, atendendo alguém no balcão, ou rodando o lote
de processamento de um consultório.

Sete listas governam isso. Quatro decidem encerramento:

| Lista | O que faz |
|---|---|
| `$script:Protegidos` | nunca encerrado, em nenhuma circunstância |
| `$script:SessaoPreservar` | preservado na otimização de sessão |
| `$script:SessaoRemoto` | acesso remoto — encerrar derruba o próprio atendimento do TI |
| `$script:SessaoNaoCompactar` | não tem a memória compactada (engasgaria a tela) |

**`SessaoPreservar` e `SessaoNaoCompactar` terminam concatenando `SessaoRemoto`.**
Quem mexe numa mexe nas três.

### A assimetria que governa toda decisão de lista

> **Proteger algo que não precisava deixa peso morto. Não proteger algo que
> precisava interrompe atendimento. Os dois erros não custam o mesmo, e a dúvida se
> resolve protegendo.**

Uma edição descuidada no projeto de origem já marcou `sqlservr`, `VspAppTomcat`,
`MiradaRTx` e `TeamViewer` como dispensáveis — o que teria derrubado o banco de um
sistema de imagem e o acesso remoto do TI no meio de um atendimento.

### Nome de regex casa demais

As listas são expressões regulares casadas contra `ProcessName`, sem âncora por
padrão. Consequências medidas neste código:

- `ARIA` (Varian) casa com `Soarian` (Siemens). Protege por acidente.
- `^mfe` pega o McAfee, e qualquer coisa que comece com `mfe`.
- `Box` casaria com `GameBar`? Não — mas casa com `Sandboxie`, e `Sync` casa com
  meia dúzia de coisas do Windows.

> **Nome de três ou quatro letras numa lista de proteção quase sempre precisa de `^`
> e `$`.** `et`, o processo de planilha do WPS Office, é o caso extremo.

### Escreva o teste antes da entrada

Não é cerimônia. É o teste que faz aparecer a entrada que não casa — grafia
diferente, âncora faltando, nome curto demais. Lista primeiro e teste depois só
confirma o que já se acreditava.

---

## 3. Instantâneo: dado guardado nunca decide encerrar processo

O Módulo 5 consulta o Windows uma vez por operação em vez de uma vez por processo.
É o que faz o «otimizar essa sessão» caber em segundos: no projeto de origem, o
Aplicar caiu de **26,6 s para 2,6 s**, medido.

E é a origem de **seis defeitos encontrados de uma vez**, todos da mesma família.

> **Um retrato descreve o instante em que foi tirado. A decisão de encerrar acontece
> depois — e o que mudou no meio é exatamente o que importa proteger.**

### A regra, em três partes

**1. Quem decide lê ao vivo.** O caminho vem de `Get-CaminhoAoVivo`, que lê `.Path`
do processo na hora e só usa o instantâneo como reforço, para processo cujo `.Path`
não abre. Nunca o contrário.

**2. Quem guarda declara o prazo.** `$script:SnapMaxSeg = 15` segundos. Vencido,
reconstrói. A guarda de leitura confere o **conteúdo**, não só a existência:
instantâneo sem nenhum processo vivo não é resposta, e nenhum dos dois lados confia
em que o outro se comportou.

**3. Consulta que falhou não se guarda.** Cachear o vazio desligaria a proteção pela
operação inteira, em silêncio. E atenção ao detalhe da linguagem: **um
`pscustomobject` vazio é verdadeiro em PowerShell**, então `if ($cache)` não
distingue «tenho resposta» de «tenho um objeto vazio».

### O que pode ser guardado

O retrato **da lista que se mostra na tela** — `Get-ProcessosSessao` e
`Get-GruposSessao`. Eles montam o plano e o relatório; o executor redecide cada PID
ao vivo, um por um. Essa separação é o que torna o cache seguro.

> **Se algum dia alguém usar essa lista para decidir um encerramento, o prazo dela
> passa a importar e o comentário no código está errado.**

### Lista de proteção: na dúvida, protege

Havia um filtro que descartava PID listado na sinalização se ele não estivesse no
instantâneo. A intenção era não proteger PID morto. O efeito era descartar PID
**vivo** porque o retrato envelheceu.

> **Proteger um PID que já morreu não custa nada — não há o que encerrar.**

### As duas leituras da sinalização

| leitura | custo medido | traz | não traz |
|---|---|---|---|
| `-Rapido` | 9 ms | PID que o app acabou de listar | descendência |
| completo | 278 ms | o worker nascido de um PID listado | o que entrou na lista depois |

`Get-EmUsoOperacao` **une** as duas. O código anterior escolhia uma **ou** outra por
`if ($Protegidos.Count -eq 0)` — que trocava uma pela outra exatamente quando o
arquivo estava vazio, isto é, quando não havia nada a proteger.

---

## 4. Diagnóstico não pode emudecer

`Get-Service -ErrorAction Stop` tropeça no primeiro serviço que o usuário não pode
consultar. O tropeço vira terminante, o `catch` engole, e a lista sai **vazia**.

E lista vazia ali não parece erro: parece *«esta máquina não hospeda serviço
nenhum»*. O aviso de papel de servidor — o que impede otimizar uma estação que serve
outras pessoas — desaparecia calado.

> **Verificação que não aconteceu tem de dizer que não aconteceu.** Silêncio é lido
> como «está tudo bem», e é a forma mais barata de mentir sem querer.

---

## 5. Texto que o Windows devolve é localizado

Uma versão filtrava `sc query` pela palavra `STATE`. O Windows em português devolve
`ESTADO`. Resultado: **zero agentes de segurança** relatados em toda estação pt-BR,
sem erro nenhum.

> **Cmdlet que devolve OBJETO é seguro. Comando que devolve TEXTO é localizado, e
> comparar esse texto é um defeito que só aparece na máquina de outra pessoa.**

Num projeto universal isso deixa de ser um caso e passa a ser uma categoria. Prefira
`Get-Service` a `sc query`, `Get-CimInstance` a `wmic`,
`[Environment]::GetFolderPath()` a nome de pasta escrito à mão. Quando não houver
alternativa, compare pela parte que **não** se traduz.

O mesmo vale para formato regional: data, separador decimal e separador de milhar.
Use `[Globalization.CultureInfo]::InvariantCulture` em toda conversão que precise
ser estável.

---

## 6. Resolver por marca, não por caminho

> **Código que localiza pasta por caminho fixo quebra na primeira reorganização.
> Código que a localiza por marca sobrevive.**

Marca é uma condição verificável no destino — um componente de caminho, um arquivo
que só existe ali. A comparação é por **componente exato**, nunca por prefixo:

```powershell
return ($Caminho.Split([char]92, [char]47) -contains $script:MarcaEcossistema)
```

O `Split` por `[char]92` em vez de uma classe de caractere escrita é deliberado. A
versão anterior usava `[\\/]` e uma camada de ferramenta comeu uma barra,
transformando em `[/]`: passou a casar **só** com barra normal, e a função devolvia
falso para **todo** caminho do Windows.

> **O parse passou, a regex compilou, e a proteção simplesmente não aconteceu.** Foi
> um teste NEGATIVO que pegou. Teste positivo não vê proteção que parou de proteger.

E daí a guarda que a suíte ganhou: **alternativa vazia num regex de alternância casa
com qualquer coisa.** `'Spotify' -match 'a||b'` devolve `True`. Numa lista de
proteção isso vira «protege tudo», os testes negativos passam todos, e nada quebra —
o modo de falha mais difícil de notar que existe neste código.

---

## 7. Não se contorna antivírus

O Cortex XDR bloqueou o app numa estação como *«Suspicious script detected»*. A
decisão foi **não rodar naquela máquina**.

Isso não impede melhorar o código por motivo legítimo — a compilação em tempo de
execução foi removida porque era **desnecessária**, e o método alternativo já
funcionava em todas as máquinas. Mas a diferença importa: remover código
desnecessário é higiene; ofuscar para escapar de detecção é outra coisa.

Se algum dia fizer sentido rodar lá, o caminho é o TI liberar por hash depois de
olhar o código.

---

## 8. Nada de dado de paciente, credencial ou identificador

**Não existe neste repositório, e não pode passar a existir.** A pasta sincroniza
com o Google Drive. `.gitignore` protege o git — **não protege o Drive**.

Este projeto nasceu de uma varredura que barrou 47 ocorrências de identificador de
infraestrutura de um hospital, e foi criado **sem histórico compartilhado** porque
histórico não se desfaz depois. Ver `PROVENIENCIA.md`.

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Verificar-Vazamento.ps1
```

Ele confere por **categoria** — a forma de um identificador de rede —, nunca por
lista de nomes. Uma lista de nomes num arquivo versionado **põe no repositório
exatamente o que ela existe para manter fora**, e foi assim que a primeira versão
desta seção nasceu errada.

---

## 9. Ao alterar qualquer lista, testar contra processos reais

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Sessao.ps1
```

A suíte confere o parse, o BOM, **a integridade das sete listas**, e exercita as
proteções contra uma árvore de processos de verdade que ela mesma cria e derruba.

O bloco 11 imprime o que **ainda não** está coberto. Não conta como falha, de
propósito: uma suíte que falha por trabalho pendente é ignorada em uma semana.
