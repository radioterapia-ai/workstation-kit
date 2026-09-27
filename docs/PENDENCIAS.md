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
