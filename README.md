# Workstation Kit

Prepara, diagnostica e otimiza estações de trabalho Windows. **Roda sem privilégio
de administrador**, em máquina corporativa gerenciada.

Serve qualquer computador — de clínica, de consultório, de recepção ou pessoal —, em
qualquer país.

> **É um app de Windows, e é para rodar logo após logar na máquina** — antes de
> abrir o trabalho do dia. Rodar no meio do dia pode encerrar trabalho em curso; o
> app mede o momento e avisa quando não é a hora.
>
> Ele conhece software de radioterapia de vários fabricantes porque foi ali que
> nasceu, e essa proteção continua valendo. Mas não é um app de radioterapia: é o
> que prepara a máquina antes de ela ser usada, para qualquer trabalho.

Versão atual: **1.0** · PowerShell 5.1 + WinForms

> **Estado: fork recém-aberto.** O aplicativo roda e a suíte de teste passa, mas a
> camada de idioma ainda não existe e 23 entradas de lista de proteção estão por
> cobrir. Ver `docs/PENDENCIAS.md`, que separa o que foi verificado do que ainda
> não foi.

---

## Distribuição

Dois arquivos, na mesma pasta:

```
Start_WorkstationKit.cmd     o atalho que abre o app
WorkstationKit.ps1           o aplicativo inteiro
```

O `.ps1` é **um arquivo só**, e é deliberado: a distribuição é cópia de arquivo, não
instalação. No projeto de origem, atualizar a cópia na pasta da rede era o que
distribuía a correção para 11 máquinas sem ninguém passar de mesa em mesa.

---

## Os módulos

| | |
|---|---|
| **Diagnóstico** | retrato da máquina, e o que dá para resolver |
| **Arquivos grandes** | somente leitura; lista e dá veredicto |
| **Limpeza** | temporários, caches e Downloads antigos |
| **Otimizar sessão** | encerra o que não está em uso, ajusta prioridade, compacta memória |
| **Persistência** | o que volta sozinho a cada logon |

Nomes em vez de números: o módulo de preparação de ambiente saiu, e renumerar o resto
quebraria toda referência.

Antes de aplicar, a Limpeza e o Otimizar sessão dizem se **é a hora certa**: quanto
tempo a sessão está aberta, se há carga em andamento, e se o trabalho do dia já
começou.

## Nada é definitivo, exceto os arquivos que você apagar

| o que | volta por |
|---|---|
| processos encerrados, tarefas paradas, prioridade de CPU, memória compactada | o próximo logon |
| itens tirados da inicialização, ajustes de desempenho | o botão **Desfazer** |
| Downloads antigos | a Lixeira |
| temporários e caches | não voltam — é a limpeza |

A Lixeira é esvaziada **primeiro**, e o que sai de Downloads vai para ela depois; por
isso continua recuperável.

**Não desinstala nada.** Quando um programa parece dispensável, o app **encerra o
processo** em vez de desinstalar: libera a memória agora, e no próximo logon está tudo
de volta. Detecção que exigiria ação definitiva — mapeamento de rede morto, credencial
órfã — sai como achado, com o comando, para você decidir.

---

## O que ele não faz

- não exige administrador, em nenhuma função
- não desinstala nada e não altera `HKLM`
- não limpa registro: não há evidência publicada de ganho mensurável, e há risco real
- não baixa nada e não se atualiza sozinho
- não apaga arquivo fora de pasta de cache do próprio perfil — o Módulo 4 lista e dá
  veredicto, quem apaga é você, pelo Explorer
- não encerra o que alguém está usando: Citrix, prontuário, banco de imagem,
  planejamento e o acesso remoto do TI estão em lista de proteção, e o Módulo 5
  confere cada processo **ao vivo** antes de encerrar, um por um
- não contorna antivírus, e não compila código em tempo de execução

---

## Idioma

| Camada | Idioma |
|---|---|
| texto que o usuário lê | inglês, português e espanhol |
| código, comentário e documentação | português do Brasil |

A segunda linha é a regra do ecossistema em que este projeto vive e não muda por o
produto ser internacional. Ver `CLAUDE.md`.

---

## Origem

Fork do `WORKSTATION_RT`, aberto em 27/09/2026 a partir do commit `ad13e33`. Aquele
projeto continua existindo e continua sendo a ferramenta da clínica do autor.

O repositório foi criado **sem histórico compartilhado**, de propósito: o `.ps1` de
origem tinha 47 ocorrências de identificador de infraestrutura de um hospital real —
domínio, fileserver, dois IPs internos, cinco hostnames. Um clone gravaria isso na
história deste projeto para sempre, e história não se desfaz depois.

`docs/PROVENIENCIA.md` tem a lista do que saiu, o que ficou no lugar, e o defeito que
a própria remoção quase criou.

---

## Antes de entregar qualquer mudança

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Sessao.ps1
```

36 casos, e cada bloco nomeia o defeito real que ele pega. São **testes negativos**:
montam o caso em que a proteção precisa falhar se estiver quebrada.

> Teste positivo não vê proteção que parou de proteger.

---

## Documentação

| Arquivo | Para quê |
|---|---|
| `CLAUDE.md` | como trabalhar neste código |
| `docs/REGRAS-DE-SEGURANCA.md` | o que nunca se toca e por quê |
| `docs/PROVENIENCIA.md` | de onde veio e o que saiu |
| `docs/PENDENCIAS.md` | o trabalho da versão universal, em ordem |

---

Toda saída deste aplicativo é um rascunho sujeito a revisão.
