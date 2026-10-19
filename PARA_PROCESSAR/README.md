# PARA_PROCESSAR — montar o arquivo pra lançar no TOTVS

Esta é a **etapa 2** do dia, depois do `KONCILI`. O programa pega o AB baixado (ou um arquivo de teste), confere pedido a pedido contra o TOTVS (e contra o SRI, quando o TOTVS ainda não tem o lançamento), e monta a planilha **"Para Processar"** — o arquivo que alguém do financeiro vai de fato lançar no TOTVS.

## Canais que esta pasta sabe processar

| Canal | Tem conferência de frete? | Observação |
|---|---|---|
| FULL | sim | usa 2 vendedores no arquivo de origem (PTI) |
| ES (Coletas ES) | sim | |
| MOVEIS | não | |
| ELETRO (Eletrodoméstico) | não | não usa PTI |
| CENTER | não | |
| COLETAS CO05 | sim | |
| PRIME | não | |
| WEB | não | |

As subpastas `ES/`, `FULL/`, `MOVEIS/` e `WEB/` guardam o motor de cada canal (o código que sabe ler o formato específico de cada um) — no dia a dia você não precisa entrar nelas, só chamar `rodar_canal.py` daqui de cima.

## Antes de rodar pela primeira vez

- Rode a etapa `KONCILI` primeiro, para o dia que você quer processar (ou use `--ab` com um arquivo de teste, veja abaixo).
- Tenha acesso à unidade `X:` e ao TOTVS (o cruzamento usa o TOTVS como fonte principal e o SRI como reforço).

## Como rodar

Abra um terminal dentro da pasta `PARA_PROCESSAR`:

```
..\.venv\Scripts\python.exe rodar_canal.py FULL 22.09.2026
```

Processa **um canal, um dia**. Troque `FULL` por `ES`, `MOVEIS`, `ELETRO`, `CENTER`, `COLETAS_CO05`, `PRIME` ou `WEB`.

Processar **todos os canais do dia** de uma vez:

```
..\.venv\Scripts\python.exe rodar_canal.py TODOS 22.09.2026
```

Processar **vários dias** do mesmo canal (lista cada data, um por um):

```
..\.venv\Scripts\python.exe rodar_canal.py FULL 11.09.2026 14.09.2026 18.09.2026
```

Testar com um arquivo específico, **sem mexer no AB real** (só funciona com 1 canal + 1 dia):

```
..\.venv\Scripts\python.exe rodar_canal.py WEB 22.09.2026 --ab "X:\caminho\MERCADO WEB_teste 22.09.2026.xlsx"
```

O programa **nunca sobrescreve** um arquivo "Para Processar" que já existe, a não ser que você mande explicitamente com `--forcar`:

```
..\.venv\Scripts\python.exe rodar_canal.py ELETRO 22.09.2026 --forcar
```

## O que olhar no resultado

Ao final, a tela mostra quantas linhas foram processadas, o total (H:M) e quantas ficaram como **"Verificar"** — pedidos que não bateram automaticamente contra o TOTVS/SRI e precisam de olhar manual antes de seguir para o TOTVS.

## Erros comuns

| Mensagem / sintoma | O que fazer |
|---|---|
| Avisa que o AB do canal não está no `X:` | Rode a etapa `KONCILI` primeiro com `--salvar` pra aquele canal/dia, ou use `--ab` com um arquivo manual. |
| Já existe um "Para Processar" para aquele dia | É proposital — para não sobrescrever por engano. Use `--forcar` só se tiver certeza. |
| Muitos pedidos em "Verificar" | Pode ser dia com problema real no TOTVS/SRI, ou o calendário desse dia está errado — confira com o resultado do `KONCILI` antes de supor que o processo falhou. |

## Testar sem mexer em nada real

```
..\.venv\Scripts\python.exe rodar_canal.py --teste
```

Roda só a lógica interna (sem tocar em arquivos reais nem no TOTVS). Use depois de qualquer atualização do projeto, antes de processar pra valer.
