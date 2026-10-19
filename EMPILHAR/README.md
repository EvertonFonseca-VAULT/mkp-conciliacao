# EMPILHAR — guardar o lançamento na planilha histórica

Esta é a **etapa 3** do dia, depois do `PARA_PROCESSAR`. "Empilhar" significa guardar, numa planilha histórica grande (uma por canal, no `X:`), a linha de cada venda já lançada — pra ter um histórico consultável e auditável, sem precisar abrir o TOTVS toda vez.

O programa não mexe no TOTVS: ele pega o "Para Processar" que a etapa anterior já gerou (que por sua vez já conferiu contra o TOTVS/SRI) e cruza com o AB do dia pra montar a linha do empilhado.

## Antes de rodar

Rode `PARA_PROCESSAR` primeiro, para o canal e o dia que você quer empilhar — o empilhador depende do arquivo "Para Processar" já existir.

## Como rodar

Abra um terminal dentro da pasta `EMPILHAR`:

```
..\.venv\Scripts\python.exe empilhar_canais.py FULL 22.09.2026
```

Isso **monta e confere**, mas **não grava nada** — é o modo seguro de conferir antes de publicar de verdade.

Para gravar de fato na planilha histórica do `X:`:

```
..\.venv\Scripts\python.exe empilhar_canais.py FULL 22.09.2026 23.09.2026 --publicar
```

(dá pra listar vários dias na mesma chamada, como no exemplo acima).

Troque `FULL` pelo canal desejado: `ES`, `MOVEIS`, `ELETRO`, `CENTER`, `COLETAS_CO05`, `PRIME` ou `WEB`.

Rodar de novo um dia que já foi empilhado **não duplica** a linha — pode rodar quantas vezes precisar sem medo.

## Erros comuns

| Mensagem / sintoma | O que fazer |
|---|---|
| Não acha SOURCE_ID / avisa "0 ou 2 candidatos" | O cruzamento entre o AB e o Para Processar ficou ambíguo para aquele pedido — precisa de conferência manual antes de publicar. |
| Diz que o "Para Processar" não existe | Rode a etapa `PARA_PROCESSAR` primeiro, para aquele canal e dia. |
| Demorou mas não travou | Normal — o programa lê o arquivo histórico inteiro antes de gravar; em arquivos grandes isso leva uns segundos. |

## Testar sem mexer em nada real

```
..\.venv\Scripts\python.exe empilhar_canais.py CANAL --teste    # testa só um canal (troque CANAL)
..\.venv\Scripts\python.exe empilhar_canais.py --teste          # testa todos os canais de uma vez
```

Não acessa o `X:` nem grava nada de verdade — só confirma que a lógica de montar/cruzar continua certa. Use depois de qualquer atualização do projeto.
