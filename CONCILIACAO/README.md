# CONCILIACAO — conferir com o banco e ver o painel do dia

Esta é a **etapa 4**, a última do dia. Ela cruza o que já foi lançado (RB — relatório do banco) com o repasse do Mercado, calcula a diferença (`DIF = REPASSE − TOTAL_RB`) e mostra um painel (dashboard) em HTML com o que fechou sozinho e o que precisa de olhar manual.

## O jeito mais simples de rodar (recomendado para quem não mexe com código)

1. Abra a pasta `CONCILIACAO` no Explorador de Arquivos.
2. Dê **duplo clique** em `iniciar_conciliacao.bat`.
3. Vai abrir uma janela preta com um menu numerado (FULL, COLETAS ES, MOVEIS, ELETRODOMESTICO, CENTER, COLETAS CO05, PRIME, WEB). Digite o número do canal e aperte Enter.
4. Escolha `1` para um único dia, ou `2` para um intervalo de datas, e digite a(s) data(s) no formato `dd.mm`.
5. No fim aparece um **RESUMO**: canal, data, total (H:M), quantos pedidos ficaram em "Verificar" e se bateu com o calendário.
6. Um arquivo HTML é gerado (o dashboard) — abra ele no navegador pra ver os detalhes, pedido por pedido.

## Antes de rodar, pela primeira vez

A conciliação exige três arquivos de entrada, todos com a **data do dia que você está processando** (campo "Dt.Ref"):

- `AUXILIAR\rb.xlsx`
- `AUXILIAR\rbx.xlsx`
- `C:\TEMP\Conciliador - Backup\CONCILIADOR_RB.xlsx`

Se a data desses arquivos não bater com o dia pedido, o programa **recusa rodar** — isso é proposital, para não gerar um resultado errado silenciosamente. Exporte esses relatórios do dia certo antes de tentar de novo.

## Rodando por linha de comando (alternativa ao duplo clique)

```
..\.venv\Scripts\python.exe interativo.py FULL 22.09
..\.venv\Scripts\python.exe interativo.py FULL --data-inicio 22.09 --data-fim 30.09
```

## Como ler o painel (dashboard)

- Cada pedido que **não** bateu automaticamente cai na fila **VERIFICAR**, na ordem Frete → Comissão → Banco — nada some do radar, tudo que não fechou sozinho fica visível pra conferência manual.
- Clicar num valor da fila (Frete, Comissão, Banco) **copia** o valor, já formatado, pronto pra colar no TOTVS.
- As marcações que você fizer no painel (ajustado, natureza) ficam guardadas no próprio navegador, por canal e por dia — não precisa refazer se fechar e abrir de novo a mesma página, mas se trocar de computador ou limpar o navegador, elas se perdem.
- O painel funciona **offline**: é um único arquivo HTML, não precisa de internet pra abrir.

## A fila VERIFICAR some com um clique ("buscar no ML/MP")

Alguns pedidos da fila VERIFICAR podem ser buscados automaticamente no site do Mercado Livre/Mercado Pago. Para isso:

1. Rode `abrir_chrome_debug.bat` (duplo clique) — ele abre um Chrome especial, conectado ao programa.
2. Faça login normalmente no Mercado Livre e no Mercado Pago nesse Chrome, se pedir.
3. Rode a conciliação com `--scraper` (ou marque a opção equivalente no assistente), e deixe essa janela do Chrome aberta — o programa usa ela pra "olhar" os pedidos sozinho.

## Erros comuns

| Mensagem / sintoma | O que fazer |
|---|---|
| Recusa rodar, fala de `Dt.Ref` | Os arquivos `rb.xlsx`, `rbx.xlsx` ou `CONCILIADOR_RB.xlsx` não são do dia pedido — exporte o relatório certo. |
| Muitos pedidos em VERIFICAR, mais do que o normal | Confira se o calendário daquele dia está certo (às vezes o alvo em si está errado) antes de supor que o processo falhou. |
| O duplo clique no `.bat` fecha sozinho / não encontra `python.exe` | Confira se a pasta `MARKETPLACE\.venv` existe (veja o Setup no `README.md` da raiz do projeto) — o `.bat` usa o Python de dentro dela. |
| Scraper não encontra o pedido | Confirme que o Chrome aberto pelo `abrir_chrome_debug.bat` está logado no Mercado Livre e no Mercado Pago, e que ele continua aberto durante o processamento. |

## Testar sem mexer em nada real

```
..\.venv\Scripts\python.exe dashboard.py
```

Gera o dashboard com dados de exemplo, sem tocar em arquivo nenhum do dia a dia. Deve terminar com `dashboard.py: ok`.
