# ADR-0002: BARRA CRED e KABUM não entram no registro `REPASSES` (por enquanto)

## Status
Aceito (14/10/2026)

## Contexto
Revisão de arquitetura (14/10/2026) apontou `empilhar_barra_cred.py` e `empilhar_kabum.py` como
candidatos a migrar pro registro `REPASSES` de `empilhar_repasses.py` — ambos são, como LIVELO e
COLOMBO, marketplaces de repasse fechado (sem cruzamento diário contra TOTVS/AB), e os dois
docstrings se autodeclaram `"gohorse -- regra provisoria, estruturar depois"` (13/10/2026, um dia
antes desta revisão).

Investigação mostrou que a semelhança é só de família (repasse fechado), não de mecânica:

- **Descoberta/dedupe**: `empilhar_repasses.py` varre `pasta_fonte/<ano>/<mes>.<ano>/`, resolve
  **1 data por arquivo** (`resolver_periodo`) e decide o que falta comparando **datas** contra a
  aba Calendário (`listar_periodos`/`datas_dados`/`datas_cal`). BARRA CRED pega só o último
  ano/mês/arquivo da pasta e dedupa por **chave natural de cada linha** (Pedido/TID + Parcela +
  Valor) contra todas as linhas já gravadas no destino — não existe "data pendente" nem aba
  Calendário nesse arquivo. KABUM varre a árvore inteira (todo arquivo `DD.MM.AAAA KABUM*.xlsx`)
  e também dedupa por linha (TID+Parcela+Valor), mas escreve a aba `EMPILHADO` agrupando as linhas
  novas por **dia encontrado nelas**, não por data pendente de um calendário pré-existente.
- **Escrita de célula**: `empilhar_repasses.py._escrever` só copia `number_format`. BARRA CRED e
  KABUM copiam `number_format` + `font` + `alignment` + `fill` + `border` da última linha
  autêntica — perdê-los mudaria a aparência visual das planilhas (fonte/cor/borda ficariam no
  padrão do openpyxl, não as da planilha mantida à mão).

## Decisão
BARRA CRED e KABUM continuam em arquivos próprios, fora de `REPASSES`. Unificação completa
exigiria o motor compartilhado suportar 2 estratégias de descoberta/dedupe (período-por-arquivo
contra calendário vs linha-individual contra chave natural) — isso tocaria `main()`/
`listar_periodos()`, hoje usado em produção por `LIVELO_CENTER`, `LIVELO_LG` e `COLOMBO_CENTER`,
sem ambiente real disponível nesta sessão pra validar contra dado de produção.

O único problema real e sem ambiguidade encontrado — zero autoteste nos dois arquivos — foi
corrigido nesta mesma revisão: `empilhar_barra_cred.py --teste` e `empilhar_kabum.py --teste`
cobrem a mecânica pura (montagem de linha, chave de dedupe, conversão de data) sem precisar abrir
nenhum `.xlsx`.

## Por quê
Teste da deleção: se a unificação fosse forçada mesmo assim, a complexidade não desaparece — ela
só migra pra dentro do motor compartilhado como um `if estrategia == "linha"` por todo lugar que
hoje assume período-por-arquivo, tornando `empilhar_repasses.py` mais raso (mais `if`, por
canal/estratégia) em vez de mais profundo, e arriscando regressão silenciosa nos 3 canais que já
rodam em produção. "Gohorse, estruturar depois" é dívida deliberada e datada (1 dia antes desta
revisão) — não dado como esquecimento a corrigir às pressas sem quem mantém os dois lados por
perto pra revisar o diff.

## Consequência
Se um 3º marketplace de repasse fechado aparecer com a MESMA mecânica de BARRA CRED/KABUM (dedupe
por linha, não por período), aí sim há 2 adaptadores reais pro mesmo seam — momento de extrair um
motor "repasse fechado por linha" separado de `empilhar_repasses.py` (que continua sendo o motor
"repasse fechado por período"), não de forçar os 2 debaixo do mesmo `REPASSES`.
