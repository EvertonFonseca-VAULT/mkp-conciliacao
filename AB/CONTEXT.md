# Mkp/AB — Conciliação de Vendas

Concilia dados de venda dos canais contra os registros internos (TOTVS/SRI) e empilha o resultado no ledger de acompanhamento. Além disso, empilha os repasses fechados de marketplaces que não passam por esse cruzamento diário (ver Repasse fechado).

## Language

**Canal**:
Escopo: só dentro do Mercado. Uma sub-divisão de venda do Mercado com formato de dados e regras de reconciliação próprias, cruzada diariamente contra TOTVS/SRI via AB. Ativos hoje: CENTER, COLETAS05, MOVEIS, PRIME, WEB, FULL, ES — a lista cresce conforme a empresa entra em novos canais do Mercado.
_Avoid_: Plataforma, marketplace, loja (quando usados como sinônimo de Canal neste contexto).

**Marketplace**:
A plataforma de venda em si (Mercado, LIVELO, e as que vierem). Vocabulário do negócio (mesmo termo usado em `Controle_Conciliacao_Recebimentos.xlsx`, coluna MARKETPLACE).

**Loja**:
Sub-divisão de venda de um Marketplace (coluna LOJA do `Controle_Conciliacao_Recebimentos.xlsx`). CENTER e WEB existem tanto como Loja do LIVELO quanto como Canal do Mercado — são entidades diferentes que coincidem de nome; não confundir.

**Repasse fechado**:
Família de Marketplaces (LIVELO, COLOMBO, e os que vierem) sem cruzamento diário contra TOTVS/AB — o arquivo que o marketplace manda já é o dado final, seja ele mensal com resumo embutido (LIVELO) ou uma exportação diária sem resumo, com a data do repasse vindo do nome do arquivo (COLOMBO). Cadência (mensal/diária) e resumo-ou-não são configuração por chave, não um eixo que separa motores. Motor em `empilhar_repasses.py`, registro `REPASSES` (chave = Marketplace_Loja, ex. `LIVELO_CENTER`, `COLOMBO_CENTER`). Documentado em ADR-0001.
_Avoid_: tratar como Canal, ou fundir o motor com `empilhar_canais.py` — são seams diferentes (ver ADR-0001). Também evitar: criar um motor novo por marketplace quando a diferença é só configuração (de onde vem a data, se tem resumo) — isso cabe em `REPASSES`.
