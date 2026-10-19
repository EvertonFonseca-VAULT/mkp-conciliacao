# DOCS — decisões e vocabulário do projeto

Esta pasta não tem programas pra rodar. Ela guarda anotações sobre **por que** o projeto foi construído de um certo jeito — útil pra quem for entender ou alterar o código no futuro (pessoa ou assistente de IA).

- `ADR/` — "Architecture Decision Records": um arquivo por decisão técnica importante, explicando o problema, a decisão tomada e o porquê. Por exemplo, `0001-motor-separado-pra-repasse-fechado.md` explica por que o empilhamento de canais do Mercado e o de marketplaces como a LIVELO são dois programas separados, e não um só.
- `AGENTS/` — instruções de como um assistente de IA (como o Claude) deve explorar e usar a documentação deste projeto. `domain.md` explica onde achar o vocabulário do negócio (arquivo `CONTEXT.md`, na raiz do projeto).

Se você não mexe em código, pode ignorar esta pasta — ela não afeta o dia a dia de rodar `KONCILI`, `PARA_PROCESSAR`, `EMPILHAR` ou `CONCILIACAO`.

## Vocabulário do negócio

O glossário dos termos usados neste projeto (Canal, Marketplace, Loja, Repasse fechado, etc.) está em `CONTEXT.md`, na raiz da pasta `MARKETPLACE`.
