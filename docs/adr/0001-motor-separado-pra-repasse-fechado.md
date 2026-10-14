# ADR-0001: Motor separado pra marketplaces de repasse fechado (não fundir com empilhar_canais.py)

## Status
Aceito (08/10/2026)

## Contexto
`empilhar_canais.py` resolve o empilhado dos Canais do Mercado: cruzamento diário do AB contra o
Para Processar do TOTVS pra achar o número TOTVS de cada linha. LIVELO (primeiro marketplace da
família "repasse fechado") não tem esse passo — o arquivo mensal já vem fechado pelo próprio
marketplace, com o resumo de IR/Comissão/Repasse embutido.

Review de arquitetura (08/10/2026) cogitou colocar os dois debaixo de um `grupo="repasse_fechado"`
dentro do `CANAIS` existente.

## Decisão
Dois módulos profundos separados: `empilhar_canais.py` (Mercado, cruzamento diário) e
`empilhar_repasses.py` (repasse fechado, mensal, sem cruzamento). Cada um com seu próprio registro
(`CANAIS` / `REPASSES`).

## Por quê
Os dois não compartilham seam de verdade. `empilhar_canais.py` tem uma etapa inteira (resolver
número TOTVS via AB×Para Processar) que não existe no outro lado. Forçar os dois num `grupo` só
criaria um módulo raso, cheio de `if` pra pular o que não existe de um lado. Teste da deleção: se
a unificação fosse apagada, a complexidade não reaparece concentrada — ela só revela que nunca
foram a mesma coisa.

## Consequência
Próximo marketplace de repasse fechado (2ª entrada em `REPASSES`) entra sem tocar em
`empilhar_canais.py`. Se um marketplace futuro precisar de cruzamento diário tipo TOTVS/AB, ele é
candidato a `CANAIS`, não a `REPASSES` — a pergunta que decide onde ele entra é "tem cruzamento
diário contra um sistema interno, ou o arquivo mensal já vem fechado?".
