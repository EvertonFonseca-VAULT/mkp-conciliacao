# Conciliação de Marketplace (Mercado Livre / Mercado Pago)

Automatiza o ciclo diário de conciliação financeira do Mercado: baixar o AB no Koncili, gerar o **Para Processar** para o TOTVS, empilhar as bases e conciliar o repasse contra o extrato bancário num dashboard de baixas.

Antes era feito com macros, planilhas e Power Query manuais (`AB_MERCADO_TOTVS.xlsm`, `CONTROLE_V2.xlsx`, conferência pedido a pedido no ML).

## Fluxo do dia

```
Koncili ──► AB (X:)  ──► Para Processar (C:\RODAR TOTVS) ──► TOTVS
                │                                         
                ├──► Empilhado FULL / ES (X:, com backup)
                └──► Conciliação (AB/RB) ──► dashboard HTML ──► baixas (Frete, Comissão, Banco)
```

Regra de validação em todos os canais: **H:M do Para Processar = RESULTADO do Koncili = VALOR_EB do calendário**.

## Canais

| Canal | Conta no Koncili | Vendedor | PTI | Frete | Saída do Para Processar |
|---|---|---|---|---|---|
| FULL | ver `CANAIS_KONCILI` | 751 (PTI tem 750 + 751) | `C:\TEMP\PTI FULL*.csv` | sim | `C:\RODAR TOTVS\PROCESSAR\FULL` |
| ES | ver `CANAIS_KONCILI` | 763 | `C:\TEMP\PTI ES.csv` | sim | `C:\RODAR TOTVS\PROCESSAR\COLETAS ES` |
| MOVEIS | `WEBCO MÓVEIS` | 755 | `C:\TEMP\PTI MOVEIS.csv` | não | `C:\RODAR TOTVS\PROCESSAR\MOVEIS` |
| ELETRO | `WEBCO ELETRO` | 756 | não usa | não | `C:\RODAR TOTVS\ELETRODOMESTICO` |
| CENTER | `CONTINENTALCENTER` | 944 | `C:\TEMP\PTI CENTER.csv` | não | `C:\RODAR TOTVS\PROCESSAR\CENTER` |
| COLETAS CO05 | `LTWE20240219130512` | 847 | `C:\TEMP\PTI COLETAS 05.csv` | sim | `C:\RODAR TOTVS\PROCESSAR\COLETAS CO05` |
| PRIME | `WEBCOPRIME` | 657 | `C:\TEMP\PTI_PRIME.csv` | não | `C:\RODAR TOTVS\PROCESSAR\PRIME` |
| WEB | `WEB CONTINENTAL` | 578 | `C:\TEMP\PTI_WEB.csv` | — | só baixa o AB (implantação, sem Para Processar automatizado ainda) |

Nos códigos de vendedor, o PTI usa `000` na frente (`000755`); CMM/MOV usam sem zeros.

## Estrutura

```
Mkp/
├── requirements.txt
└── AB/
    ├── koncili.py            # login no Koncili, acha o repasse, lê o RESULTADO, baixa o AB e confere com o calendário
    ├── rodar_canal.py        # Para Processar de FULL, ES, MOVEIS, ELETRO, CENTER, COLETAS CO05 e PRIME (ou TODOS) num comando
    ├── empilhar.py           # empilhado FULL/ES: ensaio + gravação (com backup em C:\BACKUP EMPILHADO)
    ├── empilhar_gravar.py    # grava as linhas novas direto no XML do xlsx, sem duplicar dias
    ├── pipeline.py           # execução por lote do FULL com status em SQLite (só FULL)
    ├── FULL/                 # cruzamento AB × TOTVS/SRI + PTI (2 vendedores) e layout de 40 colunas
    ├── ES/                   # layout de 60 colunas (também lê 43 em português e 35 do WEB/CENTER)
    ├── MOVEIS/               # modelo_ab.py: validação, pagamento, frete e Para Processar (compartilhado)
    ├── WEB/                  # exemplos de Para Processar do WEB (implantação)
    └── RB/                   # conciliação com o RB, classificação e dashboard
```

Layouts do AB bruto: FULL 40 colunas · ES/CO05 60 · CENTER/WEB 35 · MOVEIS/ELETRO 43 ou 48 (português).

## Uso diário

```bash
cd AB

# 1. Baixa o AB do Koncili e confere com o calendário
python koncili.py ES 22.09.2026                  # só confere
python koncili.py ELETRO 22.09.2026 --salvar     # grava no X: (não sobrescreve)

# 2. Para Processar (não sobrescreve sem --forcar)
python rodar_canal.py FULL 22.09.2026                       # 1 canal, 1 dia único
python rodar_canal.py TODOS 22.09.2026                      # todos os canais, 1 dia
python rodar_canal.py FULL 11.09.2026 14.09.2026 18.09.2026 # 1 canal, vários dias — lista cada data (não é intervalo)

# 3. Empilhado FULL / ES
python empilhar.py FULL 22.09.2026                       # ensaio local
python empilhar.py FULL 22.09.2026 --publicar            # troca o do X: (com backup)

# 4. Conciliação + dashboard (time: duplo clique em RB/iniciar_conciliacao.bat)
cd RB && python interativo.py
```

Variáveis de ambiente para o Koncili: `KONCILI_USUARIO` e `KONCILI_SENHA`. Testes offline: `python koncili.py --teste`, `python rodar_canal.py --teste`, `python empilhar.py --teste`, `python dashboard.py`.

## Conciliação e dashboard (`AB/RB`)

- `DIF = REPASSE − TOTAL_RB`; contas **Original (103925)** e **Transitória (X103925)**.
- STATUS, nesta ordem: REBATE → SEM DIFF → EST COM → MEDIACAO → VERIFICAR.
- Os VERIFICAR são buscados no ML/MP pelo scraper (`scraper_ml.py`, precisa do Chrome com debug remoto via `abrir_chrome_debug.bat`) e recebem ação sugerida e histórico de baixa.
- Dashboard: HTML de um arquivo só, offline. A fila VERIFICAR segue a ordem das baixas: **Frete, Comissão, Banco** (Banco = RBX + DIF), com clique para copiar o valor. Naturezas: Frete 2209, Tarifa 2804, Rebate 1107.
- Estado do analista (ajustados, marcações) fica no `localStorage` do navegador, por canal e dia.

## Onde mexer (mudança → arquivo)

| Quero mudar… | Abra |
|---|---|
| Pasta, prefixo do AB, PTI, frete ou nome de saída de um canal do Para Processar | `AB/rodar_canal.py` → `CANAIS_PP` |
| Conta do Koncili, pasta ou layout do AB baixado | `AB/koncili.py` → `CANAIS_KONCILI` |
| Regra de cruzamento com o TOTVS/SRI (chave SHIPPING_ID) | `AB/FULL/processar_full.py` |
| Layout de colunas do AB (35, 40, 43, 48 ou 60) | `AB/ES/processar_es.py` (`carregar_ab`), `AB/ES/tipagem_es.py`, `AB/FULL/tipagem_full.py` |
| Categorias do AB (Deduzir Cupom, Taxa financiamento…) e fórmulas de validação/pagamento/frete | `AB/MOVEIS/modelo_ab.py` |
| Leitura do PTI (separador, encoding, vários arquivos) | `AB/FULL/modelo_ab_full.py` → `carregar_pti` |
| Colunas do empilhado ou a gravação no xlsx grande | `AB/empilhar.py`, `AB/empilhar_gravar.py` |
| Ordem de STATUS (REBATE, SEM DIFF, EST COM…) e cálculo de DIF | `AB/RB/cruzamento.py` |
| Config por canal da conciliação (loja, vendedor, onde está o Para Processar) | `AB/RB/canais.py` |
| Históricos e códigos de natureza do TOTVS | `AB/RB/historicos.py`, `AB/RB/dashboard.py` |
| Colunas, cores, copiar valor ou linhas CMM/MOV do dashboard | `AB/RB/dashboard.py` |
| O que o scraper lê no ML/MP e como classifica as linhas | `AB/RB/scraper_ml.py`, `AB/RB/classificar_ml.py` |
| Layout do CMM de rebate | `AB/RB/rebate_cmm.py`, `AB/RB/rebate_processar.py` |

## Módulos (documentação)

### Frente PROCESSAR (AB → Para Processar)

**`AB/koncili.py`** — Baixa o AB no Koncili e confere com o calendário. Precisa de `playwright` e das variáveis `KONCILI_USUARIO` / `KONCILI_SENHA`.
- `login(page, usuario, senha)`: Playwright headless. Não usa `networkidle`, porque a página nunca assenta.
- `cabecalhos_api(page)`: a API autentica pelo header `gumgatoken`, que a própria página envia.
- `achar_repasse(req, hdr, conta, data_iso)`: `POST /reconciliation/api/view/conciliation/search` e pega o repasse cujo período começa na data.
- `resultado(req, hdr, id_repasse)`: `kpis.result` = RESULTADO.
- `baixar_ab(page, id_repasse, destino)`: botão `#btn-view-spreadsheet`.
- `soma_ab`, `caminho_destino(canal, data)`, `_ajustar_layout` (aba e colunas no padrão do X:; CENTER/WEB usam `COLUNAS_WEB`, 35 colunas), `conferir(...)` (AB × RESULTADO × calendário).
- Contas em `CANAIS_KONCILI`. Nunca imprimir token ou sessão em log.

**`AB/rodar_canal.py`** — Comando único do Para Processar (FULL, ES, MOVEIS, ELETRO, CENTER ou TODOS).
- `CANAIS_PP`: configuração por canal (pasta e prefixo do AB, PTI, frete, pasta e nome de saída, ano no nome). CENTER adicionado 22/09/2026 (vendedor 944, PTI `C:\TEMP\PTI CENTER.csv`, sem frete — mesmo molde do MOVEIS).
- `caminhos(canal, data)`: devolve o AB de entrada e o nome do arquivo de saída.
- `processar(canal, data, pasta_saida, forcar)`: FULL usa `rodar_full.rodar` (2 vendedores no PTI); os demais usam `rodar_es.rodar`. Grava a aba "Para Processar" e imprime linhas, H:M e quantos "Verificar". Não sobrescreve sem `--forcar`.
- `main(argv)`: se o AB de um canal não está no X:, avisa e segue com os outros.

**`AB/FULL/processar_full.py`** — Motor do cruzamento (tradução das macros do PB BASE.xlsm).
- `carregar_ab`, `marcar_flags` e `filtrar_validos` (linhas válidas, colunas T/U do Para_Processar).
- `carregar_totvs()`: lê `TOTVS_2022`, aba "01-01 - Listagem do Browse" (cabeçalho nas linhas 1-4).
- `cruzar_totvs(df, totvs)`: chave **SHIPPING_ID**; marca `verificar` quando não acha.
- `resolver_pendentes_via_sri(pendentes)`: fallback no SRI para o que ficou "Verificar".
- `gerar_para_processar(cruzado)`: monta a aba final. `_limpar_id` remove o `.0` dos IDs.

**`AB/FULL/modelo_ab_full.py`** — `montar_ab_pp(para_processar)` cria os apelidos que a Validação espera (ARRUMAR, NOVO CREDITO, NOVO DEBITO, `CUSTO DE OFERECER DESCONTO` = 0). `carregar_pti(caminho)` lê o PTI (latin-1, separador `;`, aceita curinga com vários arquivos). Os outros canais reaproveitam.

**`AB/MOVEIS/modelo_ab.py`** — Fórmulas do MODELO AB MERCADO, compartilhadas por FULL, ES e MOVEIS.
- `conferencia_ab_original` e `calcular_ab_para_processar` (colunas P:U).
- `montar_validacao`: agrega por pedido, classifica (`Deduzir Cupom`, `Deduzir Taxa financiamento CMM`…) e devolve `Verificar` quando não fecha com o PTI.
- `montar_pagamento`, `montar_frete`, `montar_para_processar` (Pagamento + Frete + linha de ajuste) e `resumo_acao_manual`.
- As colunas de saída do Pagamento estão em `COLUNAS_PAGAMENTO` (15 colunas, H:M = valores).

**`AB/FULL/rodar_full.py`** (+ `rodar_full.bat`, `escolher_full.py`, `criar_pastas.bat`) — `rodar(caminho_ab, caminho_pti)` devolve (validação, pagamento, frete, processar). `--data DD.MM.AAAA` resolve AB e saída sozinho. `escolher_full.py` pergunta mês e arquivo. `tipagem_full.py` corrige colunas de valor que vêm como texto.

**`AB/ES/`** — `processar_es.carregar_ab` é o único ponto que muda por layout: traduz o export em português (43 colunas), aceita as 60 em outra ordem e completa as 35 do CENTER/WEB. `rodar_es.rodar(caminho_ab, caminho_pti, frete)`: com `caminho_pti=None` só faz o cruzamento TOTVS/SRI (ELETRO); com `frete=False` não lança frete (MOVEIS). `tipagem_es.py` guarda `LAYOUT_ES` e o mapa `PT_PARA_EN`; `modelo_ab_es.py` só reexporta `montar_ab_pp` e `carregar_pti` do FULL; `escolher_es.py` é o seletor de arquivo.

**`AB/empilhar.py` + `AB/empilhar_gravar.py`** — Empilhado FULL e ES: o AB original linha a linha mais o que o processo calcula (25 colunas, incluindo Nº QUE ESTÁ NO TOTVS, OPERAÇÃO, VALOR, SEQUENCIA e CHAVE_PRIMÁRIA).
- `empilhar.py`: `cruzar_dia`, `montar_linhas` e `ensaio` (monta e confere a soma de VALOR × calendário; cache em `AB/empilhar_cache/`).
- `empilhar_gravar.py`: acrescenta direto no XML do .xlsx (arquivo de 400+ MB; o openpyxl perderia os valores em cache das fórmulas). Textos como `inlineStr`, remove `calcChain`, corrige SEQUENCIA na cauda e acrescenta uma linha por dia na aba CALENDÁRIO. Backup em `C:\BACKUP EMPILHADO` antes de publicar. Rodar de novo com um dia já empilhado não duplica.
- Nunca abrir e regravar o empilhado com openpyxl.

**`AB/pipeline.py`** — Controle de lotes do FULL em SQLite (`conciliacao.db`): `sync` (lotes novos), `rodar <lote>` e `status`. Alimenta a tabela de lotes do dashboard. Diferente do lote "manual" do `rodar_canal.py` (lista de datas na linha de comando, sem persistir nada) — este rastreia status (pendente/rodado/erro) por lote no banco. Só existe pro FULL.

### Frente Conciliação (`AB/RB`)

**Entrada e orquestração**
- `interativo.py`: assistente (`iniciar_conciliacao.bat`). Pergunta canal e data e chama `rodar_conciliacao_rb.processar`. Modo direto: `python interativo.py --canal FULL --data DD.MM.AAAA [--scraper] [--saida-pasta X]`. Exige `AUXILIAR\rb.xlsx`, `rbx.xlsx` e `C:\TEMP\Conciliador - Backup\CONCILIADOR_RB.xlsx` com **Dt.Ref igual ao dia**; senão recusa.
- `rodar_conciliacao_rb.py`: `processar(canal, data, pasta_saida, usar_scraper)` agrega, cruza, enriquece, aplica rebate e salva. `_ordenar_fila_verificar` divide a fila em 2 grupos por sinal; `_arquivar_rbx` guarda cópia do rbx.
- `canais.py`: fonte única da config por canal (raiz do histórico TOTVS, onde está o Para Processar, chave do RB, loja do calendário, vendedor, prefixo MOV). `resolver_para_processar` e `validar_data_rb`.

**Cálculo**
- `parser_rb.py`: `metadata_rb` (Dt.Ref) e `carregar_rb` (relatório paginado, repete o cabeçalho a cada ~57 linhas e termina com totais).
- `agregador.py`: `agregar_por_pedido` (replica a macro `SubstituirEProcessar_V2`).
- `cruzamento.py`: `cruzar_com_rb`. STATUS na ordem REBATE (CUPOM ≠ 0) → SEM DIFF (DIF = 0) → EST COM (COMISSAO + DIF = 0, DIF < 0) → MEDIACAO (DIF − MEDIACAO = 0) → VERIFICAR. `DIF = REPASSE − TOTAL_RB`; `Dif. extrato = DIF − RB − RBX + TOTAL_RB`.
- `conferir_rbx.py`: confere as 2 contas. A **Original (103925)** tem que fechar com o calendário; a **Transitória (X103925)** acumula o movimento e, quando o líquido bate com o alvo, já dá para transferir para a Original. Inclui `conferir_pedidos`.
- `calendario_alvo.py`: `buscar_alvo(canal, data_ref)` e `buscar_alvo_loja(loja, data_ref)` (o empilhado usa a segunda).
- `enriquecer.py`: histórico e ação sugeridos por pedido; `_rebate_com_sinal_do_dif`.
- `historicos.py`, `extrair_historicos.py` e `historicos_totvs.json`: raiz do histórico TOTVS por loja (`_REPASSE`, `_COMISSAO`, `_FRETE`, `_ESTORNO`, `_ANUNCIO`), extraída da aba HISTORICOS do CONTROLE_V2.
- `ids.py`: `limpar_id` (sem `.0`).

**Scraper ML/MP**
- `scraper_ml.py`: `resumo_pedido(numero, context)` (SHIPPING_ID do ML começa com "4"; senão usa EXTERNAL_REFERENCE/MP — se o MP não achar nada, tenta de novo o mesmo número como SHIPPING_ID no ML antes de desistir, 22/09/2026) e `resumir_pedidos` (progresso e tentativas). Usa Chrome com debug remoto (porta 9222, `abrir_chrome_debug.bat`).
- `classificar_ml.py`: sem playwright; classifica linhas (comissão, envio, rebate, desconto/bônus) por substring.
- `descricoes_ml.py`: guarda todas as linhas raspadas em `descricoes_ml_<canal>_<data>.xlsx`.

**Rebate**
- `rebate_cmm.py`: linhas de rebate para a aba Movimento_Bancario do CMM (LOJA, VENDEDOR e HISTÓRICO vêm de `CANAIS`).
- `rebate_processar.py`: layout `PROCESSAR_REBATE` (só FULL), gravado em `C:\RODAR TOTVS\PROCESSAR\REBATE\`.
- `rebate_fora_da_regra.py`: lê o resultado do TOTVS (SpreadsheetML disfarçado de .xls ou .xlsx) e monta o CMM dos "Fora da Regra".
- `conferir_rbx.py` e `test_interativo.py` completam a pasta (conferência e teste do assistente).

**`AB/RB/dashboard.py`** — HTML único e offline. `gerar_dashboard(resultado, fila_verificar, rebate_cmm, …)` preenche o template com placeholders `@@NOME@@`.
- `_faixa_status`: mostra se o dia fecha com o calendário. Alvo e Calculado copiam só os dígitos (212.638,47 → `21263847`).
- `_linhas_tabela`: colunas da fila VERIFICAR, na ordem `Aj. | RBX | RB | Dif. extrato | AB total | Pedido | Descrição | DIF | Frete | Comissão | BANCO | [Marcação] | Desc/Bônus | Repasse | Rebate`. O cabeçalho (`<thead>`) fica na mesma função e tem que seguir a mesma ordem.
- BANCO = RBX + DIF (sem RBX, só a DIF). Frete aparece em valor absoluto quando a DIF é positiva. Frete, Comissão e BANCO copiam ao clicar (`_valor_copia`: vírgula decimal, sem milhar, ex.: `586,94`).
- Naturezas marcáveis, que geram linhas CMM e MOV na gaveta do rodapé: Frete 2209, Tarifa 2804 e Rebate 1107, todas na conta X103925, banco 237, agência 3708.
- Scroll da fila VERIFICAR persiste em `localStorage` (22/09/2026): sobrevive ao reload automático do `visibilitychange`.
- `localStorage`: `ajustados_<canal>_<data>` (+ `_aba`, `_natureza`, `_natureza_extra`, `_drawer`, `_scroll`) e `concfull:tema`.
- Autoteste: `python dashboard.py` → `dashboard.py: ok`.

## Cuidados conhecidos

- VPN caída derruba o acesso ao X:.
- `os.replace` não funciona de C: para X:; usar `shutil.move`.
- RB e RBX precisam ter a Dt.Ref do dia; o `interativo.py` recusa se não bater.
- O calendário às vezes está errado (ex.: MOVEIS 04/09). Conferir com o Koncili antes de achar que o processo errou.
- Editar arquivo `.py` por script via heredoc no Bash estraga barras invertidas (`\2026` vira octal). Usar o editor.

## Dados e segurança

- **Dados nunca são versionados**: `.xlsx`, `.xls`, `.db`, `Dados/` e `Projeto Imports/` estão no `.gitignore`. Contêm valores reais de repasse e pedido.
- Nenhuma credencial no código; o Koncili lê usuário e senha do ambiente.
- Documentação completa do código: nota **Documentação do Código - Conciliação Mercado** no Obsidian (vault `Manuais`).

## Qualidade e auditoria do processo

Processo financeiro sem CI/QA dedicado — a segurança vem de disciplina no próprio código, não de ferramenta externa:

- **Autoteste em cada módulo**: 26 dos 37 arquivos `.py` têm um bloco `__main__` com `assert` (387 no total) que roda sozinho — `python cruzamento.py`, `python dashboard.py` etc. imprimem `"<arquivo>: ok"` ou quebram na hora, sem precisar de framework de teste. Regra: lógica não trivial (branch, loop, parser, caminho de dinheiro) não é considerada pronta sem esse autoteste.
- **Validação contra dado real, não só teoria**: mais de 20 comentários no código documentam "achado real"/"confirmado com dado real" — casos em que a lógica foi comparada pedido a pedido contra o TOTVS/Koncili/ML antes de ser aceita (ex.: o achado do `ids.py` que motivou o commit `37e9f16`, ou o `MEDIACAO` removido dos STATUS em `cruzamento.py` depois de um pedido real que "batia por construção" sem nunca ter sido lançado).
- **Guarda-corpos que travam antes de processar errado**: `canais.validar_data_rb` recusa rodar se o RB exportado for de outro dia (senão o cruzamento zera TOTAL_RB pra quase todo mundo e vira uma fila VERIFICAR gigante e falsa); `interativo.py` exige `Dt.Ref` do dia em `rb.xlsx`/`rbx.xlsx`/`CONCILIADOR_RB.xlsx` antes de aceitar rodar.
- **Fila VERIFICAR como rede de segurança**: nada que não bateu automaticamente (REBATE/SEM DIFF/EST COM) sai do radar — cai em VERIFICAR pra conferência manual em vez de ser dado como resolvido por engano (foi assim que o STATUS MEDIACAO foi descartado, 22/09/2026).
- **Correção pela raiz, não pelo sintoma**: bug corrigido na função compartilhada que todos os canais usam (não só onde o sintoma apareceu) — antes de mexer, todos os usos da função são checados (`grep`) pra decidir o alcance certo do fix. Config por canal também é centralizada num único lugar (`canais.py`, `CANAIS_KONCILI`, `CANAIS_PP`) desde 16-18/09/2026, exatamente pra evitar que os mesmos dados divirjam em dicts espalhados sem ninguém notar.
- **Commit revisado, nunca em lote às cegas**: cada commit é conferido por `git diff` antes do `git add` e fica restrito aos arquivos da mudança em questão — mudanças pendentes de outra frente de trabalho não entram junto só porque estavam no working tree.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate       # Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium   # login do Koncili e scraper do ML
```

Nunca usou venv? Veja a aula básica em [`docs/VENV.md`](docs/VENV.md) (para que serve, como ativar e desativar).

O SRI (fallback do cruzamento em `processar_full.py`) e o scraper do ML/MP usam o projeto separado **Screaper Mercado**, clonado em `Desktop\Projetos\Screaper Mercado` (`Pagina.py` e o venv com playwright). O `interativo.py --scraper` troca para o Python desse venv sozinho. `koncili.py` também precisa de `playwright`, no venv atual ou no do Screaper.

Precisa de acesso à unidade `X:` (AB e empilhados) e ao TOTVS_2022 (cruzamento por SHIPPING_ID, com SRI como alternativa).

## Pendências

- [ ] **WEB**: Para Processar (falta o PTI do WEB e o código do vendedor; exemplos em `AB/WEB`)
- [ ] Empilhado de MOVEIS, ELETRO, PRIME, WEB e CENTER (layouts diferentes)
- [ ] Para Processar do MOVEIS/ELETRO dos dias 12 a 17 e 21/09
- [ ] ELETRO 11/09: H:M 95.010,82 contra 94.393,55 do calendário (Δ 617,27) — conferir com o Koncili
- [ ] Scraper: `MELIPAYMENTS-COLLECTIONATTEMPT…` retorna "não encontrado"
- [ ] Banco central do PTI (hoje são CSVs em `C:\TEMP`)
