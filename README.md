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
| PRIME, WEB, CENTER | `WEBCOPRIME`, `WEB CONTINENTAL`, `CONTINENTALCENTER` | — | — | — | só baixa o AB (WEB em implantação) |

Nos códigos de vendedor, o PTI usa `000` na frente (`000755`); CMM/MOV usam sem zeros.

## Estrutura

```
Mkp/
├── requirements.txt
└── AB/
    ├── koncili.py            # login no Koncili, acha o repasse, lê o RESULTADO, baixa o AB e confere com o calendário
    ├── rodar_canal.py        # Para Processar de FULL, ES, MOVEIS e ELETRO (ou TODOS) num comando
    ├── empilhar.py           # empilhado FULL/ES: ensaio + gravação (com backup em C:\BACKUP EMPILHADO)
    ├── empilhar_gravar.py    # grava as linhas novas direto no XML do xlsx, sem duplicar dias
    ├── pipeline.py           # execução por lote do FULL com status em SQLite (só FULL)
    ├── FULL/                 # cruzamento AB × TOTVS/SRI + PTI (2 vendedores) e layout de 40 colunas
    ├── ES/                   # layout de 60 colunas (também lê 43 em português e 35 do WEB/CENTER)
    ├── MOVEIS/               # modelo_ab.py: validação, pagamento, frete e Para Processar (compartilhado)
    ├── WEB/                  # exemplos de Para Processar do WEB (implantação)
    └── RB/                   # conciliação com o RB, classificação e dashboard
```

Layouts do AB bruto: FULL 40 colunas · ES 60 · CENTER/WEB 35 · MOVEIS/ELETRO 43 ou 48 (português).

## Uso diário

```bash
cd AB

# 1. Baixa o AB do Koncili e confere com o calendário
python koncili.py ES 22.09.2026                  # só confere
python koncili.py ELETRO 22.09.2026 --salvar     # grava no X: (não sobrescreve)

# 2. Para Processar (não sobrescreve sem --forcar)
python rodar_canal.py TODOS 22.09.2026
python rodar_canal.py FULL 11.09.2026 18.09.2026

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

## Dados e segurança

- **Dados nunca são versionados**: `.xlsx`, `.xls`, `.db`, `Dados/` e `Projeto Imports/` estão no `.gitignore`. Contêm valores reais de repasse e pedido.
- Nenhuma credencial no código; o Koncili lê usuário e senha do ambiente.
- Documentação completa do código: nota **Documentação do Código - Conciliação Mercado** no Obsidian (vault `Manuais`).

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate       # Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium   # login do Koncili e scraper do ML
```

Precisa de acesso à unidade `X:` (AB e empilhados) e ao TOTVS_2022 (cruzamento por SHIPPING_ID, com SRI como alternativa).

## Pendências

- [ ] **WEB**: Para Processar (falta o PTI do WEB e o código do vendedor; exemplos em `AB/WEB`)
- [ ] Empilhado de MOVEIS, ELETRO, PRIME, WEB e CENTER (layouts diferentes)
- [ ] Para Processar do MOVEIS/ELETRO dos dias 12 a 17 e 21/09
- [ ] ELETRO 11/09: H:M 95.010,82 contra 94.393,55 do calendário (Δ 617,27) — conferir com o Koncili
- [ ] Scraper: `MELIPAYMENTS-COLLECTIONATTEMPT…` retorna "não encontrado"
- [ ] Banco central do PTI (hoje são CSVs em `C:\TEMP`)
