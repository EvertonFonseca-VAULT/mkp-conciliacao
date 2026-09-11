# Conciliação de Marketplace

Automatização do processo de conciliação financeira de marketplace (AB, PTI, PB, TOTVS, CMM/MOV), hoje feito majoritariamente via planilhas, macros e Power Query manuais.

## Objetivo

Substituir o processo manual — disperso em dezenas de macros e planilhas auxiliares sem documentação — por um banco central versionado, importações automatizadas e, no futuro, a conciliação ponta a ponta.

## Estado atual

O que existe hoje é a base do banco e a importação do Magalu. As demais etapas do processo (extração automática do AB/PTI, geração do PROCESSAR, importação no ERP, conciliação TOTVS/CMM/MOV) ainda são manuais e ficam como próximos passos — não há orquestração (Airflow/dbt/etc.) implementada ainda.

## Estrutura

```
Mkp/
├── sqlserver/           # implementação: conexão, import, consultas, schema
│   ├── conexao.py        # login Azure AD (cache de token/AuthenticationRecord)
│   ├── importar.py        # roda o schema.sql e importa as planilhas do Magalu
│   ├── funcoes.py          # consultas/resumos prontos sobre o banco
│   ├── schema.sql           # DDL das tabelas (ver "Banco" abaixo)
│   ├── testar_conexao.py     # smoke test de conexão
│   └── README.md              # setup, driver, autenticação — detalhes técnicos
├── Dados/magalu/          # bases novas do Magalu a importar (não versionado)
└── Projeto Imports/        # histórico consolidado + planilha de canais (não versionado)
```

Dados brutos (`Dados/`, `Projeto Imports/`, qualquer `.xlsx`) não são versionados — ver `.gitignore`. Contêm valores reais de repasse/pedido.

## Banco (Microsoft Fabric SQL database)

- **`canais`** — cadastro canal + loja + vertente (1P/3P/SEM ID), extraído da planilha de planejamento.
- **`conciliacao_diaria`** — uma linha por canal + data: extrato bancário (`valor_eb`), arquivo de baixa (`valor_ab`), valor no sistema interno (`valor_rb`) e a diferença calculada entre eles. `valor_eb`/`valor_rb` ainda são preenchidos manualmente.
- **`mkp_magalu`** — dado transacional detalhado (pedido/parcela) do Magalu, alimentado pelo `importar.py`.
- **`importacoes_log`** — controla quais arquivos já foram importados (hash SHA-256), pra evitar reimport duplicado.

Detalhes de conexão, autenticação e passo a passo de execução: [`sqlserver/README.md`](sqlserver/README.md).

## Convenções

- **Dados nunca são versionados** — apenas código.
- **Commits pequenos e frequentes**, com mensagem descrevendo a mudança de forma clara.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Próximos passos

- [ ] Import automático de EB e RB (hoje manuais em `conciliacao_diaria`)
- [ ] Mapear outros marketplaces além do Magalu
- [ ] Resolver os 11 canais com `loja = 'SEM ID'` (nomes reais de loja pendentes)
- [ ] Automatizar extração (AB/PTI) e geração do PROCESSAR
- [ ] Importação automática no ERP (TOTVS) e conciliação (CMM/MOV)
