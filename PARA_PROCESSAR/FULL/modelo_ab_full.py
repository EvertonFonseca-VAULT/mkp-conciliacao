"""Validação/Pagamento/Frete/Para Processar do Mercado FULL — reaproveita
as funções genéricas de `AB/MOVEIS/modelo_ab.py` (mesma fórmula, canais
diferentes; ver docstring de lá).

Diferença do MOVEIS: o AB do FULL já vem no layout final em inglês (não
tem uma "AB Original" em português por trás) — não passa por
`calcular_ab_para_processar`. O `ab_pp` aqui é a própria saída de
`processar_full.gerar_para_processar`, só com os apelidos de coluna que
`montar_validacao`/`montar_pagamento` esperam (ARRUMAR, NOVO CREDITO,
NOVO DEBITO) e `CUSTO DE OFERECER DESCONTO`=0 (conceito que não existe
nesse canal).

Validado contra `C:\\RODAR TOTVS\\PROCESSAR\\PROCESSAR_MERCADO FULL
01.09.xlsx`: 340/340 linhas, 184/184 pedidos batendo em NET_CREDIT_AMOUNT/
NET_DEBIT_AMOUNT/SHIPPING_FEE_AMOUNT (os valores realmente lançados no
TOTVS). GROSS_AMOUNT/MP_FEE_AMOUNT batem 183/184 — 1 pedido no arquivo
real quebra o próprio padrão dele (isolado, não afeta o valor líquido).
"""

import glob

import pandas as pd


def montar_ab_pp(para_processar):
    """`para_processar` = saída de `processar_full.gerar_para_processar`
    (15 colunas, EXTERNAL_REFERENCE já resolvido via TOTVS/SRI)."""
    df = para_processar.copy()
    df["ARRUMAR"] = df["EXTERNAL_REFERENCE"].astype(str).str.strip()
    df["NOVO CREDITO"] = df["NET_CREDIT_AMOUNT"]
    df["NOVO DEBITO"] = df["NET_DEBIT_AMOUNT"]
    df["CUSTO DE OFERECER DESCONTO"] = 0
    return df


def carregar_pti(caminho):
    """PTI do FULL: export TOTVS em latin-1, separado por `;`, valores em
    formato pt-BR (ponto de milhar, vírgula decimal). `caminho` pode ter
    curinga ("PTI FULL*.csv"): o FULL tem 2 vendedores (750 e 751) e cada um
    sai num arquivo -- junta todos (e ignora linha repetida)."""
    arquivos = sorted(glob.glob(caminho)) or [caminho]
    pti = pd.concat([pd.read_csv(a, dtype=str, encoding="latin-1", sep=";") for a in arquivos], ignore_index=True).drop_duplicates()
    for col in ["VALOR REAIS", "VALOR", "FRETE"]:
        pti[col] = pd.to_numeric(
            pti[col].str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce"
        )
    return pti
