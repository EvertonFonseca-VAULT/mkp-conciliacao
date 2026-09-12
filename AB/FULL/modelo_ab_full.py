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

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "MOVEIS"))

import pandas as pd
from modelo_ab import montar_validacao, montar_pagamento, montar_frete, montar_para_processar, resumo_acao_manual


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
    formato pt-BR (ponto de milhar, vírgula decimal)."""
    pti = pd.read_csv(caminho, dtype=str, encoding="latin-1", sep=";")
    for col in ["VALOR REAIS", "VALOR", "FRETE"]:
        pti[col] = pd.to_numeric(
            pti[col].str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce"
        )
    return pti


def _demo():
    para_processar = pd.DataFrame({
        "DATE": ["2026-09-01"] * 2,
        "SOURCE_ID": ["1", "2"],
        "EXTERNAL_REFERENCE": ["P1", "P2"],
        "RECORD_TYPE": ["release", "release"],
        "DESCRIPTION": ["payment", "payment"],
        "NET_CREDIT_AMOUNT": [90, 0],
        "NET_DEBIT_AMOUNT": [0, 200],
        "GROSS_AMOUNT": [100, -200],
        "MP_FEE_AMOUNT": [-10, 0],
        "FINANCING_FEE_AMOUNT": [0, 0],
        "SHIPPING_FEE_AMOUNT": [0, 0],
        "TAXES_AMOUNT": [0, 0],
        "COUPON_AMOUNT": [0, 0],
        "INSTALLMENTS": ["1", "1"],
        "PAYMENT_METHOD": ["visa", "visa"],
    })
    ab_pp = montar_ab_pp(para_processar)
    pti = pd.DataFrame({"PED. MKP": ["P1"], "VALOR REAIS": [100], "VALOR": [100], "FRETE": [0]})

    validacao = montar_validacao(ab_pp, pti)
    categorias = dict(zip(validacao["PED_MKP"], validacao["categoria"]))
    assert categorias["P1"] == "Manter igual", categorias
    assert categorias["P2"] == "Estorno", categorias
    assert list(resumo_acao_manual(validacao)["PED_MKP"]) == ["P2"]
    print("ok")


if __name__ == "__main__":
    _demo()
