"""Cruza o agregado por pedido com o RB e classifica o STATUS de cada
pedido, replicando a aba "CONCILIADOR MERCADO" do CONTROLE_V2.xlsx.

STATUS (nessa ordem):
    REBATE existe (coluna CUPOM != 0)      -> REBATE, independente do DIF
        (correcao do Everton: nao e so quando o cupom explica a diferenca --
        todo pedido com valor em CUPOM e rebate, ponto, mesmo se o banco ja
        bateu certo sem sobrar diferenca)
    DIF = 0                                -> SEM DIFF
    COMISSAO + DIF = 0, COMISSAO != 0, DIF < 0  -> EST COM (estorno de comissao)
    DIF - MEDIACAO = 0, DIF != 0           -> MEDIACAO
    senao                                  -> VERIFICAR (fila de ajuste manual)
"""

import numpy as np
import pandas as pd

from ids import limpar_id

# NAO CONFIRMADO: o texto exato de DESCRIPTION para linhas de mediacao pode
# vir como "mediation" (ingles, visto em linhas residuais geradas pelo
# proprio pipeline Python) ou "Mediação" (portugues, usado na formula real
# do CONTROLE_V2). Comparamos por substring case-insensitive ate validar
# contra um dia real com pedidos MEDIACAO.
_MARCADOR_MEDIACAO = "medi"


def cruzar_com_rb(agregado, rb, para_processar):
    rb_por_pedido = rb.groupby("PED_MKP").agg(
        SALDO_TOTVS=("SALDO_TITULO", "sum"),
        ENTRADA=("ENTRADAS", "sum"),
        SAIDA=("SAIDAS", "sum"),
    )
    rb_por_pedido["SAIDA"] = -rb_por_pedido["SAIDA"]
    rb_por_pedido["TOTAL_RB"] = rb_por_pedido["ENTRADA"] + rb_por_pedido["SAIDA"]

    para_processar = para_processar.copy()
    para_processar["EXTERNAL_REFERENCE"] = para_processar["EXTERNAL_REFERENCE"].apply(limpar_id)

    rebate = para_processar.groupby("EXTERNAL_REFERENCE")["COUPON_AMOUNT"].sum()

    descricao = para_processar["DESCRIPTION"].astype(str).str.lower()
    mediacao = (
        para_processar[descricao.str.contains(_MARCADOR_MEDIACAO, na=False)]
        .groupby("EXTERNAL_REFERENCE")["GROSS_AMOUNT"].sum()
    )

    df = agregado.merge(rb_por_pedido, how="left", left_on="PEDIDO_MKP", right_index=True)
    df[["SALDO_TOTVS", "ENTRADA", "SAIDA", "TOTAL_RB"]] = (
        df[["SALDO_TOTVS", "ENTRADA", "SAIDA", "TOTAL_RB"]].fillna(0.0)
    )

    df["REBATE"] = df["PEDIDO_MKP"].map(rebate).fillna(0.0).round(2)
    df["MEDIACAO"] = df["PEDIDO_MKP"].map(mediacao).fillna(0.0).round(2)
    df["TOTAL_PED"] = (df["REPASSE"] + df["COMISSAO"] + df["FRETE"]).round(2)
    df["DIF"] = (df["REPASSE"] - df["TOTAL_RB"]).round(2)

    df["STATUS"] = np.select(
        [
            df["REBATE"] != 0,
            df["DIF"] == 0,
            (df["COMISSAO"] + df["DIF"] == 0) & (df["COMISSAO"] != 0) & (df["DIF"] < 0),
            (df["DIF"] - df["MEDIACAO"] == 0) & (df["DIF"] != 0),
        ],
        ["REBATE", "SEM DIFF", "EST COM", "MEDIACAO"],
        default="VERIFICAR",
    )

    colunas = [
        "PEDIDO_MKP", "REPASSE", "COMISSAO", "FRETE", "TOTAL_PED",
        "SALDO_TOTVS", "ENTRADA", "SAIDA", "TOTAL_RB",
        "REBATE", "MEDIACAO", "DIF", "STATUS",
    ]
    return df[colunas]
