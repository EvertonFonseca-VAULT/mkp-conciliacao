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
    # DESCRICAO: todas as descricoes da coluna E (AB) daquele pedido, unidas
    # -- pro Everton ver "mediation"/"refund"/etc na fila VERIFICAR sem
    # precisar abrir o AB pra procurar o pedido. Pedido do Everton, 18/09/2026.
    descricoes_por_pedido = para_processar.groupby("EXTERNAL_REFERENCE")["DESCRIPTION"].apply(
        lambda s: ", ".join(sorted(set(s.dropna().astype(str))))
    )

    df = agregado.merge(rb_por_pedido, how="left", left_on="PEDIDO_MKP", right_index=True)
    df[["SALDO_TOTVS", "ENTRADA", "SAIDA", "TOTAL_RB"]] = (
        df[["SALDO_TOTVS", "ENTRADA", "SAIDA", "TOTAL_RB"]].fillna(0.0)
    )

    df["REBATE"] = df["PEDIDO_MKP"].map(rebate).fillna(0.0).round(2)
    df["MEDIACAO"] = df["PEDIDO_MKP"].map(mediacao).fillna(0.0).round(2)
    df["DESCRICAO"] = df["PEDIDO_MKP"].map(descricoes_por_pedido).fillna("")
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
        "REBATE", "MEDIACAO", "DESCRICAO", "DIF", "STATUS",
    ]
    return df[colunas]


if __name__ == "__main__":
    agregado = pd.DataFrame({
        "PEDIDO_MKP": ["1", "2"], "REPASSE": [100.0, 44.05],
        "COMISSAO": [0.0, 0.0], "FRETE": [0.0, 0.0],
    })
    rb = pd.DataFrame({
        "PED_MKP": ["1"], "SALDO_TITULO": [0.0], "ENTRADAS": [100.0], "SAIDAS": [0.0],
    })
    # pedido "2": 2 linhas no AB (refund + mediation), confirmado com dado
    # real do pedido 2000017591518558 (08/09/2026) -- DESCRICAO junta as 2,
    # separadas por virgula, pra aparecer inteiro na fila VERIFICAR.
    para_processar = pd.DataFrame({
        "EXTERNAL_REFERENCE": ["2", "2"],
        "DESCRIPTION": ["refund", "mediation"],
        "COUPON_AMOUNT": [0.0, 0.0],
        "GROSS_AMOUNT": [0.0, 0.0],
    })
    resultado = cruzar_com_rb(agregado, rb, para_processar)
    linha1 = resultado[resultado["PEDIDO_MKP"] == "1"].iloc[0]
    assert linha1["STATUS"] == "SEM DIFF", linha1
    assert linha1["DESCRICAO"] == "", linha1  # sem linha no AB pro pedido "1"
    linha2 = resultado[resultado["PEDIDO_MKP"] == "2"].iloc[0]
    assert linha2["DESCRICAO"] == "mediation, refund", linha2  # ordem alfabetica
    print("cruzamento.py: ok")
