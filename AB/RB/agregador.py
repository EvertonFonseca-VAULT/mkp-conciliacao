"""Agrega o "Para Processar" (saida do pipeline, por canal/dia) por pedido
(EXTERNAL_REFERENCE), replicando a macro SubstituirEProcessar_V2 do
AB_MERCADO_TOTVS.xlsm (aba Planilha1)."""

import pandas as pd

from ids import limpar_pedido_ab


def agregar_por_pedido(para_processar):
    para_processar = para_processar.copy()
    para_processar["EXTERNAL_REFERENCE"] = para_processar["EXTERNAL_REFERENCE"].apply(limpar_pedido_ab)

    grupo = para_processar.groupby("EXTERNAL_REFERENCE", dropna=False)

    soma = grupo[["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT"]].sum()
    agregado = soma.reset_index().rename(columns={"EXTERNAL_REFERENCE": "PEDIDO_MKP"})

    agregado["REPASSE"] = (
        agregado["GROSS_AMOUNT"] + agregado["MP_FEE_AMOUNT"]
        + agregado["FINANCING_FEE_AMOUNT"] + agregado["SHIPPING_FEE_AMOUNT"]
    ).round(2)
    agregado["COMISSAO"] = (-agregado["MP_FEE_AMOUNT"]).round(2)
    agregado["FRETE"] = (-agregado["SHIPPING_FEE_AMOUNT"]).round(2)

    return agregado[["PEDIDO_MKP", "REPASSE", "COMISSAO", "FRETE"]]
