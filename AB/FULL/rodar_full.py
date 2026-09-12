"""Roda a pipeline completa do Mercado FULL (AB bruto -> TOTVS/SRI ->
Validação -> Pagamento/Frete -> Para Processar) e salva o resultado em
Excel pra conferência manual.

Uso:
    python rodar_full.py "<caminho do AB>" "<caminho do PTI>" [saida.xlsx]

Se os caminhos não forem passados, usa os do lote 01.09.2026 (mesmo usado
na validação contra PROCESSAR_MERCADO FULL 01.09.xlsx).
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "MOVEIS"))

import pandas as pd

from processar_full import (
    carregar_ab, marcar_flags, filtrar_validos,
    carregar_totvs, cruzar_totvs, resolver_pendentes_via_sri, gerar_para_processar,
)
from modelo_ab_full import montar_ab_pp, carregar_pti
from modelo_ab import montar_validacao, montar_pagamento, montar_frete, montar_para_processar, resumo_acao_manual

CAMINHO_AB_PADRAO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB\FULL\2026\09.2026\MERCADO FULL 01.09.2026.xlsx"
CAMINHO_PTI_PADRAO = r"C:\TEMP\PTI FULL.csv"

# H:M no layout de 15 colunas (Pagamento/Frete/Para Processar): GROSS_AMOUNT
# até COUPON_AMOUNT. A soma dessas 6 colunas deve bater com NET_CREDIT-NET_DEBIT.
COLUNAS_H_M = ["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT"]


def com_conferencia(df):
    df = df.copy()
    df["SOMA H:M"] = df[COLUNAS_H_M].sum(axis=1).round(2)
    df["BATE"] = df["SOMA H:M"] == (df["NET_CREDIT_AMOUNT"] - df["NET_DEBIT_AMOUNT"]).round(2)
    return df


def rodar(caminho_ab, caminho_pti):
    ab = carregar_ab(caminho_ab)
    validos = filtrar_validos(marcar_flags(ab))

    cruzado = cruzar_totvs(validos, carregar_totvs())
    pendentes = cruzado[cruzado["verificar"]]
    if len(pendentes):
        resolvidos = resolver_pendentes_via_sri(pendentes)
        cruzado.loc[resolvidos.index, "totvs_encontrado"] = resolvidos["totvs_encontrado"]
        cruzado.loc[resolvidos.index, "verificar"] = False

    para_processar = gerar_para_processar(cruzado)
    ab_pp = montar_ab_pp(para_processar)
    pti = carregar_pti(caminho_pti)

    validacao = montar_validacao(ab_pp, pti)
    pagamento = montar_pagamento(validacao, ab_pp)
    frete = montar_frete(validacao, ab_pp)
    processar = montar_para_processar(validacao, pagamento, frete, ab_pp)
    return validacao, pagamento, frete, processar


def main():
    caminho_ab = sys.argv[1] if len(sys.argv) > 1 else CAMINHO_AB_PADRAO
    caminho_pti = sys.argv[2] if len(sys.argv) > 2 else CAMINHO_PTI_PADRAO
    saida = sys.argv[3] if len(sys.argv) > 3 else "resultado_full.xlsx"

    validacao, pagamento, frete, processar = rodar(caminho_ab, caminho_pti)

    with pd.ExcelWriter(saida) as writer:
        validacao.to_excel(writer, sheet_name="Validacao", index=False)
        com_conferencia(pagamento).to_excel(writer, sheet_name="Pagamento", index=False)
        com_conferencia(frete).to_excel(writer, sheet_name="Frete", index=False)
        com_conferencia(processar).to_excel(writer, sheet_name="Para Processar", index=False)
        resumo_acao_manual(validacao).to_excel(writer, sheet_name="Acao Manual", index=False)

    print(f"AB: {caminho_ab}")
    print(f"PTI: {caminho_pti}")
    print(f"Pedidos na Validação: {len(validacao)}")
    print(f"Linhas em Para Processar: {len(processar)}")
    print(f"Fila de Ação Manual: {len(resumo_acao_manual(validacao))}")
    print(f"Salvo em: {os.path.abspath(saida)}")


if __name__ == "__main__":
    main()
