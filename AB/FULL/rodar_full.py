"""Roda a pipeline completa do Mercado FULL (AB bruto -> TOTVS/SRI ->
Validação -> Pagamento/Frete -> Para Processar) e salva a aba "Para
Processar" (formato de importação no TOTVS) em Excel.

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
from modelo_ab import montar_validacao, montar_pagamento, montar_frete, montar_para_processar

CAMINHO_AB_PADRAO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB\FULL\2026\09.2026\MERCADO FULL 03.09.2026.xlsx"
CAMINHO_PTI_PADRAO = r"C:\TEMP\PTI FULL.csv"


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

    _validacao, _pagamento, _frete, processar = rodar(caminho_ab, caminho_pti)

    with pd.ExcelWriter(saida) as writer:
        processar.to_excel(writer, sheet_name="Para Processar", index=False)

    print(f"AB: {caminho_ab}")
    print(f"PTI: {caminho_pti}")
    print(f"Linhas em Para Processar: {len(processar)}")
    print(f"Salvo em: {os.path.abspath(saida)}")


if __name__ == "__main__":
    main()
