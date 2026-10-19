"""Roda a pipeline completa do Mercado FULL (AB bruto -> TOTVS/SRI ->
Validação -> Pagamento/Frete -> Para Processar) e salva a aba "Para
Processar" (formato de importação no TOTVS) em Excel.

Uso do dia a dia (via rodar_full.bat, igual o iniciar_conciliacao.bat do
AB/RB): python rodar_full.py --data DD.MM.AAAA -- resolve sozinho o AB de
entrada (pasta REDE, ver CAMINHO_AB_PADRAO) e já salva a saída no nome/pasta
exatos que o AB/RB (resolver_para_processar) espera achar, fechando a ponte
entre os dois pipelines sem precisar mover arquivo na mao.

Uso manual/depuracao (caminhos explicitos, nao mexe na saida automatica):
    python rodar_full.py "<caminho do AB>" "<caminho do PTI>" [saida.xlsx]

Sem nenhum argumento, usa os caminhos do lote 01.09.2026 (mesmo usado na
validacao contra PROCESSAR_MERCADO FULL 01.09.xlsx).
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
# Curinga: o FULL tem 2 vendedores no PTI (750 e 751), um arquivo cada (ex.: "PTI FULL.csv" e "PTI FULL 750.csv").
CAMINHO_PTI_PADRAO = r"C:\TEMP\PTI FULL*.csv"

# Mesma pasta/prefixo que canais.py (AB/RB) usa em CANAIS["FULL"] pra achar o
# "Para Processar" do dia -- salvando aqui com esse nome exato, o AB/RB acha
# sozinho, sem precisar mover o arquivo na mao. Pedido do Everton, 17/09/2026.
PASTA_SAIDA_PADRAO = r"C:\RODAR TOTVS\PROCESSAR\FULL"
PREFIXO_SAIDA_PADRAO = "PROCESSAR_MERCADO FULL"

# Pasta REDE onde o AB bruto do dia fica (Capa!B2 do PB BASE.xlsm, ver Passo a
# passo - Conciliacao Mercado FULL.md) -- <ANO>/<MM>.<ANO>/MERCADO FULL <data>.xlsx
PASTA_AB_REDE_PADRAO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB\FULL"


def caminhos_do_dia(data):
    """`data` = "DD.MM.AAAA". Deriva o caminho do AB bruto (pasta REDE,
    <ano>/<mes>.<ano>/) e o caminho de saida no padrao que o AB/RB espera
    (PASTA_SAIDA_PADRAO/PREFIXO_SAIDA_PADRAO <data>.xlsx)."""
    dia, mes, ano = data.split(".")
    caminho_ab = os.path.join(PASTA_AB_REDE_PADRAO, ano, f"{mes}.{ano}", f"MERCADO FULL {data}.xlsx")
    saida = os.path.join(PASTA_SAIDA_PADRAO, f"{PREFIXO_SAIDA_PADRAO} {data}.xlsx")
    return caminho_ab, CAMINHO_PTI_PADRAO, saida


def rodar(caminho_ab, caminho_pti):
    ab = carregar_ab(caminho_ab)
    validos = filtrar_validos(marcar_flags(ab))

    cruzado = cruzar_totvs(validos, carregar_totvs())
    pendentes = cruzado[cruzado["verificar"]]
    if len(pendentes):
        resolvidos = resolver_pendentes_via_sri(pendentes)
        cruzado.loc[resolvidos.index, "totvs_encontrado"] = resolvidos["totvs_encontrado"]
        # resolvidos["verificar"] (nao False fixo): sem o scraper SRI disponivel,
        # resolver_pendentes_via_sri devolve o pendente sem resolver (verificar
        # continua True) -- fixo em False perderia a rede de seguranca da fila
        # VERIFICAR pra quem roda sem o SRI. Achado real, colega sem Screaper
        # Mercado, CENTER 09.10.2026.
        cruzado.loc[resolvidos.index, "verificar"] = resolvidos["verificar"]

    para_processar = gerar_para_processar(cruzado)
    ab_pp = montar_ab_pp(para_processar)
    pti = carregar_pti(caminho_pti)

    validacao = montar_validacao(ab_pp, pti)
    pagamento = montar_pagamento(validacao, ab_pp)
    frete = montar_frete(validacao, ab_pp)
    processar = montar_para_processar(validacao, pagamento, frete, ab_pp)
    return validacao, pagamento, frete, processar


def main():
    # --data DD.MM.AAAA (uso do dia a dia, via rodar_full.bat): deriva AB de
    # entrada e saida sozinho, no padrao que o AB/RB espera. Sem --data, cai
    # no modo manual de sempre (caminhos explicitos ou default fixo de teste).
    if len(sys.argv) > 2 and sys.argv[1] == "--data":
        caminho_ab, caminho_pti, saida = caminhos_do_dia(sys.argv[2])
    else:
        caminho_ab = sys.argv[1] if len(sys.argv) > 1 else CAMINHO_AB_PADRAO
        caminho_pti = sys.argv[2] if len(sys.argv) > 2 else CAMINHO_PTI_PADRAO
        saida = sys.argv[3] if len(sys.argv) > 3 else "resultado_full.xlsx"

    os.makedirs(os.path.dirname(saida) or ".", exist_ok=True)
    _validacao, _pagamento, _frete, processar = rodar(caminho_ab, caminho_pti)

    with pd.ExcelWriter(saida) as writer:
        processar.to_excel(writer, sheet_name="Para Processar", index=False)

    print(f"AB: {caminho_ab}")
    print(f"PTI: {caminho_pti}")
    print(f"Linhas em Para Processar: {len(processar)}")
    print(f"Salvo em: {os.path.abspath(saida)}")


if __name__ == "__main__":
    main()
