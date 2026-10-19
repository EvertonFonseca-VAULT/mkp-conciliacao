"""Roda a pipeline completa do Mercado COLETAS ES (AB bruto -> TOTVS/SRI ->
Validação -> Pagamento/Frete -> Para Processar) e salva a aba "Para
Processar" (formato de importação no TOTVS) em Excel.

Uso do dia a dia (via escolher_es.bat, igual escolher_full.bat do FULL):
    python rodar_es.py --data DD.MM.AAAA -- resolve sozinho o AB de entrada
    (pasta REDE, ver PASTA_AB_REDE_PADRAO) e ja salva a saida no nome/pasta
    que o TOTVS espera (C:\\RODAR TOTVS\\PROCESSAR\\COLETAS ES).

Uso manual/depuracao (caminhos explicitos, nao mexe na saida automatica):
    python rodar_es.py "<caminho do AB>" "<caminho do PTI>" [saida.xlsx]
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "MOVEIS"))

import pandas as pd

from processar_es import (
    carregar_ab, marcar_flags, filtrar_validos,
    carregar_totvs, cruzar_totvs, resolver_pendentes_via_sri, gerar_para_processar,
)
from modelo_ab_es import montar_ab_pp, carregar_pti
from modelo_ab import montar_validacao, montar_pagamento, montar_frete, montar_para_processar

CAMINHO_PTI_PADRAO = r"C:\TEMP\PTI ES.csv"

# Mesmo padrao de rede/saida que rodar_full.py usa pro FULL -- confirmado
# com os arquivos ja existentes em C:\RODAR TOTVS\PROCESSAR\COLETAS ES
# (nome "PROCESSAR_MERCADO COLETAS ES <data>.xlsx"). Pedido do Everton,
# 18/09/2026.
PASTA_SAIDA_PADRAO = r"C:\RODAR TOTVS\PROCESSAR\COLETAS ES"
PREFIXO_SAIDA_PADRAO = "PROCESSAR_MERCADO COLETAS ES"
PASTA_AB_REDE_PADRAO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB\COLETAS ES"


def caminhos_do_dia(data):
    """`data` = "DD.MM.AAAA". Deriva o caminho do AB bruto (pasta REDE,
    <ano>/<mes>.<ano>/) e o caminho de saida no padrao que o TOTVS espera
    (PASTA_SAIDA_PADRAO/PREFIXO_SAIDA_PADRAO <data>.xlsx) -- mesma funcao
    de rodar_full.caminhos_do_dia, so que com as constantes do ES."""
    dia, mes, ano = data.split(".")
    caminho_ab = os.path.join(PASTA_AB_REDE_PADRAO, ano, f"{mes}.{ano}", f"MERCADO COLETAS ES {data}.xlsx")
    saida = os.path.join(PASTA_SAIDA_PADRAO, f"{PREFIXO_SAIDA_PADRAO} {data}.xlsx")
    return caminho_ab, CAMINHO_PTI_PADRAO, saida


def rodar(caminho_ab, caminho_pti, frete=True, ajustar_cupom=False):
    """`caminho_pti=None` (ELETRO): so o cruzamento TOTVS/SRI, sem PTI/frete --
    devolve (None, None, None, para_processar). `frete=False` (MOVEIS): roda a
    Validacao/Pagamento mas nao lanca Frete. `ajustar_cupom=True` (WEB, ver
    CANAIS_PP["WEB"] em rodar_canal.py): cruza os pedidos do Para Processar
    final de volta com o AB original (por EXTERNAL_REFERENCE resolvido no
    TOTVS) e soma o EFFECTIVE_COUPON_AMOUNT do pedido na linha de
    COUPON_AMOUNT (coluna M) -- uma vez por pedido, na primeira linha que
    aparecer com essa referencia.

    ⚠️ Testado em 13/10/2026 (WEB 01.10.2026): aplicar o ajuste ANTES da
    classificacao (direto no AB, em COUPON_AMOUNT) nao muda o H:M (o cupom
    se cancela dentro de `_recalcular`, MOVEIS/modelo_ab.py) e so piora a
    classificacao (Verificar de 6 pra 59 de 128). Por isso o ajuste e feito
    DEPOIS, direto na linha final do Para Processar -- ai sim o H:M cai
    exatamente o valor do cupom."""
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
    if caminho_pti is None:
        return None, None, None, para_processar
    ab_pp = montar_ab_pp(para_processar)
    pti = carregar_pti(caminho_pti)

    validacao = montar_validacao(ab_pp, pti)
    pagamento = montar_pagamento(validacao, ab_pp)
    frete = montar_frete(validacao, ab_pp) if frete else montar_frete(validacao, ab_pp, pedidos=[])
    processar = montar_para_processar(validacao, pagamento, frete, ab_pp)
    if ajustar_cupom:
        processar = _somar_cupom_efetivo(processar, cruzado)
    return validacao, pagamento, frete, processar


def _somar_cupom_efetivo(processar, cruzado):
    """Cruza o Para Processar de volta com o AB original (chave =
    EXTERNAL_REFERENCE do Para Processar == totvs_encontrado do AB cruzado)
    e soma o EFFECTIVE_COUPON_AMOUNT do pedido em COUPON_AMOUNT (coluna M) --
    uma vez por pedido, na primeira linha que aparecer com essa referencia
    (pedido pode ter linha de Pagamento + Frete + Ajuste, cupom entra uma
    vez so)."""
    processar = processar.copy()
    cupom_por_pedido = cruzado.groupby("totvs_encontrado")["EFFECTIVE_COUPON_AMOUNT"].sum()
    cupom_por_pedido = cupom_por_pedido[cupom_por_pedido != 0]
    primeira_linha = processar.reset_index().groupby("EXTERNAL_REFERENCE")["index"].first()
    for pedido, valor in cupom_por_pedido.items():
        idx = primeira_linha.get(pedido)
        if idx is not None:
            processar.loc[idx, "COUPON_AMOUNT"] += valor
    return processar


def main():
    # --data DD.MM.AAAA (uso do dia a dia, via escolher_es.bat): deriva AB de
    # entrada e saida sozinho, no padrao que o TOTVS espera. Sem --data, cai
    # no modo manual de sempre (caminhos explicitos).
    if len(sys.argv) > 2 and sys.argv[1] == "--data":
        caminho_ab, caminho_pti, saida = caminhos_do_dia(sys.argv[2])
    else:
        caminho_ab = sys.argv[1]
        caminho_pti = sys.argv[2] if len(sys.argv) > 2 else CAMINHO_PTI_PADRAO
        saida = sys.argv[3] if len(sys.argv) > 3 else "resultado_es.xlsx"

    os.makedirs(os.path.dirname(saida) or ".", exist_ok=True)
    _validacao, _pagamento, _frete, processar = rodar(caminho_ab, caminho_pti)

    with pd.ExcelWriter(saida) as writer:
        processar.to_excel(writer, sheet_name="Para Processar", index=False)

    print(f"AB: {caminho_ab}")
    print(f"PTI: {caminho_pti}")
    print(f"Linhas em Para Processar: {len(processar)}")
    print(f"Salvo em: {os.path.abspath(saida)}")


def _teste():
    """Offline: _somar_cupom_efetivo isolado (sem AB/TOTVS/PTI reais)."""
    processar = pd.DataFrame({
        "EXTERNAL_REFERENCE": ["P1", "P1", "P2", "P3"],  # P1: Pagamento + Frete
        "COUPON_AMOUNT": [10.0, 0.0, 5.0, 2.0],
    })
    cruzado = pd.DataFrame({
        "totvs_encontrado": ["P1", "P1", "P2", "P9"],  # P1 em 2 linhas do AB, P9 nao existe no Para Processar
        "EFFECTIVE_COUPON_AMOUNT": [-1.0, -2.0, 0.0, -99.0],
    })
    ajustado = _somar_cupom_efetivo(processar, cruzado)
    # P1: cupom do AB = -1-2=-3, entra 1x so na 1a linha (10 -> 7); P2: cupom 0, nao muda; P3: sem match, nao muda.
    assert ajustado["COUPON_AMOUNT"].tolist() == [7.0, 0.0, 5.0, 2.0]
    print("rodar_es.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main()
