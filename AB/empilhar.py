"""Empilhado do Mercado (ARQUIVOS REPASSES COM CHAVES PRIMARIAS) -- ENSAIO.

O empilhado (aba Base de FULL MERCADO_2024.xlsx / COLETAS_ES_MERCADO_2024.xlsx)
e o AB ORIGINAL linha a linha (nao o Para Processar) mais o que o processo
calcula: Nº QUE ESTA NO TOTVS (resultado do cruzamento TOTVS/SRI), OPERACAO
(RECORD_TYPE_DESCRIPTION), VALOR, SEQUENCIA e CHAVE_PRIMARIA. Mapeamento
validado em 21/09/2026 com o FULL de 31/08 (1300 linhas do empilhado == 1300 do
AB, 0 diferenca de valor, soma de VALOR == calendario).

Este passo so MONTA as linhas do dia e confere a soma de VALOR contra o
calendario -- nao grava nada no empilhado. Guarda cada dia num pickle (cache)
pra a gravacao usar depois.

Uso:
    python empilhar.py ensaio FULL 11.09.2026 14.09.2026 ...
    python empilhar.py --teste
"""

import os
import pickle
import sys
import time

_AQUI = os.path.dirname(os.path.abspath(__file__))
for sub in ("FULL", "ES", "MOVEIS", "RB"):
    sys.path.insert(0, os.path.join(_AQUI, sub))
sys.path.insert(0, _AQUI)

import pandas as pd

import processar_es
import processar_full
from calendario_alvo import buscar_alvo_loja
from koncili import CANAIS_KONCILI, caminho_destino

CACHE = os.path.join(_AQUI, "empilhar_cache")
CARREGAR_AB = {"FULL": processar_full.carregar_ab, "ES": processar_es.carregar_ab}
CHAVE_SUFIXO = {"FULL": "FULL_MERCADO", "ES": "COLETAS_MERCADO"}  # CHAVE_PRIMARIA = <SEQUENCIA>_<sufixo>

# Colunas do empilhado, na ordem do arquivo (25; a 26a fica em branco).
COLUNAS = [
    "DATA_REPASSE", "DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "DESCRIÇÃO 1", "DESCRIÇÃO 2", "VALOR REPASSE", "ESTORNO",
    "GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT",
    "INSTALLMENTS", "PAYMENT_METHOD", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID",
    "Nº QUE ESTÁ NO TOTVS", "OPERAÇÃO", "VALOR", "SEQUENCIA", "CHAVE_PRIMÁRIA",
]


def cruzar_dia(canal, data, totvs):
    """AB do dia + TOTVS/SRI -> DataFrame do AB com `totvs_encontrado`
    (mesmos passos do rodar_full/rodar_es, ate antes do Para Processar)."""
    ab = CARREGAR_AB[canal](caminho_destino(canal, data))
    validos = processar_full.filtrar_validos(processar_full.marcar_flags(ab))
    cruz = processar_full.cruzar_totvs(validos, totvs)
    pend = cruz[cruz["verificar"]]
    if len(pend):
        r = processar_full.resolver_pendentes_via_sri(pend)
        cruz.loc[r.index, "totvs_encontrado"] = r["totvs_encontrado"]
        cruz.loc[r.index, "verificar"] = False
    return cruz


def montar_linhas(cruz, data):
    """Linhas do empilhado (SEQUENCIA/CHAVE ficam vazias: sao numeradas na gravacao)."""
    dia, mes, ano = (int(p) for p in data.split("."))
    num = lambda c: pd.to_numeric(cruz[c], errors="coerce").fillna(0)
    df = pd.DataFrame({
        "DATA_REPASSE": pd.Timestamp(ano, mes, dia),
        "DATE": cruz["RELEASE_DATE"],
        "SOURCE_ID": cruz["SOURCE_ID"],
        "EXTERNAL_REFERENCE": cruz["EXTERNAL_REFERENCE"],
        "DESCRIÇÃO 1": cruz["RECORD_TYPE"],
        "DESCRIÇÃO 2": cruz["DESCRIPTION"],
        "VALOR REPASSE": num("NET_CREDIT_AMOUNT"),
        "ESTORNO": num("NET_DEBIT_AMOUNT"),
        "GROSS_AMOUNT": num("GROSS_AMOUNT"), "MP_FEE_AMOUNT": num("MP_FEE_AMOUNT"),
        "FINANCING_FEE_AMOUNT": num("FINANCING_FEE_AMOUNT"), "SHIPPING_FEE_AMOUNT": num("SHIPPING_FEE_AMOUNT"),
        "TAXES_AMOUNT": num("TAXES_AMOUNT"), "COUPON_AMOUNT": num("COUPON_AMOUNT"),
        "INSTALLMENTS": cruz["INSTALLMENTS"], "PAYMENT_METHOD": cruz["PAYMENT_METHOD"],
        "SHIPPING_ID": cruz["SHIPPING_ID"], "SHIPMENT_MODE": cruz["SHIPMENT_MODE"],
        "ORDER_ID": cruz["ORDER_ID"], "PACK_ID": cruz["PACK_ID"],
        "Nº QUE ESTÁ NO TOTVS": cruz["totvs_encontrado"],
        "OPERAÇÃO": cruz["RECORD_TYPE"].astype(str) + "_" + cruz["DESCRIPTION"].astype(str),
        "VALOR": num("NET_CREDIT_AMOUNT") - num("NET_DEBIT_AMOUNT"),
        "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None,
    })
    return df[COLUNAS].reset_index(drop=True)


def ensaio(canal, datas):
    """Monta e confere cada dia; devolve {data: (linhas, soma_valor, alvo)}."""
    os.makedirs(CACHE, exist_ok=True)
    totvs = processar_full.carregar_totvs()
    print(f"TOTVS: {len(totvs)} titulos carregados", flush=True)
    resumo = {}
    for data in datas:
        t0 = time.time()
        df = montar_linhas(cruzar_dia(canal, data, totvs), data)
        dia, mes, ano = data.split(".")
        alvo = buscar_alvo_loja(CANAIS_KONCILI[canal]["loja"], f"{dia}/{mes}/{ano}")
        soma = round(df["VALOR"].sum(), 2)
        status = "sem linha no calendario" if alvo is None else ("OK" if abs(soma - alvo) < 0.01 else "DIFERE")
        pickle.dump(df, open(os.path.join(CACHE, f"{canal}_{data}.pkl"), "wb"))
        resumo[data] = (len(df), soma, alvo)
        print(f"  {canal} {data}: {len(df)} linhas | soma VALOR {soma:,.2f} | calendario {('-' if alvo is None else f'{alvo:,.2f}')} | {status} ({time.time()-t0:.0f}s)", flush=True)
    return resumo


def _teste():
    """Offline: 25 colunas e mapeamento de uma linha sintetica."""
    assert len(COLUNAS) == 25 and COLUNAS[22] == "VALOR"
    cruz = pd.DataFrame({
        "RELEASE_DATE": ["2026-09-01T10:00:00.000-03:00"], "SOURCE_ID": [1], "EXTERNAL_REFERENCE": [2000], "RECORD_TYPE": ["release"],
        "DESCRIPTION": ["refund"], "NET_CREDIT_AMOUNT": [0.0], "NET_DEBIT_AMOUNT": [640.0], "GROSS_AMOUNT": [0.0], "MP_FEE_AMOUNT": [0.0],
        "FINANCING_FEE_AMOUNT": [0.0], "SHIPPING_FEE_AMOUNT": [0.0], "TAXES_AMOUNT": [0.0], "COUPON_AMOUNT": [0.0], "INSTALLMENTS": [1],
        "PAYMENT_METHOD": ["visa"], "SHIPPING_ID": [47], "SHIPMENT_MODE": ["me2"], "ORDER_ID": [9], "PACK_ID": [None], "totvs_encontrado": ["47"],
    })
    df = montar_linhas(cruz, "01.09.2026")
    assert df.loc[0, "VALOR"] == -640 and df.loc[0, "ESTORNO"] == 640 and df.loc[0, "OPERAÇÃO"] == "release_refund"
    assert df.loc[0, "Nº QUE ESTÁ NO TOTVS"] == "47" and str(df.loc[0, "DATA_REPASSE"])[:10] == "2026-09-01"
    print("empilhar.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    elif len(sys.argv) > 3 and sys.argv[1] == "ensaio" and sys.argv[2].upper() in CARREGAR_AB:
        ensaio(sys.argv[2].upper(), sys.argv[3:])
    else:
        sys.exit("uso: python empilhar.py ensaio <FULL|ES> DD.MM.AAAA [...]  |  python empilhar.py --teste")
