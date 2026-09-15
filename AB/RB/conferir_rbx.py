"""Confere as 2 contas usadas na conciliacao do dia: a conta ORIGINAL
(103925, sem "X") tem que fechar com o valor do calendario, e a conta
TRANSITORIA (X103925) tem que zerar (entradas - saidas = 0). Pedido do
Everton, 16/09/2026.

Mecanica (validada com dados reais do FULL 01.09.2026, arquivos rb.xlsx +
rbx.xlsx que o Everton salva manualmente em AUXILIAR/): o dinheiro entra
pedido a pedido na TRANSITORIA durante o dia; no fim do dia sai de la em
lancamentos agregados (ex. MKP_FULL_REPASSE, MKP_FULL_CANCELAMENTO) e entra
na ORIGINAL com os mesmos rotulos. Pedidos cancelados DEPOIS de ja terem
sido lancados na ORIGINAL deixam um debito residual la (CAN-REC MKP...),
que o lancamento agregado de cancelamento compensa exatamente -- por isso
o LIQUIDO da conta original (nao so a linha _REPASSE isolada) bate com o
alvo do calendario.

A "chave" de busca no OPERACAO (substring, case-insensitive) NAO e a raiz
inteira do canal -- e o pedaco que aparece nos rotulos agregados de fim de
dia (ex. FULL usa "FU", nao "MKP_MELI_FU"). Vem de CANAIS[canal]["chave_rbx"]
em canais.py (fonte unica dessa config por canal) -- so confirmado pro FULL
ate agora, ver comentarios la.
"""

import os

from canais import CANAIS
from parser_rb import carregar_rb

TOLERANCIA = 0.01


def _liquido(df, chave):
    mask = df["OPERACAO"].str.contains(chave, case=False, na=False)
    linhas = df.loc[mask]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2), len(linhas)


def conferir(canal, alvo, pasta_auxiliar):
    """None se nao tiver rb.xlsx/rbx.xlsx em pasta_auxiliar (feature
    manual/opcional -- so roda quando o Everton gerou os 2 arquivos) ou se
    o canal ainda nao tiver chave definida."""
    chave = CANAIS.get(canal, {}).get("chave_rbx")
    if not chave:
        return None
    caminho_rb = os.path.join(pasta_auxiliar, "rb.xlsx")
    caminho_rbx = os.path.join(pasta_auxiliar, "rbx.xlsx")
    if not (os.path.exists(caminho_rb) and os.path.exists(caminho_rbx)):
        return None

    original = carregar_rb(caminho_rb)  # conta ORIGINAL (103925)
    transitoria = carregar_rb(caminho_rbx)  # conta TRANSITORIA (X103925)

    liquido_original, qtd_original = _liquido(original, chave)
    liquido_transitoria, qtd_transitoria = _liquido(transitoria, chave)

    return {
        "chave": chave,
        "liquido_original": liquido_original,
        "qtd_original": qtd_original,
        "original_bate": alvo is not None and abs(liquido_original - alvo) < TOLERANCIA,
        "liquido_transitoria": liquido_transitoria,
        "qtd_transitoria": qtd_transitoria,
        "transitoria_zera": abs(liquido_transitoria) < TOLERANCIA,
    }


if __name__ == "__main__":
    import tempfile

    import openpyxl

    def _escrever(caminho, conta, linhas):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Dt.Ref: 01/09/2026"])
        ws.append(["Hora: 10:00:00"])
        ws.append(["Emissão: 01/09/2026"])
        ws.append(["BANCO:", "237 - BRADESCO SA", "AGENCIA:", "3708", "CONTA:", conta, "SALDO INICIAL:", "0,00"])
        ws.append(["DATA", "OPERACAO", "DOCUMENTO", "PREFIXO/TITULO", "ENTRADAS", "SAIDAS",
                    "SALDO ATUAL", "USUARIO", "PED MKP", "Vendedor", "Emissao Titulo", "Saldo Titulo"])
        for operacao, entradas, saidas, pedido in linhas:
            ws.append(["01/09/2026", operacao, pedido, "", entradas, saidas, 0, "", pedido, "", "", 0])
        ws.append(["", "NAO CONCILIADOS", "CONCILIADOS", "TOTAL"])
        wb.save(caminho)

    with tempfile.TemporaryDirectory() as tmp:
        _escrever(os.path.join(tmp, "rb.xlsx"), "103925", [
            ("MKP_FULL_REPASSE", 100.0, 0, ""),
            ("MKP_FULL_CANCELAMENTO", 10.0, 0, ""),
            ("CAN-REC MKP MKP_MELI_FU REF PARC 1", 0, 10.0, "1"),
        ])
        _escrever(os.path.join(tmp, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 100.0, 0, "1"),
            ("MKP_FULL_REPASSE", 0, 100.0, ""),
            ("MKP_FULL_CANCELAMENTO", 0, 10.0, ""),
            ("MKP_ME_FULL_RECEBIMENTO DE REBATE", 10.0, 0, ""),
        ])

        r = conferir("FULL", alvo=100.0, pasta_auxiliar=tmp)
        assert r["liquido_original"] == 100.0, r
        assert bool(r["original_bate"]), r
        assert r["liquido_transitoria"] == 0.0, r
        assert bool(r["transitoria_zera"]), r

        assert not bool(conferir("FULL", alvo=999.0, pasta_auxiliar=tmp)["original_bate"])
        assert conferir("XPTO", alvo=100.0, pasta_auxiliar=tmp) is None  # canal sem chave
        assert conferir("FULL", alvo=100.0, pasta_auxiliar=tmp + "_nada") is None  # sem arquivos

    print("conferir_rbx.py: ok")
