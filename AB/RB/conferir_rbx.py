"""Confere as 2 contas usadas na conciliacao do dia: a conta ORIGINAL
(103925, sem "X") tem que fechar com o valor do calendario, e a conta
TRANSITORIA (X103925) acumula o movimento do dia -- NAO compara contra
zero. Quando o liquido dela bate com o MESMO alvo do calendario, e sinal
de que ja pode fazer a transferencia manual pra Original (que ai sim zera
a Transitoria). Pedido do Everton, 16-17/09/2026 -- correcao de regra:
a 1a versao comparava a Transitoria contra zero (transitoria_zera), o que
estava errado.

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
from ids import limpar_id
from parser_rb import carregar_rb
from rebate_cmm import MOTIVO_REBATE

TOLERANCIA = 0.01

# Nao e por canal (ao contrario de chave_rbx) -- cancelamento de RA
# (Recebimento Antecipado) na transitoria, ex. "CAN-EXCLUSAO DE TITULO RA",
# nao pertence a reconciliacao de nenhum canal especifico (nao tem a chave
# tipo "FU") e por isso ficava invisivel: nunca entrava em
# liquido_transitoria nem aparecia em lugar nenhum do dashboard, so
# escondido dentro do "NAO zerou". Pedaco comum do rotulo, nao o texto
# inteiro -- podem existir outras variacoes de cancelamento de RA. Pedido
# do Everton, 17/09/2026.
CHAVE_RA = "TITULO RA"


def _liquido(df, chave):
    mask = df["OPERACAO"].str.contains(chave, case=False, na=False)
    linhas = df.loc[mask]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2), len(linhas)


def _valores_lancados(df):
    """Valor liquido (ENTRADAS - SAIDAS) lancado por pedido, agrupado pela
    chave PED_MKP quando preenchido, senao DOCUMENTO (limpar_id) -- mesmo
    criterio de antes (_pedidos_lancados), so que agora soma o valor em vez
    de so marcar presenca: lancamento de frete/cobranca feito via MOV (ex.
    MKP_ME_FULL_TARIFA DE COBRANCA) vem do TOTVS com PED_MKP vazio e o
    pedido SO no DOCUMENTO (confirmado nos pedidos 2000017834651356,
    2000018022992924, 2000018087525336 do rbx.xlsx do FULL 02.09.2026).
    DOCUMENTO passa por limpar_id igual PED_MKP (celula numerica no Excel
    vira float, senao "123.0" nunca bate com "123"). Usado pro dashboard
    mostrar o valor real lancado em cada conta ao lado da fila VERIFICAR,
    e calcular a Diferenca extrato (DIF - valor RB - valor RBX). Pedido do
    Everton, 16/09/2026."""
    chave = df["PED_MKP"].where(df["PED_MKP"] != "", df["DOCUMENTO"].apply(limpar_id))
    liquido = (df["ENTRADAS"] - df["SAIDAS"]).groupby(chave).sum()
    return {k: round(v, 2) for k, v in liquido.items() if k}


def conferir(canal, alvo, pasta_auxiliar, rebate_total=None):
    """None se nao tiver rb.xlsx/rbx.xlsx em pasta_auxiliar (feature
    manual/opcional -- so roda quando o Everton gerou os 2 arquivos) ou se
    o canal ainda nao tiver chave definida."""
    cfg = CANAIS.get(canal, {})
    chave = cfg.get("chave_rbx")
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
    liquido_ra, qtd_ra = _liquido(transitoria, CHAVE_RA)
    liquido_transitoria_simulado = round(liquido_transitoria + liquido_ra, 2)
    # Correcao de regra, 17/09/2026: a Transitoria NAO zera durante o dia --
    # ela acumula o movimento. Quando o liquido dela (chave do canal + RA)
    # bate com o MESMO alvo do calendario usado em original_bate, e sinal
    # de que ja pode fazer a transferencia manual pra Original (que ai sim
    # zera a Transitoria). Comparar contra 0 (antigo transitoria_zera)
    # estava errado.
    #
    # 2o estado do ciclo, 17/09/2026: depois que a transferencia manual e
    # feita de verdade, a Transitoria volta a 0 -- "zerada" tambem e um
    # estado CERTO (transferencia ja feita), nao o mesmo "ainda nao bate
    # com o alvo" de quando esta so acumulando no meio do dia.
    zerada = abs(liquido_transitoria_simulado) < TOLERANCIA
    pronta_para_transferencia = alvo is not None and abs(liquido_transitoria_simulado - alvo) < TOLERANCIA

    # Rebate: 2 formas de cair no RBX -- direto do ML (chave_cupom_rebate,
    # pedido "deu certo via processar") ou lancado a mao no CMM
    # (prefixo_mov + MOTIVO_REBATE, pra quem nao deu certo via processar).
    # Soma dos 2 tem que bater com o Rebate (formula) do card -- semaforo no
    # dashboard. Pedido do Everton, 18/09/2026.
    chave_cupom_rebate = cfg.get("chave_cupom_rebate")
    liquido_cupom_rebate, qtd_cupom_rebate = (
        _liquido(transitoria, chave_cupom_rebate) if chave_cupom_rebate else (0.0, 0)
    )
    prefixo_mov = cfg.get("prefixo_mov")
    chave_cmm_rebate = f"{prefixo_mov}_{MOTIVO_REBATE}" if prefixo_mov else None
    liquido_cmm_rebate, qtd_cmm_rebate = (
        _liquido(transitoria, chave_cmm_rebate) if chave_cmm_rebate else (0.0, 0)
    )
    liquido_rebate_rbx = round(liquido_cupom_rebate + liquido_cmm_rebate, 2)
    rebate_bate = (
        rebate_total is not None and abs(liquido_rebate_rbx - rebate_total) < TOLERANCIA
    )

    return {
        "chave": chave,
        "liquido_original": liquido_original,
        "qtd_original": qtd_original,
        "original_bate": alvo is not None and abs(liquido_original - alvo) < TOLERANCIA,
        "liquido_transitoria": liquido_transitoria,
        "qtd_transitoria": qtd_transitoria,
        # Cancelamento de RA na transitoria (CHAVE_RA) -- fora da chave do
        # canal, mas ja somado em liquido_transitoria_simulado abaixo (ver
        # comentario da constante). Pedido do Everton, 17/09/2026.
        "liquido_ra": liquido_ra,
        "qtd_ra": qtd_ra,
        "liquido_transitoria_simulado": liquido_transitoria_simulado,
        "qtd_transitoria_simulado": qtd_transitoria + qtd_ra,
        "zerada": zerada,
        "pronta_para_transferencia": pronta_para_transferencia,
        # Valor liquido ja lancado por pedido em cada conta (PED_MKP ou
        # DOCUMENTO, ver _valores_lancados) -- usado pro dashboard mostrar,
        # pedido a pedido da fila VERIFICAR, quanto ja foi lancado no RB
        # (original) e/ou no RBX (transitoria). Pedido do Everton,
        # 15-16/09/2026.
        "valores_rb": _valores_lancados(original),
        "valores_rbx": _valores_lancados(transitoria),
        "liquido_cupom_rebate": liquido_cupom_rebate,
        "qtd_cupom_rebate": qtd_cupom_rebate,
        "liquido_cmm_rebate": liquido_cmm_rebate,
        "qtd_cmm_rebate": qtd_cmm_rebate,
        "liquido_rebate_rbx": liquido_rebate_rbx,
        "rebate_bate": rebate_bate,
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
        for operacao, entradas, saidas, documento, ped_mkp in linhas:
            ws.append(["01/09/2026", operacao, documento, "", entradas, saidas, 0, "", ped_mkp, "", "", 0])
        ws.append(["", "NAO CONCILIADOS", "CONCILIADOS", "TOTAL"])
        wb.save(caminho)

    with tempfile.TemporaryDirectory() as tmp:
        _escrever(os.path.join(tmp, "rb.xlsx"), "103925", [
            ("MKP_FULL_REPASSE", 100.0, 0, "", ""),
            ("MKP_FULL_CANCELAMENTO", 10.0, 0, "", ""),
            ("CAN-REC MKP MKP_MELI_FU REF PARC 1", 0, 10.0, "1", "1"),
        ])
        _escrever(os.path.join(tmp, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 100.0, 0, "1", "1"),
            ("MKP_FULL_REPASSE", 0, 100.0, "", ""),
            ("MKP_FULL_CANCELAMENTO", 0, 10.0, "", ""),
            ("MKP_ME_FULL_RECEBIMENTO DE REBATE", 10.0, 0, "", ""),
            # lancamento via MOV (frete/cobranca): pedido SO no DOCUMENTO,
            # PED_MKP vazio -- e o bug real do FULL 02.09.2026.
            ("MKP_ME_FULL_TARIFA DE COBRANCA", 0, 5.0, "2000017834651356", ""),
            # cancelamento de RA -- nao tem "FU" (chave_rbx do FULL), fica
            # de fora do liquido_transitoria/qtd_transitoria de proposito.
            ("CAN-EXCLUSAO DE TITULO RA", 0, 7.0, "", ""),
        ])

        r = conferir("FULL", alvo=100.0, pasta_auxiliar=tmp)
        # rebate (formula) 10.0 bate com RBX: so a linha CMM (sem cupom
        # nesse fixture) -- ver bloco isolado mais abaixo pro caso com as
        # 2 fontes juntas.
        assert r["liquido_cmm_rebate"] == 10.0, r
        assert r["qtd_cmm_rebate"] == 1, r
        assert r["liquido_cupom_rebate"] == 0.0, r
        assert r["qtd_cupom_rebate"] == 0, r
        assert r["liquido_rebate_rbx"] == 10.0, r
        assert r["rebate_bate"] is False  # sem rebate_total, nada pra comparar
        assert bool(conferir("FULL", alvo=100.0, pasta_auxiliar=tmp, rebate_total=10.0)["rebate_bate"])
        assert not bool(conferir("FULL", alvo=100.0, pasta_auxiliar=tmp, rebate_total=999.0)["rebate_bate"])
        assert r["liquido_original"] == 100.0, r
        assert bool(r["original_bate"]), r
        assert r["liquido_transitoria"] == -5.0, r
        # -5.0 (chave) + -7.0 (RA) = -12.0, nao bate com o alvo (100.0) --
        # ainda NAO esta pronta pra transferencia (nao e "nao zerou": a
        # Transitoria nao compara contra 0, so contra o alvo do calendario)
        assert r["liquido_transitoria_simulado"] == -12.0, r
        assert not bool(r["pronta_para_transferencia"]), r
        assert "transitoria_zera" not in r  # regra antiga (comparar contra 0) removida
        # so a linha CAN-REC tem PED MKP preenchido: entradas 0 - saidas 10
        assert r["valores_rb"] == {"1": -10.0}, r
        # "1" vem do PED_MKP da linha REC MKP (100-0=100); "2000017834651356"
        # vem do DOCUMENTO da linha MOV, PED_MKP vazio ali (0-5=-5) -- sem a
        # uniao, ficaria de fora (falso negativo no dashboard).
        assert r["valores_rbx"] == {"1": 100.0, "2000017834651356": -5.0}, r
        # cancelamento de RA: fora da chave "FU" do canal, so aparece aqui
        assert r["liquido_ra"] == -7.0, r
        assert r["qtd_ra"] == 1, r

        assert not bool(conferir("FULL", alvo=999.0, pasta_auxiliar=tmp)["original_bate"])
        assert conferir("XPTO", alvo=100.0, pasta_auxiliar=tmp) is None  # canal sem chave
        assert conferir("FULL", alvo=100.0, pasta_auxiliar=tmp + "_nada") is None  # sem arquivos

    # liquido_transitoria_simulado (chave + RA) batendo exatamente com o
    # alvo (90.0 + 10.0 = 100.0) -- sinal de que ja pode fazer a
    # transferencia manual pra Original. Prova que pronta_para_transferencia
    # compara contra o alvo, nao contra zero.
    with tempfile.TemporaryDirectory() as tmp2:
        _escrever(os.path.join(tmp2, "rb.xlsx"), "103925", [
            ("MKP_FULL_REPASSE", 100.0, 0, "", ""),
        ])
        _escrever(os.path.join(tmp2, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 90.0, 0, "1", "1"),
            ("CAN-EXCLUSAO DE TITULO RA", 10.0, 0, "", ""),
        ])
        r2 = conferir("FULL", alvo=100.0, pasta_auxiliar=tmp2)
        assert r2["liquido_transitoria"] == 90.0, r2
        assert r2["liquido_ra"] == 10.0, r2
        assert r2["liquido_transitoria_simulado"] == 100.0, r2
        assert bool(r2["pronta_para_transferencia"]), r2  # bate com o alvo -> pronta
        assert not bool(r2["zerada"]), r2  # 100.0 nao e zero

        assert not bool(conferir("FULL", alvo=999.0, pasta_auxiliar=tmp2)["pronta_para_transferencia"])
        assert conferir("FULL", alvo=None, pasta_auxiliar=tmp2)["pronta_para_transferencia"] is False

    # 2o estado do ciclo: transferencia manual ja feita, a Transitoria voltou
    # a 0 (chave -10.0 + RA +10.0 = 0.0) -- "zerada" e certo mesmo sem bater
    # com o alvo (alvo continua sendo o valor do dia, nao zero).
    with tempfile.TemporaryDirectory() as tmp3:
        _escrever(os.path.join(tmp3, "rb.xlsx"), "103925", [
            ("MKP_FULL_REPASSE", 100.0, 0, "", ""),
        ])
        _escrever(os.path.join(tmp3, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 0, 10.0, "1", "1"),
            ("CAN-EXCLUSAO DE TITULO RA", 10.0, 0, "", ""),
        ])
        r3 = conferir("FULL", alvo=100.0, pasta_auxiliar=tmp3)
        assert r3["liquido_transitoria"] == -10.0, r3
        assert r3["liquido_ra"] == 10.0, r3
        assert r3["liquido_transitoria_simulado"] == 0.0, r3
        assert bool(r3["zerada"]), r3  # transferencia ja feita
        assert not bool(r3["pronta_para_transferencia"]), r3  # 0.0 != alvo (100.0)

    # Rebate com as 2 fontes juntas no RBX: cupom direto do ML (chave_cupom_
    # rebate) + lancamento manual no CMM (prefixo_mov + MOTIVO_REBATE) --
    # soma tem que bater com o Rebate (formula). Isolado num tmp proprio pra
    # nao mexer nos numeros de liquido_transitoria ja testados acima ("FU"
    # tambem casa com "FULL", entao a linha de cupom conta pra la).
    with tempfile.TemporaryDirectory() as tmp4:
        _escrever(os.path.join(tmp4, "rb.xlsx"), "103925", [])
        _escrever(os.path.join(tmp4, "rbx.xlsx"), "X103925", [
            ("MKP CUPOM MELI FULL REPASSE", 15.0, 0, "", ""),
            ("MKP_ME_FULL_RECEBIMENTO DE REBATE", 10.0, 0, "", ""),
        ])
        r4 = conferir("FULL", alvo=None, pasta_auxiliar=tmp4, rebate_total=25.0)
        assert r4["liquido_cupom_rebate"] == 15.0, r4
        assert r4["qtd_cupom_rebate"] == 1, r4
        assert r4["liquido_cmm_rebate"] == 10.0, r4
        assert r4["qtd_cmm_rebate"] == 1, r4
        assert r4["liquido_rebate_rbx"] == 25.0, r4
        assert bool(r4["rebate_bate"]), r4  # 25.0 == 25.0 -> bate

    print("conferir_rbx.py: ok")
