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
em canais.py (fonte unica dessa config por canal).

chave_rbx pode ser 1 string OU uma lista/tupla de strings alternativas
(OR) -- FULL tem sorte de "FU" aparecer em TODOS os tipos de lancamento
(pedido a pedido "MELI_FU", agregado "MKP_FULL_..."), mas o ES nao: pedido
a pedido usa "MKP_MELI_ES" (sem "CO"), agregado/manual usa "MKP_MELI_CO_"
(com "CO", mas sem "ES") -- confirmado pelo Everton, 18/09/2026, com rotulos
reais (REC MKP MKP_MELI_ES REF PARC 1, MKP_MELI_CO_ES_REPASSE,
MKP_MELI_CO_RECEBIMENTO DE REBATE etc). O "_" no final de "MKP_MELI_CO_" e
proposital: sem ele bateria tambem com "MKP_MELI_CO05_..." (canal "Mercado
Coletas 05", confirmado que lanca nos MESMOS rb.xlsx/rbx.xlsx do ES).
"""

import os

import pandas as pd

from canais import CANAIS
from ids import limpar_id
from parser_rb import carregar_rb
from rebate_cmm import MOTIVO_REBATE

TOLERANCIA = 0.01


def _mask_chave(operacao, chave):
    """OPERACAO contendo QUALQUER UMA das chaves (chave = 1 string ou uma
    lista/tupla de strings alternativas, ver docstring do modulo) -- usado
    em vez de um .str.contains() direto em todo lugar que recebe a chave do
    canal, pra suportar canais como o ES que nao tem 1 raiz curta comum a
    todos os tipos de lancamento."""
    chaves = [chave] if isinstance(chave, str) else list(chave)
    mask = pd.Series(False, index=operacao.index)
    for c in chaves:
        mask |= operacao.str.contains(c, case=False, na=False)
    return mask

# Nao e por canal (ao contrario de chave_rbx) -- cancelamento de RA
# (Recebimento Antecipado) na transitoria, ex. "CAN-EXCLUSAO DE TITULO RA",
# nao pertence a reconciliacao de nenhum canal especifico (nao tem a chave
# tipo "FU") e por isso ficava invisivel: nunca entrava em
# liquido_transitoria nem aparecia em lugar nenhum do dashboard, so
# escondido dentro do "NAO zerou". Pedaco comum do rotulo, nao o texto
# inteiro -- podem existir outras variacoes de cancelamento de RA. Pedido
# do Everton, 17/09/2026.
CHAVE_RA = "TITULO RA"

# Cancelamento pedido-a-pedido na Original ("CAN-REC MKP <raiz> REF PARC N",
# confirmado com dado real do FULL 09/09/2026 -- mesmo padrao em outros
# canais: MERCADO CENTER, MKP_MELI_COL05) -- pedido que ja tinha sido
# lancado (e ja transferido) num dia anterior e foi cancelado hoje, deixando
# so o debito residual na Original, sem entrada correspondente na
# Transitoria do dia (o credito dele ja saiu da Transitoria fazia tempo).
# Isso "infla" o liquido_transitoria_simulado de hoje em relacao ao alvo, na
# mesma medida do debito -- por isso o valor de transferencia real e
# liquido_transitoria_simulado + liquido_cancelado_original (cancelado vem
# negativo, entao soma = subtrai o excedente). Pedido do Everton,
# 18/09/2026.
#
# CORRECAO 22/09/2026 (achado real do Everton, FULL 14/09/2026, R$ 1.219,82):
# quando o Everton JA lanca o par de compensacao (credito na Original +
# debito na Transitoria, rotulo tipo "MKP_MELI_FU_CANCELAMENTOS", mesmo doc
# nos 2), o debito na Transitoria ja entra em liquido_transitoria por causa
# do chave_rbx normal (ela tambem contem "FU") -- subtrair
# liquido_cancelado_original de novo em cima disso conta o mesmo lancamento
# 2x. So faz sentido subtrair a PARTE que ainda nao tem par na Transitoria
# (cenario original desta constante, pedido de dia anterior sem nada
# lancado hoje) -- ver CHAVE_CANCELAMENTO_AGREGADO e
# _cancelamento_agregado_transitoria.
CHAVE_CANCELAMENTO_ORIGINAL = "CAN-REC"

# Lancamento agregado de compensacao ("MKP_FULL_CANCELAMENTO",
# "MKP_MELI_FU_CANCELAMENTOS") -- sai da Transitoria e entra na Original com
# o mesmo rotulo (ver docstring do modulo), pra compensar exatamente os
# CAN-REC do dia. Substring generico (singular/plural).
CHAVE_CANCELAMENTO_AGREGADO = "CANCELAMENTO"


def _liquido(df, chave):
    mask = _mask_chave(df["OPERACAO"], chave)
    linhas = df.loc[mask]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2), len(linhas)


def _liquido_vendedor(df, chave, vendedor):
    """Como _liquido, mas tambem exige VENDEDOR == vendedor -- usado so pro
    RA (CHAVE_RA), que ao contrario de chave_rbx nao tem nenhuma chave de
    OPERACAO que identifique o canal ("CAN-EXCLUSAO DE TITULO RA" e
    generico). O VENDEDOR e o unico jeito confiavel de saber se aquele RA e
    do canal certo ou de outro produto/canal. Bug real reportado pelo
    Everton, 18/09/2026: FULL 09/09/2026 tinha RA de outro canal (vendedor
    "000423"/vazio, nenhum "751") contando como se fosse do FULL --
    Transitoria mostrando R$ 130.062,82 em vez de R$ 136.068,20."""
    mask = (
        _mask_chave(df["OPERACAO"], chave)
        & (df["VENDEDOR"].astype(str).str.strip() == vendedor)
    )
    linhas = df.loc[mask]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2), len(linhas)


def _cancelados_original(df, chave):
    """Cancelamentos do canal na Original (CHAVE_CANCELAMENTO_ORIGINAL +
    chave do canal, ver constante acima) -- liquido, quantidade e a lista
    pedido/valor (pro dashboard listar cada um). Pedido do Everton,
    18/09/2026."""
    mask = (
        df["OPERACAO"].str.contains(CHAVE_CANCELAMENTO_ORIGINAL, case=False, na=False)
        & _mask_chave(df["OPERACAO"], chave)
    )
    linhas = df.loc[mask]
    detalhe = [
        (str(row["PED_MKP"]) if row["PED_MKP"] else str(row["DOCUMENTO"]),
         round(row["ENTRADAS"] - row["SAIDAS"], 2))
        for _, row in linhas.iterrows()
    ]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2), len(linhas), detalhe


def _cancelamento_agregado_transitoria(df, chave):
    """Liquido do lancamento agregado de cancelamento JA na Transitoria hoje
    (CHAVE_CANCELAMENTO_AGREGADO + chave do canal) -- a parte de
    liquido_cancelado_original que esse lancamento ja cobre (ver correcao
    22/09/2026 na docstring de CHAVE_CANCELAMENTO_ORIGINAL)."""
    mask = (
        df["OPERACAO"].str.contains(CHAVE_CANCELAMENTO_AGREGADO, case=False, na=False)
        & _mask_chave(df["OPERACAO"], chave)
    )
    linhas = df.loc[mask]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2)


def _rebate_cancelado_rbx(df, chaves_rebate):
    """Cupom/CMM de rebate cancelado hoje (rotulo comecando com "CAN-", ex.
    "CAN-MKP CUPOM MELI FULL REPASSE") -- pedido que recebeu rebate direto
    do ML/CMM em outro dia e teve o cupom estornado hoje (ex.: entrou em
    disputa). Ja esta DENTRO de liquido_cupom_rebate/liquido_cmm_rebate (a
    soma normal por chave inclui o CAN-, ele so casa a chave de novo) -- essa
    funcao existe pra EXPLICAR a diferenca contra rebate_total (formula), que
    so soma pedidos com STATUS==REBATE hoje e nao sabe desse estorno (o
    pedido cancelado pode nem aparecer como REBATE hoje, so como VERIFICAR ou
    nem aparecer). Achado real do Everton, FULL 14/09/2026: pedido
    2000018111220588 (VERIFICAR, REBATE=0 na formula de hoje) com
    CAN-MKP CUPOM MELI FULL REPASSE de -136,16 na RBX -- rebate_bate dava
    'nao bate' por uma diferenca 100% explicada."""
    chaves = [c for c in chaves_rebate if c]
    if not chaves:
        return 0.0
    mask = _mask_chave(df["OPERACAO"], chaves) & df["OPERACAO"].str.upper().str.startswith("CAN-")
    linhas = df.loc[mask]
    return round((linhas["ENTRADAS"] - linhas["SAIDAS"]).sum(), 2)


def _valores_lancados(df, chave):
    """Valor liquido (ENTRADAS - SAIDAS) lancado por pedido, agrupado pela
    chave PED_MKP quando preenchido, senao DOCUMENTO (limpar_id) -- mesmo
    criterio de antes (_pedidos_lancados), so que agora soma o valor em vez
    de so marcar presenca: lancamento de frete/cobranca feito via MOV (ex.
    MKP_ME_FULL_TARIFA DE COBRANÇA) vem do TOTVS com PED_MKP vazio e o
    pedido SO no DOCUMENTO (confirmado nos pedidos 2000017834651356,
    2000018022992924, 2000018087525336 do rbx.xlsx do FULL 02.09.2026).
    DOCUMENTO passa por limpar_id igual PED_MKP (celula numerica no Excel
    vira float, senao "123.0" nunca bate com "123"). Usado pro dashboard
    mostrar o valor real lancado em cada conta ao lado da fila VERIFICAR,
    e calcular a Diferenca extrato (DIF - valor RB - valor RBX). Pedido do
    Everton, 16/09/2026.

    Filtra por OPERACAO contendo a `chave` do canal ANTES de agrupar --
    rb.xlsx/rbx.xlsx sao compartilhados entre TODOS os canais do MKP (nao so
    o FULL), entao sem esse filtro entravam junto milhares de linhas de
    outros canais/produtos (ex. rotulos "SERVICEFEE"/"AMAZONFEES" de
    Amazon), inflando a lista muito alem dos pedidos do canal atual -- bug
    real reportado pelo Everton, 18/09/2026 (guia 2 do dashboard com RBX
    gigante e RB comparativamente sumido na rolagem)."""
    df = df[_mask_chave(df["OPERACAO"], chave)]
    chave_pedido = df["PED_MKP"].where(df["PED_MKP"] != "", df["DOCUMENTO"].apply(limpar_id))
    liquido = (df["ENTRADAS"] - df["SAIDAS"]).groupby(chave_pedido).sum()
    return {k: round(v, 2) for k, v in liquido.items() if k}


def _valores_lancados_por_pedidos(df, pedidos):
    """Como _valores_lancados, mas filtra por um conjunto de PEDIDO_MKP
    especificos (isin) em vez de por chave de canal (chave_rbx) -- usado
    quando o canal ainda nao tem chave_rbx confirmada (ver conferir_pedidos).
    PED_MKP/DOCUMENTO ja identifica o pedido especifico sozinho: ao contrario
    dos agregados de fim de dia (que so tem o rotulo em OPERACAO, sem
    pedido), aqui nao precisa isolar por canal pra achar UM pedido que ja se
    sabe o numero (os da fila VERIFICAR). Pedido do Everton, 18/09/2026."""
    pedidos = {str(p).strip() for p in pedidos}
    chave_pedido = df["PED_MKP"].where(df["PED_MKP"] != "", df["DOCUMENTO"].apply(limpar_id))
    mask = chave_pedido.isin(pedidos)
    liquido = (df.loc[mask, "ENTRADAS"] - df.loc[mask, "SAIDAS"]).groupby(chave_pedido[mask]).sum()
    return {k: round(v, 2) for k, v in liquido.items() if k}


def conferir_pedidos(pasta_auxiliar, pedidos):
    """Comparativo pedido a pedido (valores_rb/valores_rbx), SEM depender do
    chave_rbx do canal -- serve pra liberar esse comparativo (semaforo +
    Diferenca extrato na fila VERIFICAR) num canal que ainda nao tem
    chave_rbx confirmada (ex. ES), sem esperar isso ser confirmado com dado
    real (a conferencia Original x Transitoria continua bloqueada, essa sim
    precisa da chave -- ver conferir()). None se rb.xlsx/rbx.xlsx nao
    existem ainda em pasta_auxiliar (mesmo criterio de conferir()). Pedido
    do Everton, 18/09/2026."""
    caminho_rb = os.path.join(pasta_auxiliar, "rb.xlsx")
    caminho_rbx = os.path.join(pasta_auxiliar, "rbx.xlsx")
    if not (os.path.exists(caminho_rb) and os.path.exists(caminho_rbx)):
        return None
    original = carregar_rb(caminho_rb)
    transitoria = carregar_rb(caminho_rbx)
    return {
        "valores_rb": _valores_lancados_por_pedidos(original, pedidos),
        "valores_rbx": _valores_lancados_por_pedidos(transitoria, pedidos),
    }


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
    # RA filtrado por VENDEDOR (ver constante "vendedor" em CANAIS e
    # docstring de _liquido_vendedor) -- sem vendedor configurado, cai no
    # comportamento antigo (qualquer RA, sem filtro).
    vendedor = cfg.get("vendedor")
    liquido_ra, qtd_ra = (
        _liquido_vendedor(transitoria, CHAVE_RA, vendedor) if vendedor
        else _liquido(transitoria, CHAVE_RA)
    )
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

    # Cancelamento pedido-a-pedido na Original (ver CHAVE_CANCELAMENTO_ORIGINAL)
    # -- pedido de dia anterior, cancelado hoje, sem entrada correspondente
    # na Transitoria de hoje. valor_transferencia e o numero real a comparar
    # com o alvo (liquido_cancelado_original ja vem negativo, entao somar
    # = descontar o excedente). Pedido do Everton, 18/09/2026.
    liquido_cancelado_original, qtd_cancelado_original, cancelados_original = (
        _cancelados_original(original, chave)
    )
    # So desconta a parte que ainda NAO tem o lancamento agregado de
    # compensacao na Transitoria hoje -- esse lancamento (quando existe) ja
    # entra em liquido_transitoria pela chave normal do canal; descontar
    # liquido_cancelado_original inteiro por cima conta o mesmo debito 2x
    # (achado real do Everton, FULL 14/09/2026, R$ 1.219,82 -- ver correcao
    # 22/09/2026 na docstring de CHAVE_CANCELAMENTO_ORIGINAL).
    liquido_cancelamento_agregado_transitoria = _cancelamento_agregado_transitoria(transitoria, chave)
    ajuste_cancelamento_nao_coberto = round(
        liquido_cancelado_original - liquido_cancelamento_agregado_transitoria, 2
    )
    valor_transferencia = round(liquido_transitoria_simulado + ajuste_cancelamento_nao_coberto, 2)
    pronta_para_transferencia = alvo is not None and abs(valor_transferencia - alvo) < TOLERANCIA

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
    # Rebate cancelado hoje (CAN-...) na RBX -- ja esta dentro de
    # liquido_rebate_rbx, so serve pra ajustar rebate_total (formula) antes
    # de comparar, senao o card mostra "nao bate" numa diferenca ja
    # explicada. Ver _rebate_cancelado_rbx.
    liquido_rebate_cancelado_rbx = _rebate_cancelado_rbx(transitoria, [chave_cupom_rebate, chave_cmm_rebate])
    rebate_bate = (
        rebate_total is not None
        and abs(liquido_rebate_rbx - round(rebate_total + liquido_rebate_cancelado_rbx, 2)) < TOLERANCIA
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
        # Cancelamentos pedido-a-pedido na Original hoje (pedido de dia
        # anterior cancelado hoje) e o valor de transferencia ja ajustado
        # por eles -- ver comentario acima de valor_transferencia. Pedido
        # do Everton, 18/09/2026.
        "liquido_cancelado_original": liquido_cancelado_original,
        "qtd_cancelado_original": qtd_cancelado_original,
        "cancelados_original": cancelados_original,
        # Parte do cancelamento que o lancamento agregado na Transitoria ja
        # cobre hoje (0 quando ainda nao foi lancado) -- ver correcao
        # 22/09/2026 acima de CHAVE_CANCELAMENTO_ORIGINAL.
        "liquido_cancelamento_agregado_transitoria": liquido_cancelamento_agregado_transitoria,
        "valor_transferencia": valor_transferencia,
        # Valor liquido ja lancado por pedido em cada conta (PED_MKP ou
        # DOCUMENTO, ver _valores_lancados) -- usado pro dashboard mostrar,
        # pedido a pedido da fila VERIFICAR, quanto ja foi lancado no RB
        # (original) e/ou no RBX (transitoria). Pedido do Everton,
        # 15-16/09/2026.
        "valores_rb": _valores_lancados(original, chave),
        "valores_rbx": _valores_lancados(transitoria, chave),
        "liquido_cupom_rebate": liquido_cupom_rebate,
        "qtd_cupom_rebate": qtd_cupom_rebate,
        "liquido_cmm_rebate": liquido_cmm_rebate,
        "qtd_cmm_rebate": qtd_cmm_rebate,
        "liquido_rebate_rbx": liquido_rebate_rbx,
        # Rebate cancelado hoje (ja dentro de liquido_rebate_rbx) -- so pra
        # mostrar/explicar a diferenca. Ver _rebate_cancelado_rbx.
        "liquido_rebate_cancelado_rbx": liquido_rebate_cancelado_rbx,
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
        for linha in linhas:
            # Vendedor e opcional (6a posicao) -- "" por padrao, so
            # preenchido quando o teste precisa dele (ver CANAIS["FULL"]
            # ["vendedor"], usado pro filtro do RA). Pedido do Everton,
            # 18/09/2026.
            operacao, entradas, saidas, documento, ped_mkp, *resto = linha
            vendedor = resto[0] if resto else ""
            ws.append(["01/09/2026", operacao, documento, "", entradas, saidas, 0, "", ped_mkp, vendedor, "", 0])
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
            # vendedor "751" (FULL, ver CANAIS) -- confirma que RA do
            # vendedor certo continua contando normalmente.
            ("CAN-EXCLUSAO DE TITULO RA", 0, 7.0, "", "", "751"),
            # Lancamento de OUTRO canal (Amazon), com pedido preenchido --
            # rb.xlsx/rbx.xlsx sao compartilhados entre todos os canais do
            # MKP. Antes do filtro por chave em _valores_lancados, isso
            # vazava pro valores_rbx (bug real reportado pelo Everton,
            # 18/09/2026: guia 2 do dashboard com milhares de linhas de
            # outros canais).
            ("CAN-MKP_AMAZO_CE REPASSE", 0, 30.0, "AMZ1", "AMZ1"),
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
        # "AMZ1" (linha de outro canal, Amazon) NAO aparece -- confirma o
        # filtro por chave em _valores_lancados.
        assert r["valores_rbx"] == {"1": 100.0, "2000017834651356": -5.0}, r
        assert "AMZ1" not in r["valores_rbx"], r
        # cancelamento de RA: fora da chave "FU" do canal, so aparece aqui
        assert r["liquido_ra"] == -7.0, r
        assert r["qtd_ra"] == 1, r

        # Cancelamento pedido-a-pedido na Original (mesma linha CAN-REC ja
        # usada acima em valores_rb, saida 10.0, pedido "1"). Esse fixture ja
        # tem o lancamento agregado de compensacao ("MKP_FULL_CANCELAMENTO")
        # nos 2 lados (Original +10 e Transitoria -10) -- o -10 da Transitoria
        # ja esta dentro de liquido_transitoria (-5.0 acima), entao NAO
        # desconta de novo: valor_transferencia = liquido_transitoria_simulado
        # sozinho (-12.0), nao -22.0 (correcao 22/09/2026 -- a versao antiga
        # deste teste tinha o mesmo bug real do Everton, FULL 14/09/2026).
        assert r["liquido_cancelado_original"] == -10.0, r
        assert r["qtd_cancelado_original"] == 1, r
        assert r["cancelados_original"] == [("1", -10.0)], r
        assert r["liquido_cancelamento_agregado_transitoria"] == -10.0, r  # ja coberto
        assert r["valor_transferencia"] == -12.0, r

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
            ("CAN-EXCLUSAO DE TITULO RA", 10.0, 0, "", "", "751"),
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
            ("CAN-EXCLUSAO DE TITULO RA", 10.0, 0, "", "", "751"),
        ])
        r3 = conferir("FULL", alvo=100.0, pasta_auxiliar=tmp3)
        assert r3["liquido_transitoria"] == -10.0, r3
        assert r3["liquido_ra"] == 10.0, r3
        assert r3["liquido_transitoria_simulado"] == 0.0, r3
        assert bool(r3["zerada"]), r3  # transferencia ja feita
        assert not bool(r3["pronta_para_transferencia"]), r3  # 0.0 != alvo (100.0)

    # Filtro de RA por VENDEDOR (cenario real, FULL 09/09/2026): RA de OUTRO
    # canal (vendedor "000423"/vazio) fica de fora, RA do FULL (vendedor
    # "751") continua contando -- Transitoria mostrava R$ 130.062,82 em vez
    # de R$ 136.068,20 (diferenca 6005.38, exatamente o RA do canal errado).
    # Pedido do Everton, 18/09/2026.
    with tempfile.TemporaryDirectory() as tmp_ra_vendedor:
        _escrever(os.path.join(tmp_ra_vendedor, "rb.xlsx"), "103925", [])
        _escrever(os.path.join(tmp_ra_vendedor, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 130062.82, 0, "1", "1"),
            # RA de outro canal (vendedor "000423") -- NAO pode contar.
            ("CAN-EXCLUSAO DE TITULO RA", 0, 6005.38, "", "", "000423"),
            # RA do FULL de verdade (vendedor "751") -- esse SIM conta.
            ("CAN-EXCLUSAO DE TITULO RA", 500.0, 0, "", "", "751"),
        ])
        r_ra = conferir("FULL", alvo=None, pasta_auxiliar=tmp_ra_vendedor)
        assert r_ra["liquido_ra"] == 500.0, r_ra  # so a linha do vendedor 751
        assert r_ra["qtd_ra"] == 1, r_ra
        assert r_ra["liquido_transitoria_simulado"] == 130562.82, r_ra  # 130062.82 + 500.0

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

    # Achado real do Everton (FULL 14/09/2026, pedido 2000018111220588,
    # R$ 136,16, aqui com 15.0): cupom recebido em outro dia, cancelado HOJE
    # ("CAN-MKP CUPOM MELI FULL REPASSE") -- o pedido nem aparece como
    # REBATE hoje (foi pra VERIFICAR, reserve_for_dispute), entao
    # rebate_total (formula) so sabe do bruto de outro pedido ("1" = 25.0),
    # nao do cancelamento. Sem a correcao, rebate_bate comparava
    # liquido_rebate_rbx (10.0, ja liquido do cancelamento) contra 25.0 e
    # dava "nao bate" (diferenca de 15.0, exatamente o cancelado).
    with tempfile.TemporaryDirectory() as tmp9c:
        _escrever(os.path.join(tmp9c, "rb.xlsx"), "103925", [])
        _escrever(os.path.join(tmp9c, "rbx.xlsx"), "X103925", [
            ("MKP CUPOM MELI FULL REPASSE", 25.0, 0, "1", "1"),
            ("CAN-MKP CUPOM MELI FULL REPASSE", 0, 15.0, "2000018111220588", "2000018111220588"),
        ])
        # rebate_total = 25 (formula, so sabe do pedido "1"); RBX liquido =
        # 25 - 15 = 10 (ja com o cancelamento). Sem a correcao, rebate_bate
        # comparava 10 contra 25 e dava "nao bate" (diferenca de 15, exatamente
        # o cancelado). Com a correcao: 10 == 25 + (-15) -> bate.
        r9c = conferir("FULL", alvo=None, pasta_auxiliar=tmp9c, rebate_total=25.0)
        assert r9c["liquido_rebate_rbx"] == 10.0, r9c
        assert r9c["liquido_rebate_cancelado_rbx"] == -15.0, r9c
        assert bool(r9c["rebate_bate"]), r9c

    # Cenario real (FULL 09/09/2026): pedido lancado (e ja transferido) num
    # dia ANTERIOR, cancelado HOJE -- so o debito residual fica na
    # Original, sem entrada correspondente na Transitoria de hoje. Isso
    # "infla" a Transitoria de hoje em relacao ao alvo; o valor de
    # transferencia real e a Transitoria menos esse cancelado, nao o bruto.
    # Pedido do Everton, 18/09/2026.
    with tempfile.TemporaryDirectory() as tmp5:
        _escrever(os.path.join(tmp5, "rb.xlsx"), "103925", [
            ("CAN-REC MKP MKP_MELI_FU REF PARC 1", 0, 20.0, "9", "9"),
        ])
        _escrever(os.path.join(tmp5, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 120.0, 0, "1", "1"),
        ])
        r5 = conferir("FULL", alvo=100.0, pasta_auxiliar=tmp5)
        assert r5["liquido_transitoria_simulado"] == 120.0, r5
        assert abs(r5["liquido_transitoria_simulado"] - 100.0) >= 0.01  # bruto NAO bate sozinho
        assert r5["liquido_cancelado_original"] == -20.0, r5
        assert r5["qtd_cancelado_original"] == 1, r5
        # sem lancamento agregado na Transitoria ainda (nenhuma linha
        # "CANCELAMENTO" no fixture) -- desconta o cancelado INTEIRO, igual
        # antes: ajustado (120.0 + -20.0 = 100.0) bate exato.
        assert r5["liquido_cancelamento_agregado_transitoria"] == 0.0, r5
        assert r5["valor_transferencia"] == 100.0, r5
        assert bool(r5["pronta_para_transferencia"]), r5

    # Achado real do Everton (FULL 14/09/2026, R$ 1.219,82): 2 CAN-REC na
    # Original (707.55 + 512.27) JA compensados hoje pelo lancamento agregado
    # "MKP_MELI_FU_CANCELAMENTOS" nos 2 lados (credito na Original, debito na
    # Transitoria) -- o debito da Transitoria ja esta em liquido_transitoria,
    # entao NAO desconta de novo: valor_transferencia = a Transitoria bruta,
    # que ja bate com o alvo (calendario "0,00" de diferenca).
    with tempfile.TemporaryDirectory() as tmp8:
        _escrever(os.path.join(tmp8, "rb.xlsx"), "103925", [
            ("CAN-REC MKP MKP_MELI_FU REF PARC 1", 0, 707.55, "1", "1"),
            ("CAN-REC MKP MKP_MELI_FU REF PARC 1", 0, 512.27, "2", "2"),
            ("MKP_MELI_FU_CANCELAMENTOS", 1219.82, 0, "121982", ""),
        ])
        _escrever(os.path.join(tmp8, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_FU REF PARC 1", 622068.27, 0, "3", "3"),
            ("MKP_MELI_FU_CANCELAMENTOS", 0, 1219.82, "121982", ""),
        ])
        r8 = conferir("FULL", alvo=620848.45, pasta_auxiliar=tmp8)
        assert r8["liquido_cancelado_original"] == -1219.82, r8
        assert r8["qtd_cancelado_original"] == 2, r8
        assert r8["liquido_cancelamento_agregado_transitoria"] == -1219.82, r8
        assert r8["liquido_transitoria_simulado"] == 620848.45, r8
        assert r8["valor_transferencia"] == 620848.45, r8  # NAO 619628.63 (desconto duplicado)
        assert bool(r8["pronta_para_transferencia"]), r8

    # conferir_pedidos: comparativo pedido a pedido sem chave_rbx (canal tipo
    # ES, ainda sem chave confirmada) -- filtra so pelos pedidos pedidos,
    # ignorando de proposito outro pedido presente no arquivo (999, de outro
    # canal/produto) que nao foi solicitado.
    with tempfile.TemporaryDirectory() as tmp6:
        _escrever(os.path.join(tmp6, "rb.xlsx"), "103925", [
            ("QUALQUER ROTULO SEM CHAVE", 0, 10.0, "1", "1"),
        ])
        _escrever(os.path.join(tmp6, "rbx.xlsx"), "X103925", [
            ("OUTRO ROTULO SEM CHAVE", 50.0, 0, "1", "1"),
            ("LANCAMENTO DE OUTRO CANAL", 999.0, 0, "999", "999"),
        ])
        r6 = conferir_pedidos(tmp6, ["1"])
        assert r6["valores_rb"] == {"1": -10.0}, r6
        assert r6["valores_rbx"] == {"1": 50.0}, r6
        assert "999" not in r6["valores_rbx"], r6  # so o pedido pedido, nao "todo mundo"
        assert conferir_pedidos(tmp6 + "_nada", ["1"]) is None  # sem arquivos

    # chave_rbx composta (lista/tupla, OR) -- cenario real do ES (18/09/2026):
    # "MKP_MELI_ES" (pedido a pedido) + "MKP_MELI_CO_" (agregado + manual
    # CMM), com uma linha de OUTRO canal (CO05, mesmos rb.xlsx/rbx.xlsx)
    # que tem que ficar de fora por causa do "_" no final da 2a chave.
    with tempfile.TemporaryDirectory() as tmp7:
        _escrever(os.path.join(tmp7, "rb.xlsx"), "103925", [
            ("CAN-REC MKP MKP_MELI_ES REF PARC 1", 0, 5.0, "1", "1"),
        ])
        _escrever(os.path.join(tmp7, "rbx.xlsx"), "X103925", [
            ("REC MKP MKP_MELI_ES REF PARC 1", 100.0, 0, "1", "1"),
            # 2a variante de pedido a pedido, sem "ES" -- confirmada pelo
            # Everton, 18/09/2026, so bate por causa de "MKP_MELI_CO REF PARC"
            # (com espaco, nao "_"), senao ficaria de fora igual o CO05.
            ("REC MKP MKP_MELI_CO REF PARC 2", 20.0, 0, "2", "2"),
            ("MKP_MELI_CO_ES_REPASSE", 0, 90.0, "", ""),
            ("MKP_MELI_CO_RECEBIMENTO DE REBATE", 8.0, 0, "", ""),
            # CO05 (outro canal, mesmos arquivos) -- NAO pode entrar: "CO05"
            # nao tem "_"/espaco logo apos "CO", entao nao bate com nenhuma
            # das chaves "MKP_MELI_CO_"/"MKP_MELI_CO REF PARC".
            ("MKP_MELI_CO05_ES_REPASSE", 0, 500.0, "", ""),
        ])
        chave_es = CANAIS["ES"]["chave_rbx"]
        r7 = conferir("ES", alvo=None, pasta_auxiliar=tmp7)  # chave_rbx real do canais.py
        assert r7["liquido_transitoria"] == round(100.0 + 20.0 - 90.0 + 8.0, 2), r7  # CO05 de fora
        assert r7["qtd_transitoria"] == 4, r7  # nao 5 (CO05 excluido)
        assert r7["liquido_original"] == -5.0, r7
        assert _mask_chave(pd.Series(["MKP_MELI_CO05_ES_REPASSE"]), chave_es).sum() == 0  # confirma a exclusao

    print("conferir_rbx.py: ok")
