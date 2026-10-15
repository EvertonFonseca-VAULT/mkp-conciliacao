"""Config por canal especifica de CONCILIACAO (raiz do historico TOTVS, onde
fica o "Para Processar" do dia, chave de busca no RB transitorio/original).
Antes cada um desses vivia num dict separado (canais.py/conferir_rbx.py/
calendario_alvo.py) e podia divergir sem ninguem notar -- foi assim que
varias das inconsistencias de rotulo desta sessao passaram batido.
Consolidado aqui em 16/09/2026, pedido do Everton (organizacional, nao
mudou nenhum valor).

Nome da loja no calendario de planejamento (loja_calendario) NAO mora aqui
-- e fato comum a KONCILI/EMPILHAR/CONCILIACAO, dono e
COMUM/calendario_alvo.LOJA_POR_CANAL (ver revisao de arquitetura, card 2).
CANAIS[canal]["loja_calendario"] continua existindo (lido de la logo
abaixo), so pra nao quebrar quem ja consome CANAIS direto.

O RB e sempre o mesmo arquivo (exportado do TOTVS pra todos os canais) --
por isso nao entra na config por canal.
"""

import glob
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "COMUM"))
from calendario_alvo import LOJA_POR_CANAL

CAMINHO_RB_PADRAO = r"C:\TEMP\Conciliador - Backup\CONCILIADOR_RB.xlsx"

CANAIS = {
    "FULL": {
        "raiz": "MKP_MELI_FU",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\FULL",
        "prefixo": "PROCESSAR_MERCADO FULL",
        "chave_rbx": "FU",  # confirmada com dado real 2x (01/09 e 02/09)
        # loja_cmm: campo LOJA da linha de Rebate colada no CMM
        # (Movimento_Bancario) -- confirmado com dado real do FULL,
        # coincide com loja_calendario (ver LOJA_POR_CANAL em COMUM/
        # calendario_alvo.py) mas e um campo diferente (ver ES, onde
        # "COLETAS ES" != "CO"). Pedido do Everton, 18/09/2026.
        "loja_cmm": "FULL",
        # prefixo_mov: usado no HISTORICO das naturezas Frete/Comissao/Rebate
        # (dashboard.py) -- confirmado com dado real do CMM_MELI_FULL
        # (rotulos MKP_ME_FULL_DESPESAS DE FRETE ENVIO etc). Pedido do
        # Everton, 17/09/2026.
        "prefixo_mov": "MKP_ME_FULL",
        # chave_cupom_rebate: operacao no RBX quando o ML repassa o
        # cupom/rebate direto (pedido "deu certo via processar", sem
        # precisar de lancamento manual no CMM) -- ver conferir_rbx.py.
        # Pedido do Everton, 18/09/2026.
        "chave_cupom_rebate": "MKP CUPOM MELI FULL REPASSE",
        # vendedor: usado pro RA (CHAVE_RA em conferir_rbx.py) -- "CAN-
        # EXCLUSAO DE TITULO RA" nao tem nenhuma chave de OPERACAO que
        # identifique o canal (ao contrario de chave_rbx="FU"), entao o
        # VENDEDOR e o unico jeito confiavel de saber se aquele RA e do
        # FULL ou de outro canal/produto. Confirmado com dado real do FULL
        # 09/09/2026: contava RA de outro canal (vendedor "000423"/vazio)
        # como se fosse do FULL. Tambem usado no campo VENDEDOR do
        # lancamento CMM (rebate_cmm.py, confirmado com o ES 18/09/2026 --
        # mesmo valor nos dois usos). Pedido do Everton, 18/09/2026.
        "vendedor": "751",
    },
    "ES": {
        "raiz": "MKP_MELI_CO_ES",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\COLETAS ES",
        "prefixo": "PROCESSAR_MERCADO COLETAS ES",
        # loja_cmm: campo LOJA da linha de Rebate colada no CMM (diferente de
        # loja_calendario, LOJA_POR_CANAL em COMUM/calendario_alvo.py -- ver
        # comentario em FULL) -- confirmado com dado real do ES (linha CMM
        # real, 18/09/2026): "...MKP MELI CO RECEBIMENTO DE REBATE...",
        # ORIGEM=MKP/PAGADOR=MELI/LOJA=CO.
        "loja_cmm": "CO",
        "chave_cupom_rebate": "MKP_MELI_CO_ES_REBATE",
        # vendedor: usado pro filtro de RA (CHAVE_RA) E pro campo VENDEDOR
        # do lancamento CMM (confirmado pelo Everton, 18/09/2026, junto com
        # a linha CMM real do ES acima -- ao contrario do que se pensava
        # pro FULL, os dois usos sao o mesmo valor).
        "vendedor": "763",
        # prefixo_mov confirmado pelo Everton, 18/09/2026, com os HISTORICO
        # reais do CMM do ES (MKP_MELI_CO_DESPESAS DE FRETE ENVIO,
        # MKP_MELI_CO_RECEBIMENTO DE REBATE, etc -- mesmo padrao
        # prefixo_mov + "_" + MOTIVO do FULL, ver dashboard._painel_historico).
        "prefixo_mov": "MKP_MELI_CO",
        # chave_rbx: confirmado pelo Everton, 18/09/2026, com rotulos reais
        # do RB/RBX do ES -- ao contrario do FULL ("FU" sozinho cobre tudo),
        # o ES nao tem 1 raiz curta comum a todos os tipos de lancamento:
        #   - pedido a pedido (Transitoria), 2 variantes de raiz confirmadas:
        #     "REC MKP MKP_MELI_ES REF PARC N" -> "MKP_MELI_ES", e tambem
        #     "REC MKP MKP_MELI_CO REF PARC N" (sem "ES") -> precisa de
        #     "MKP_MELI_CO REF PARC" (com espaco, ver abaixo). CAN-REC e so
        #     o mesmo rotulo com prefixo de cancelamento, ja coberto pelas
        #     mesmas chaves (substring).
        #   - agregado fim de dia: "MKP_MELI_CO_ES_REPASSE" /
        #     "MKP_MELI_CO_ES_CANCELAMENTOS" -> "MKP_MELI_CO_ES"
        #   - manual CMM (Frete/Comissao/Rebate/Mediacao): "MKP_MELI_CO_..."
        #     (sem "ES") -> coberto pelo mesmo "MKP_MELI_CO_" (ver abaixo)
        # "MKP_MELI_CO_ES" e um caso particular de "MKP_MELI_CO_", entao so
        # precisa dela + as 2 variantes de pedido a pedido (3 chaves no
        # total, conferir_rbx._mask_chave faz OR entre elas). O "_"/espaco
        # no final de "MKP_MELI_CO_"/"MKP_MELI_CO REF PARC" e proposital:
        # sem isso bateria tambem com "MKP_MELI_CO05_..." (canal "Mercado
        # Coletas 05", confirmado que lanca nos MESMOS rb.xlsx/rbx.xlsx do
        # ES) -- "ES" sozinho tambem foi descartado por colidir com
        # "ESTORNO"/"DESPESAS" de outros lancamentos.
        "chave_rbx": ("MKP_MELI_ES", "MKP_MELI_CO_", "MKP_MELI_CO REF PARC"),
    },
    # PRIME: raiz, rotulos do RB/RBX (MKP_MELI_PRIME_REPASSE/_COMISSAO/_FRETE),
    # HISTORICO do CMM (MKP_MELI_PRIME_...) e chave de cupom passados pelo
    # Everton, 24/09/2026. Vendedor e loja do calendario = rodar_canal/koncili.
    "PRIME": {
        "raiz": "MKP_MELI_PRIME",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\PRIME",
        "prefixo": "PROCESSAR_MERCADO PRIME",
        # 2 raizes, igual ao ES: agregado/manual "MKP_MELI_PRIME_..." e pedido a
        # pedido "REC MKP WEBCO PRIME REF PARC N" (sem MELI) -- confirmado no
        # rbx.xlsx real, 24/09/2026 (tudo vendedor 657; nao colide com MKP_B2W_PRIME).
        "chave_rbx": ("MKP_MELI_PRIME", "WEBCO PRIME REF PARC"),
        # loja_cmm e vendedor confirmados com linha CMM real (24/09/2026):
        # MKP/MELI/PRIME, TARIFA DE COBRANCA, vendedor 657, X103925.
        "loja_cmm": "PRIME",
        "chave_cupom_rebate": "MKP CUPOM MELI PRIME REPASSE",
        "vendedor": "657",
        "prefixo_mov": "MKP_MELI_PRIME",
    },
    "WEB": {
        "raiz": "MKP_MELI_WE",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\WEB",
        "prefixo": "PROCESSAR_MERCADO WEB",
        # Confirmacao em campo: vendedor 578, sem frete, layout/AB iguais aos
        # manuais do canal WEB. A raiz e a nomenclatura dos historicos do TOTVS
        # confirmadas no arquivo historicos_totvs.json (MKP_MELI_WE / ..._REPASSE
        # / ..._COMISSAO / ..._FRETE), entao o dashboard pode montar o painel
        # corretamente sem esconder a raiz.
        "chave_rbx": ("MERCADO WEB", "MKP_MELI_WE"),
        "loja_cmm": "WEB",
        "chave_cupom_rebate": "MKP CUPOM MELI WEB REPASSE",
        "vendedor": "578",
        "prefixo_mov": "MKP_MELI_WE",
    },
    # MOVEIS/ELETRO/CENTER/CO05: configs baseadas em rodar_canal.py --
    # canais "legados" que o rodar_canal.py ja processa mas o interativo.py
    # ainda nao suportava. Marcados como "a confirmar" onde nao ha dado real
    # recente -- Everton, 29/09/2026.
    "MOVEIS": {
        "raiz": "MKP_MELI_MO",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\MOVEIS",
        "prefixo": "PROCESSAR_MERCADO MOVEIS",
        # chave_rbx: a confirmar com dado real do RBX do MOVEIS.
        # Segue padrao MKP_MELI_MO_* igual FULL/WEB/PRIME.
        "chave_rbx": ("MKP_MELI_MO",),
        "loja_cmm": "MOVEIS",
        "chave_cupom_rebate": "MKP CUPOM MELI MOVEIS REPASSE",
        "vendedor": "945",  # a confirmar -- rodar_canal nao cita vendedor do MOVEIS
        "prefixo_mov": "MKP_MELI_MO",
    },
    "ELETRO": {
        "raiz": "MKP_MELI_ED",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\ELETRO",
        "prefixo": "PROCESSAR_MERCADO ELETRODOMESTICO",
        # ELETRO so TOTVS/SRI (sem PTI), sem frete -- ver rodar_canal.py.
        # chave_rbx: a confirmar com dado real.
        "chave_rbx": ("MKP_MELI_ED",),
        "loja_cmm": "ELETRODOMESTICO",
        "chave_cupom_rebate": "MKP CUPOM MELI ELETRO REPASSE",
        "vendedor": "946",  # a confirmar
        "prefixo_mov": "MKP_MELI_ED",
    },
    "CENTER": {
        "raiz": "MKP_MELI_CE",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\CENTER",
        "prefixo": "PROCESSAR_MERCADO CENTER",
        # CENTER: TOTVS/SRI + PTI, SEM frete, vendedor 944 (rodar_canal.py).
        "chave_rbx": ("MKP_MELI_CE","MKP MERCADO CENTER"),
        "loja_cmm": "CENTER",
        "chave_cupom_rebate": "MKP CUPOM MELI CENTER REPASSE",
        "vendedor": "944",
        "prefixo_mov": "MKP_MELI_CE",
    },
    "CO05": {
        "raiz": "MKP_MELI_CO_05",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\COLETAS CO05",
        "prefixo": "PROCESSAR_MERCADO COLETAS CO05",
        # CO05: TOTVS/SRI + PTI + frete, vendedor 847, mesmo layout 60 colunas
        # do ES (so ordem diferente). chave_rbx: a confirmar -- pode colidir
        # com ES (MKP_MELI_CO_*), ver comentario do ES sobre "MKP_MELI_CO_".
        "chave_rbx": ("MKP_MELI_CO_05", "MKP_MELI_CO REF PARC"),
        "loja_cmm": "CO05",
        "chave_cupom_rebate": "MKP CUPOM MELI CO05 REPASSE",
        "vendedor": "847",
        "prefixo_mov": "MKP_MELI_CO_05",
    },
}

for _canal, _cfg in CANAIS.items():
    _cfg["loja_calendario"] = LOJA_POR_CANAL.get(_canal)
del _canal, _cfg

# MOVEIS/ELETRO/CENTER/CO05/PRIME/WEB: adicionados em 29/09/2026 (ponytail
# review) -- so FULL e ES tem raiz confirmada com dado real. Os demais
# estao em "a confirmar" (chute baseado em rodar_canal.py / padrao dos
# canais confirmados) -- entrarem em CANAIS_SEM_RAIZ_CONFIRMADA pra nao
# quebrar fluxo de conferencia_rbx/RA ate validacao real.
CANAIS_SEM_RAIZ_CONFIRMADA = {"CO05"}


def validar_data_rb(data_pedida, data_ref_rb):
    """O RB e um export manual do dia -- se o analista esqueceu de exportar
    um novo antes de rodar, o arquivo antigo casa quase nada (TOTAL_RB fica
    zerado pra maioria dos pedidos) e o cruzamento vira uma fila gigante de
    falso VERIFICAR. Detecta isso ANTES de agregar/cruzar (e antes do
    scraper, que e a parte cara)."""
    if not data_ref_rb:
        raise ValueError("Nao consegui ler a data de referencia do RB (\"Dt.Ref\" em A1).")

    pedida = data_pedida.split(".")
    ref = data_ref_rb.split("/")
    if pedida[:2] != ref[:2] or (len(pedida) > 2 and pedida[2] != ref[2]):
        raise ValueError(
            f"RB e do dia {data_ref_rb}, mas voce pediu {data_pedida}. "
            "Exporte o RB do dia certo (TOTVS) antes de rodar."
        )


def resolver_para_processar(canal, data):
    """`data` no formato do nome do arquivo (dd.mm ou dd.mm.aaaa) -- acha o
    .xlsx do dia pelo prefixo do canal, ignorando o .csv que vem junto."""
    cfg = CANAIS[canal]
    padrao = f"{cfg['pasta']}\\{cfg['prefixo']} {data}*.xlsx"
    achados = sorted(glob.glob(padrao))
    if not achados:
        raise FileNotFoundError(f"Nenhum arquivo casou com: {padrao}")
    if len(achados) > 1:
        raise FileNotFoundError(f"Mais de um arquivo casou com {padrao}: {achados}")
    return achados[0]
