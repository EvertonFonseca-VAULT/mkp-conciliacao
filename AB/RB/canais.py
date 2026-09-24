"""Config por canal -- FONTE UNICA de tudo que varia por canal (raiz do
historico TOTVS, onde fica o "Para Processar" do dia, chave de busca no RB
transitorio/original, nome da loja no calendario de planejamento). Antes
cada um desses vivia num dict separado (canais.py/conferir_rbx.py/
calendario_alvo.py) e podia divergir sem ninguem notar -- foi assim que
varias das inconsistencias de rotulo desta sessao passaram batido.
Consolidado aqui em 16/09/2026, pedido do Everton (organizacional, nao
mudou nenhum valor).

O RB e sempre o mesmo arquivo (exportado do TOTVS pra todos os canais) --
por isso nao entra na config por canal.
"""

import glob

CAMINHO_RB_PADRAO = r"C:\TEMP\Conciliador - Backup\CONCILIADOR_RB.xlsx"

CANAIS = {
    "FULL": {
        "raiz": "MKP_MELI_FU",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\FULL",
        "prefixo": "PROCESSAR_MERCADO FULL",
        "chave_rbx": "FU",  # confirmada com dado real 2x (01/09 e 02/09)
        "loja_calendario": "FULL",
        # loja_cmm: campo LOJA da linha de Rebate colada no CMM
        # (Movimento_Bancario) -- confirmado com dado real do FULL,
        # coincide com loja_calendario mas e um campo diferente (ver ES,
        # onde "COLETAS ES" != "CO"). Pedido do Everton, 18/09/2026.
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
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\COLETAS ES",
        "prefixo": "PROCESSAR_MERCADO COLETAS ES",
        # loja_calendario: chute pelo nome do canal (sem dado real do
        # calendario de planejamento pra confirmar) -- se estiver errado
        # buscar_alvo() so retorna None (sem alvo pra comparar), nao quebra
        # nada nem gera numero errado (ver calendario_alvo.py:56-58).
        "loja_calendario": "COLETAS ES",
        # loja_cmm: campo LOJA da linha de Rebate colada no CMM (diferente
        # de loja_calendario -- ver comentario em FULL) -- confirmado com
        # dado real do ES (linha CMM real, 18/09/2026): "...MKP MELI CO
        # RECEBIMENTO DE REBATE...", ORIGEM=MKP/PAGADOR=MELI/LOJA=CO.
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
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\PRIME",
        "prefixo": "PROCESSAR_MERCADO PRIME",
        "chave_rbx": "MKP_MELI_PRIME",  # nao colide com MKP_B2W_PRIME (outro marketplace)
        "loja_calendario": "PRIME",
        # loja_cmm e vendedor confirmados com linha CMM real (24/09/2026):
        # MKP/MELI/PRIME, TARIFA DE COBRANCA, vendedor 657, X103925.
        "loja_cmm": "PRIME",
        "chave_cupom_rebate": "MKP CUPOM MELI PRIME REPASSE",
        "vendedor": "657",
        "prefixo_mov": "MKP_MELI_PRIME",
    },
}

# MOVEIS/ELETRO removidos daqui (18/09/2026, pedido do Everton): por
# enquanto o projeto roda so FULL/ES, nao e pra manter os outros canais
# meio-configurados (config nunca confirmada com dado real deles). Quando
# for a hora de replicar pra outro canal, volta no historico do git (commit
# 142ffb8 em diante) pra recuperar os dicts antigos como ponto de partida.
CANAIS_SEM_RAIZ_CONFIRMADA = set()


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
