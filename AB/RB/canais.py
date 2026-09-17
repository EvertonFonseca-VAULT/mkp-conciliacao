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
    },
}

# MOVEIS/ES/ELETRO removidos daqui (18/09/2026, pedido do Everton): por
# enquanto o projeto roda so FULL, nao e pra manter os outros canais
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
