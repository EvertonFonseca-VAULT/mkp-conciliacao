"""Config por canal: RAIZ do historico TOTVS e onde fica o "Para Processar"
do dia. O RB e sempre o mesmo arquivo (exportado do TOTVS pra todos os
canais) -- por isso nao entra na config por canal."""

import glob

CAMINHO_RB_PADRAO = r"C:\TEMP\Conciliador - Backup\CONCILIADOR_RB.xlsx"

CANAIS = {
    "MOVEIS": {
        "raiz": "MKP_MELI_MO",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\MOVEIS",
        "prefixo": "PROCESSAR_MERCADO MOVEIS",
    },
    "FULL": {
        "raiz": "MKP_MELI_FU",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\FULL",
        "prefixo": "PROCESSAR_MERCADO FULL",
    },
    "ES": {
        "raiz": "MKP_MELI_CO",
        "pasta": r"C:\RODAR TOTVS\PROCESSAR\COLETAS ES",
        "prefixo": "PROCESSAR_MERCADO COLETAS ES",
    },
    "ELETRO": {
        "raiz": "MKP_MELI_ED",
        "pasta": r"C:\RODAR TOTVS\ELETRODOMESTICO",
        "prefixo": "MERCADO ELETRODOMESTICO",
    },
}

# ES: raiz no canais.py (MKP_MELI_CO) diverge da tabela HISTORICOS do
# CONTROLE_V2.xlsx (MKP_MELI_CO_ES) -- ainda nao confirmado qual esta certo
# (15/09/2026, pedido do Everton). Ate confirmar, o painel "Historico TOTVS"
# do dashboard nao mostra nada pra ES -- ver interativo.py/executar().
CANAIS_SEM_RAIZ_CONFIRMADA = {"ES"}


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
