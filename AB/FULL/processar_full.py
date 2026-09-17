"""Tradução em Python das macros do `PB BASE.xlsm` (FULL) — ver
`Modulo1_Para_Processar.bas` e `Modulo2_Para_conciliar.bas` em
`Manuais/processo conciliação Mercado FULL/VBA/`, e o passo a passo em
`Manuais/processo conciliação Mercado FULL/Passo a passo - Conciliação
Mercado FULL.md`.

Cobre a parte confirmada pelas fórmulas reais da VBA:
  - marcar_flags / filtrar_validos    (Para_Processar: colunas T/U)
  - carregar_totvs / cruzar_totvs      (Para_Processar: XLOOKUP duplo)

NÃO cobre ainda:
  - leitura do RB (extrato bancário) e o SUMIFS por PED MKP.
"""

import os

import numpy as np
import pandas as pd

from tipagem_full import LAYOUT_FULL

# Valores de DESCRIPTION que indicam saque/liquidação da conta MP (não é
# movimento por pedido) — mesma lista da coluna T da VBA. A coluna U
# (Liberação, RECORD_TYPE == "Liberações") é calculada só como informativa:
# na VBA real, o AutoFilter da Liberação (Field:=21) é chamado sem
# Criteria1, então não filtra nada de fato — só o filtro de T (Field:=20)
# é aplicado.
DESCRICOES_MOVIMENTO_INVALIDO = {
    "RESERVE_FOR_PAYOUT", "PAYOUT", "Reserva para pagamento", "Saque de dinheiro",
}


def carregar_ab(caminho):
    df = pd.read_excel(caminho, sheet_name="Relatorio")
    if list(df.columns) != LAYOUT_FULL:
        raise ValueError(f"Layout do Mercado FULL mudou. Esperado: {LAYOUT_FULL}\nEncontrado: {list(df.columns)}")
    return df


def marcar_flags(df):
    df = df.copy()
    df["movimento_valido"] = df["DESCRIPTION"].where(
        ~df["DESCRIPTION"].isin(DESCRICOES_MOVIMENTO_INVALIDO), "NÃO"
    )
    df["movimento_valido"] = df["movimento_valido"].where(df["movimento_valido"] == "NÃO", "SIM")
    df["liberacao"] = df["RECORD_TYPE"].apply(lambda v: "SIM" if v == "Liberações" else "NÃO")
    return df


def filtrar_validos(df):
    return df[df["movimento_valido"] == "SIM"].reset_index(drop=True)


# Confirmado abrindo o arquivo real: a fórmula referencia "coluna 13" da
# aba, que é a coluna com cabeçalho "Nro MKP" (a linha 1 é só um numerador
# de colunas 1..18, a linha 5 é o cabeçalho de verdade).
#
# ⚠️ Tem duas cópias com o mesmo nome: a do X: (rede) é um resíduo antigo
# (relatório de 2024, 0% de match). A boa é a LOCAL abaixo — mesmo caminho
# referenciado dentro da própria fórmula VBA — atualizada e com ~89% de
# match nos testes reais (262/294 linhas do FULL 01.09.2026).
CAMINHO_TOTVS_PADRAO = r"C:\KIT CONCILIAÇÃO MKP\TOTVS_2022.xlsx"
COLUNA_CHAVE_TOTVS = "Nro MKP"


def carregar_totvs(caminho=CAMINHO_TOTVS_PADRAO):
    """Lê a aba '01-01 - Listagem do Browse': linhas 1-4 são cabeçalho de
    relatório/numerador de colunas (não são dados), o cabeçalho real está
    na linha 5. Arquivo grande (~387 mil linhas) — carregar uma vez só."""
    return pd.read_excel(caminho, sheet_name="01-01 - Listagem do Browse", skiprows=4)


def _limpar_id(valor):
    """Formata um ID (SHIPPING_ID/EXTERNAL_REFERENCE) sem o '.0' que o
    pandas cola em colunas numéricas com valor ausente (float64)."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def cruzar_totvs(df, totvs, coluna_chave=COLUNA_CHAVE_TOTVS):
    """Chave de cruzamento é sempre o SHIPPING_ID (confirmado pelo Everton:
    é assim que ele pesquisa no SRI manualmente). Quando não tem SHIPPING_ID
    (ex.: linhas de cashback, que não têm envio associado), usa o
    EXTERNAL_REFERENCE do próprio AB direto, sem validar contra o TOTVS —
    mesma regra que ele aplica manualmente nesse caso.

    ⚠️ Não valida mais o EXTERNAL_REFERENCE cru contra o TOTVS quando existe
    SHIPPING_ID: casar por EXTERNAL_REFERENCE causava falso positivo (o
    número do AB podia coincidir com um título real de OUTRO pedido/cliente
    — achado real no lote de 03.09.2026). Se o SHIPPING_ID existir mas não
    bater direto no TOTVS, cai em `verificar=True` pra resolver via SRI
    (busca por SHIPPING_ID, mesmo fluxo manual)."""
    df = df.copy()
    chaves_totvs = set(totvs[coluna_chave].dropna().astype(str)) - {"", "nan"}

    referencia = df["EXTERNAL_REFERENCE"].apply(_limpar_id)
    shipping_id = df["SHIPPING_ID"].apply(_limpar_id)
    tem_shipping_id = shipping_id != ""

    achou_via_shipping = tem_shipping_id & shipping_id.isin(chaves_totvs)

    totvs_encontrado = np.where(achou_via_shipping, shipping_id, np.where(~tem_shipping_id, referencia, "0"))
    df["totvs_encontrado"] = pd.Series(totvs_encontrado, index=df.index, dtype=object)
    df["verificar"] = ~(achou_via_shipping | ~tem_shipping_id)
    return df


# %USERPROFILE%, nao caminho fixo com usuario -- pra rodar na maquina de
# qualquer colega que tenha a pasta "Screaper Mercado" clonada no mesmo
# lugar (Desktop\Projetos) na PROPRIA conta Windows. Pedido do Everton,
# 18/09/2026 (compartilhar o processo de processar com o time).
CAMINHO_SCRAPER_SRI = os.path.join(os.path.expanduser("~"), "Desktop", "Projetos", "Screaper Mercado")


def resolver_pendentes_via_sri(pendentes):
    """Para as linhas 'Verificar' (sem match automático no TOTVS): a
    maioria repete (mesmo pedido, várias linhas de release do MP com o
    mesmo EXTERNAL_REFERENCE/SHIPPING_ID) — dedupe antes de bater no SRI,
    cada pedido só precisa ser pesquisado uma vez. Reaproveita o scraper
    já existente em `Screaper Mercado/Pagina.py` (login + busca por
    pedmkp): achou no SRI -> usa SHIPPING_ID; não achou -> EXTERNAL_REFERENCE
    (mesma regra manual documentada no passo a passo)."""
    import sys
    if CAMINHO_SCRAPER_SRI not in sys.path:
        sys.path.insert(0, CAMINHO_SCRAPER_SRI)
    from Pagina import login, buscar_e_validar, USUARIO, SENHA
    import requests

    unicos = pendentes[["EXTERNAL_REFERENCE", "SHIPPING_ID"]].drop_duplicates()

    session = requests.Session()
    login(session, USUARIO, SENHA)

    resolvido = {}
    for external_ref, shipping_id in unicos.itertuples(index=False):
        dados, _ = buscar_e_validar(session, shipping_id)
        resolvido[(external_ref, shipping_id)] = _limpar_id(shipping_id) if dados else _limpar_id(external_ref)

    pendentes = pendentes.copy()
    chave = list(zip(pendentes["EXTERNAL_REFERENCE"], pendentes["SHIPPING_ID"]))
    pendentes["totvs_encontrado"] = pd.Series([resolvido[k] for k in chave], index=pendentes.index, dtype=object)
    pendentes["verificar"] = False
    return pendentes


# Confirmado comparando com o arquivo real gerado pela macro
# (`C:\RODAR TOTVS\MERCADO FULL 01.09.xlsx`, aba "Para Processar"): o
# Para_processar2 reescreve EXTERNAL_REFERENCE com o valor resolvido do
# TOTVS (`totvs_encontrado`) antes de salvar — por isso, nas linhas que só
# acharam match via SHIPPING_ID (fallback), o EXTERNAL_REFERENCE final não
# é mais o original, e sim o SHIPPING_ID. Só as 15 colunas abaixo (A1:O1
# na VBA) vão pro arquivo final — SHIPPING_ID/SHIPMENT_MODE/ORDER_ID/PACK_ID
# e as colunas de conferência (Q/R/S/verificar) ficam só na etapa de QA.
COLUNAS_PARA_PROCESSAR = [
    "RELEASE_DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "RECORD_TYPE", "DESCRIPTION",
    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "MP_FEE_AMOUNT",
    "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT",
    "INSTALLMENTS", "PAYMENT_METHOD",
]


def gerar_para_processar(cruzado):
    """Monta a aba final 'Para Processar' (equivalente ao que
    Para_processar2 salva em `Capa!S1`/`Capa!B3`): EXTERNAL_REFERENCE
    substituído pelo valor resolvido no TOTVS, só as 15 colunas finais."""
    df = cruzado.copy()
    df["EXTERNAL_REFERENCE"] = df["totvs_encontrado"]
    df = df[COLUNAS_PARA_PROCESSAR].reset_index(drop=True)
    df.columns = ["DATE"] + COLUNAS_PARA_PROCESSAR[1:]
    return df
