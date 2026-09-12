"""Tradução em Python das macros do `PB BASE.xlsm` (FULL) — ver
`Modulo1_Para_Processar.bas` e `Modulo2_Para_conciliar.bas` em
`Manuais/processo conciliação Mercado FULL/VBA/`, e o passo a passo em
`Manuais/processo conciliação Mercado FULL/Passo a passo - Conciliação
Mercado FULL.md`.

Cobre a parte confirmada pelas fórmulas reais da VBA:
  - marcar_flags / filtrar_validos    (Para_Processar: colunas T/U)
  - carregar_totvs / cruzar_totvs      (Para_Processar: XLOOKUP duplo)
  - classificar_diferenca              (Para_conciliar: fórmula M)

NÃO cobre ainda:
  - leitura do RB (extrato bancário) e o SUMIFS por PED MKP.
"""

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


def cruzar_totvs(df, totvs, coluna_chave=COLUNA_CHAVE_TOTVS):
    """Equivalente ao XLOOKUP duplo da coluna Q (recontando os offsets
    RC[] direto na fórmula real — corrige uma suposição anterior errada:
    a chave primária é EXTERNAL_REFERENCE, não SOURCE_ID). Sem nenhum dos
    dois, `verificar=True` (cai na fila de conferência manual / SRI).

    ⚠️ A fórmula original usa como 2º critério o resultado de uma cadeia
    de XLOOKUP em 'Aux2' que parece cruzar EXTERNAL_REFERENCE contra a
    coluna SHIPPING_ID (provável resíduo/bug de planilha, os tipos de ID
    não batem). Aqui uso o SHIPPING_ID bruto como 2º critério — mais
    parecido com o processo manual real (SRI por SHIPPING_ID, depois
    EXTERNAL_REFERENCE) do que com essa cadeia suspeita. A confirmar com
    o Everton antes de tratar isso como definitivo."""
    df = df.copy()
    chaves_totvs = set(totvs[coluna_chave].astype(str))
    referencia = df["EXTERNAL_REFERENCE"].astype(str)
    shipping_id = df["SHIPPING_ID"].astype(str)
    achou_direto = referencia.isin(chaves_totvs)
    achou_fallback = shipping_id.isin(chaves_totvs)
    df["totvs_encontrado"] = referencia.where(achou_direto, shipping_id.where(achou_fallback, 0))
    df["verificar"] = ~(achou_direto | achou_fallback)
    return df


CAMINHO_SCRAPER_SRI = r"C:\Users\everton.fonseca\Desktop\Projetos\Screaper Mercado"


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
        resolvido[(external_ref, shipping_id)] = shipping_id if dados else external_ref

    pendentes = pendentes.copy()
    chave = list(zip(pendentes["EXTERNAL_REFERENCE"], pendentes["SHIPPING_ID"]))
    pendentes["totvs_encontrado"] = [resolvido[k] for k in chave]
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


def classificar_diferenca(diferenca, rebate, mediation):
    """M = classificação Sem Diferença/Rebate/Mediation/verificar do
    Para_conciliar, mesma ordem de checagem da fórmula real."""
    if round(diferenca, 2) == 0:
        return "Sem Diferença"
    if round(diferenca - rebate, 2) == 0:
        return "Rebate"
    if round(diferenca - mediation, 2) == 0:
        return "Mediation"
    return "verificar"


def _demo():
    linhas = pd.DataFrame({
        "EXTERNAL_REFERENCE": ["1", "2", "3", "4"],
        "SHIPPING_ID": ["10", "20", "30", "40"],
        "RECORD_TYPE": ["release", "release", "release", "release"],
        "DESCRIPTION": ["payment", "Saque de dinheiro", "refund", "PAYOUT"],
    })
    flags = marcar_flags(linhas)
    assert list(flags["movimento_valido"]) == ["SIM", "NÃO", "SIM", "NÃO"]
    assert len(filtrar_validos(flags)) == 2

    totvs = pd.DataFrame({"chave": ["1", "30"]})
    cruzado = cruzar_totvs(linhas, totvs, "chave")
    assert list(cruzado["verificar"]) == [False, True, False, True]

    assert classificar_diferenca(0, 10, 5) == "Sem Diferença"
    assert classificar_diferenca(10, 10, 5) == "Rebate"
    assert classificar_diferenca(5, 10, 5) == "Mediation"
    assert classificar_diferenca(3, 10, 5) == "verificar"
    print("ok")


if __name__ == "__main__":
    _demo()
