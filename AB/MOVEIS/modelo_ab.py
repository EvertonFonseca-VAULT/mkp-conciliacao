"""Tradução em Python das fórmulas do `MODELO AB MERCADO.xlsx` (canal
MOVEIS, PTI código 000755) — ver `Manuais/MODELO AB MERCADO/MODELO AB
MERCADO - guias.md`.

Fórmulas extraídas direto do arquivo real (`C:\\TEMP\\MERCADO\\MODELO AB
MERCADO.xlsx`, linha 2 de cada aba) em 11/09/2026 — não da documentação
manual, pra evitar erro de transcrição/contagem de coluna.

Cobre:
  - conferencia_ab_original      (AB Original: AP:AS, só QA por linha)
  - calcular_ab_para_processar    (AB Para Processar: P:U)
  - montar_validacao              (Validação: agregação + classificação + U:AF)
  - resumo_acao_manual            (Resumo: fila de trabalho manual)

⚠️ Achados/decisões ao traduzir (ver conversa/checklist no .md):
  - PTI!AC ("EMISSÃO PED.", uma data) é o que a fórmula real da coluna
    "FRETE" usa — bug de planilha (a coluna certa é PTI!AD). Aqui já
    CORRIGIDO pra usar AD/FRETE de verdade (decisão do Everton).
  - RB!C:C (DOCUMENTO) é a chave usada no cruzamento RB da Validação —
    diferente do Para_conciliar do FULL, que usa RB!I:I (PED MKP). Não
    deu pra validar com dado real (a aba RB deste arquivo-modelo só tem
    4 linhas de exemplo, nenhuma com a operação usada no filtro AG$1).
  - AG$1 (RB!B:B) = o texto literal "REC MKP MKP_MELI_ES REF PARC 1"
    (é o próprio cabeçalho da coluna AG da Validação, usado como
    constante de filtro).

NÃO cobre ainda: `ProvaReal para Processar` (coluna AB da Validação) —
depende da guia `Para Processar` (outro recorte, diferente de `AB Para
Processar`), cujas colunas F/G ainda não foram identificadas.
"""

import numpy as np
import pandas as pd

TOLERANCIA = 0.05
FILTRO_RB_PADRAO = "REC MKP MKP_MELI_ES REF PARC 1"

CATEGORIAS_SEM_AJUSTE = {"Nulo", "Estorno", "AB valor liquido", "Verificar", "Manter igual", "Manter igual, dif é frete"}
CATEGORIAS_CUPOM = {"Deduzir Cupom", "Deduzir Cupom CMM"}
CATEGORIAS_TAXA_CMM = {"Deduzir Taxa financiamento CMM", "Deduzir Taxa e Cupom"}
CATEGORIAS_ACAO_MANUAL = {"AB valor liquido", "Estorno", "Verificar"}


def conferencia_ab_original(df):
    """AB Original!AP:AS — só conferência por linha, não alimenta o resto."""
    df = df.copy()
    df["Crédito menos Débito"] = df["VALOR LÍQUIDO CREDITADO"] - df["VALOR LÍQUIDO DEBITADO"]
    df["Composição repasse"] = (
        df["VALOR BRUTO DA OPERAÇÃO"]
        + df["TARIFA DO MERCADO PAGO OU DO MERCADO LIVRE"]
        + df["TARIFA POR OFERECER PARCELAMENTO SEM JUROS"]
        + df["CUSTO DE ENVIO"]
    )
    df["Diferença"] = df["Crédito menos Débito"] - df["Composição repasse"]
    df["Considerar"] = "Sim"
    return df


def calcular_ab_para_processar(pb, ab_original):
    """`pb` = saída do PB (Para_Processar/PARA_PROCESSAR2) do MOVEIS, 15
    colunas em inglês. `ab_original` precisa estar na MESMA ordem de
    linhas que `pb` — mesmo pressuposto manual da planilha real (cola-se
    os dois lado a lado, sem reordenar nenhum dos dois)."""
    if len(pb) != len(ab_original):
        raise ValueError("AB Para Processar e AB Original têm número de linhas diferente — conferir alinhamento.")
    pb = pb.reset_index(drop=True)
    ab_original = ab_original.reset_index(drop=True)

    df = pb.copy()
    df["VALIDACAO"] = df["SOURCE_ID"].astype(str) == ab_original["ID DA OPERAÇÃO NO MERCADO PAGO"].astype(str)
    df["CUSTO DE OFERECER DESCONTO"] = pd.to_numeric(ab_original["CUSTO DE OFERECER DESCONTO"], errors="coerce")
    soma_hk = df[["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT"]].sum(axis=1)
    df["NOVO CREDITO"] = soma_hk.where(soma_hk >= 0, 0)
    df["NOVO DEBITO"] = -soma_hk.where(soma_hk < 0, 0)
    df["Saldo Repasse"] = df["NOVO CREDITO"] - df["NOVO DEBITO"]
    df["ARRUMAR"] = df["EXTERNAL_REFERENCE"].astype(str).str.strip()
    return df


def _classificar(row):
    # Arredonda pra centavo antes de comparar: ruído de ponto flutuante do
    # SUM/groupby (ex.: -2.27e-13 em vez de 0) senão vaza pra "Estorno"/
    # "Verificar" em pedidos que deveriam ser "Nulo"/"Deduzir Taxa e Cupom".
    saldo = round(row["saldo_recebimento"], 2)
    tem_comissao = row["tem_comissao"]
    dif = round(row["dif_gross_pti"], 2)
    coupon = round(row["coupon"], 2)
    frete = round(row["frete"], 2)
    fin_fee = round(row["financing_fee"], 2)
    saldo_cupom = round(row["saldo_cupom"], 2)

    if saldo == 0:
        return "Nulo"
    if saldo < 0:
        return "Estorno"
    if not tem_comissao:
        return "AB valor liquido"
    if -TOLERANCIA <= dif <= TOLERANCIA and coupon == 0:
        return "Manter igual"
    if -TOLERANCIA <= dif <= TOLERANCIA and coupon != 0:
        return "Deduzir Cupom CMM"
    if -TOLERANCIA <= dif - coupon <= TOLERANCIA:
        return "Deduzir Cupom"
    if -dif == frete:
        return "Manter igual, dif é frete"
    if -TOLERANCIA <= dif + fin_fee <= TOLERANCIA and coupon == 0:
        return "Deduzir Taxa financiamento"
    if -TOLERANCIA <= dif + fin_fee <= TOLERANCIA and coupon != 0:
        return "Deduzir Taxa financiamento CMM"
    if dif == -fin_fee + coupon:
        return "Deduzir Taxa e Cupom"
    if -TOLERANCIA <= dif - saldo_cupom <= TOLERANCIA:
        return "Cupom compartilhado"
    return "Verificar"


def _recalcular(row):
    cat, e, f, g, h, i = row["categoria"], row["gross_amount"], row["mp_fee"], row["financing_fee"], row["shipping_fee"], row["coupon"]

    if cat in CATEGORIAS_SEM_AJUSTE:
        gross = e
    elif cat in CATEGORIAS_CUPOM or cat == "Cupom compartilhado":
        gross = e - i
    elif cat == "Deduzir Taxa financiamento":
        gross = e + g
    elif cat in CATEGORIAS_TAXA_CMM:
        gross = e + g - i
    else:
        gross = e

    if cat in {"Deduzir Taxa financiamento", *CATEGORIAS_TAXA_CMM}:
        comissao = f
    else:
        comissao = f + g

    cupom_cmm = i if cat in {*CATEGORIAS_CUPOM, "Deduzir Taxa e Cupom", "Deduzir Taxa financiamento CMM", "Cupom compartilhado"} else 0

    total = gross + comissao + h + cupom_cmm
    return pd.Series({
        "novo_gross_amount": gross,
        "nova_comissao": comissao,
        "custo_frete": h,
        "cmm_cupom": cupom_cmm,
        "total_repasse_recalculado": total,
    })


def montar_validacao(ab_pp, pti, rb=None, filtro_rb=FILTRO_RB_PADRAO):
    ab_pp = ab_pp.copy()
    ab_pp["ARRUMAR"] = ab_pp["ARRUMAR"].astype(str).str.strip()
    pti = pti.copy()
    pti["PED. MKP"] = pti["PED. MKP"].astype(str).str.strip()

    agg = ab_pp.groupby("ARRUMAR").agg(
        qtd_ab=("ARRUMAR", "size"),
        net_credit=("NOVO CREDITO", "sum"),
        net_debit=("NOVO DEBITO", "sum"),
        gross_amount=("GROSS_AMOUNT", "sum"),
        mp_fee=("MP_FEE_AMOUNT", "sum"),
        financing_fee=("FINANCING_FEE_AMOUNT", "sum"),
        shipping_fee=("SHIPPING_FEE_AMOUNT", "sum"),
        coupon=("COUPON_AMOUNT", "sum"),
        custo_desconto=("CUSTO DE OFERECER DESCONTO", "sum"),
    ).reset_index().rename(columns={"ARRUMAR": "PED_MKP"})

    agg["saldo_recebimento"] = agg["net_credit"] - agg["net_debit"]
    agg["tem_comissao"] = np.round(agg["mp_fee"] + agg["financing_fee"], 2) != 0
    agg["saldo_cupom"] = agg["coupon"] + agg["custo_desconto"]

    # PTI!FRETE = coluna AD (corrigido — a fórmula real da planilha usa AC,
    # "EMISSÃO PED.", por engano; ver docstring do módulo).
    pti_agg = pti.groupby("PED. MKP").agg(
        qtd_pti=("PED. MKP", "size"),
        valor_reais=("VALOR REAIS", "sum"),
        valor=("VALOR", "sum"),
        frete=("FRETE", "sum"),
    ).reset_index().rename(columns={"PED. MKP": "PED_MKP"})

    df = agg.merge(pti_agg, on="PED_MKP", how="left")
    for col in ["qtd_pti", "valor_reais", "valor", "frete"]:
        df[col] = df[col].fillna(0)

    df["validar_gross"] = (df["saldo_recebimento"] > 0) & df["tem_comissao"]
    df["dif_gross_pti"] = df["gross_amount"] - df["valor"]

    df["categoria"] = df.apply(_classificar, axis=1)
    df = pd.concat([df, df.apply(_recalcular, axis=1)], axis=1)
    df["diferenca_calculada"] = df["saldo_recebimento"] - df["total_repasse_recalculado"]
    df["lancar_frete"] = df["custo_frete"] < 0
    df["saldo_pti_final"] = np.where(
        df["categoria"].isin({"Nulo", "Estorno", "AB valor liquido", "Verificar"}),
        0, df["valor"] - df["novo_gross_amount"],
    )

    if rb is not None:
        rb = rb.copy()
        rb["DOCUMENTO"] = rb["DOCUMENTO"].astype(str).str.strip()
        rb_filtrado = rb[rb["OPERACAO"] == filtro_rb]
        rb_agg = rb_filtrado.groupby("DOCUMENTO").agg(
            entradas=("ENTRADAS", "sum"), saidas=("SAIDAS", "sum"), saldo_titulo=("Saldo Titulo", "sum"),
        ).reset_index().rename(columns={"DOCUMENTO": "PED_MKP"})
        df = df.merge(rb_agg, on="PED_MKP", how="left")
        df[["entradas", "saidas", "saldo_titulo"]] = df[["entradas", "saidas", "saldo_titulo"]].fillna(0)
        df["valor_rb"] = df["entradas"] - df["saidas"]
        df["diferenca_final"] = df["total_repasse_recalculado"] - df["valor_rb"]

    return df


COLUNAS_PAGAMENTO = [
    "DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "RECORD_TYPE", "DESCRIPTION",
    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "MP_FEE_AMOUNT",
    "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT",
    "INSTALLMENTS", "PAYMENT_METHOD",
]


def montar_pagamento(validacao, ab_pp):
    """Guia Pagamento: uma linha por pedido — coluna C é a MESMA lista de
    PED_MKP colada da Validação!A (confirmado pelo Everton), com PROCX
    (primeira ocorrência) em 'AB Para Processar' trazendo DATE/SOURCE_ID/
    RECORD_TYPE/DESCRIPTION/INSTALLMENTS/PAYMENT_METHOD. GROSS_AMOUNT/
    MP_FEE_AMOUNT/COUPON_AMOUNT vêm já **recalculados** da Validação (U/V/X);
    FINANCING_FEE/SHIPPING_FEE/TAXES ficam zerados (já absorvidos no
    recálculo)."""
    ab_pp = ab_pp.copy()
    ab_pp["_chave"] = ab_pp["EXTERNAL_REFERENCE"].astype(str).str.strip()
    primeira = ab_pp.drop_duplicates("_chave", keep="first").set_index("_chave")

    pedidos = validacao["PED_MKP"]
    lookup = primeira.reindex(pedidos)

    df = pd.DataFrame({
        "DATE": lookup["DATE"].values,
        "SOURCE_ID": lookup["SOURCE_ID"].values,
        "EXTERNAL_REFERENCE": pedidos.values,
        "RECORD_TYPE": lookup["RECORD_TYPE"].values,
        "DESCRIPTION": lookup["DESCRIPTION"].values,
        "GROSS_AMOUNT": validacao["novo_gross_amount"].values,
        "MP_FEE_AMOUNT": validacao["nova_comissao"].values,
        "FINANCING_FEE_AMOUNT": 0,
        "SHIPPING_FEE_AMOUNT": 0,
        "TAXES_AMOUNT": 0,
        "COUPON_AMOUNT": validacao["cmm_cupom"].values,
        "INSTALLMENTS": lookup["INSTALLMENTS"].values,
        "PAYMENT_METHOD": lookup["PAYMENT_METHOD"].values,
    })
    soma_hk = df[["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT"]].sum(axis=1)
    df["NET_CREDIT_AMOUNT"] = soma_hk.where(soma_hk >= 0, 0)
    df["NET_DEBIT_AMOUNT"] = -soma_hk.where(soma_hk < 0, 0)
    return df[COLUNAS_PAGAMENTO]


def montar_frete(validacao, ab_pp, pedidos=None):
    """Guia Frete: mesmo esqueleto do Pagamento (PROCX em 'AB Para
    Processar' pela primeira ocorrência do pedido), mas:
      - `DESCRIPTION` fixo em "Frete" (não vem de XLOOKUP);
      - só `SHIPPING_FEE_AMOUNT` é preenchido, com
        `SOMASE(Validação!A:A;pedido;Validação!W:W)` (Custo Frete);
        GROSS/MP_FEE/TAXES/COUPON ficam zerados.
    `pedidos`: lista de PED_MKP a colar na coluna C — como essa colagem é
    manual e pode repetir pedido (a guia real acumula lotes antigos), por
    padrão usa só os pedidos com `Lançar Frete`="Sim" do lote atual
    (✅ confirmado pelo Everton: "na guia frete eu filtro apenas o que ta
    marcado como LANÇAR FRETE"), mas aceita uma lista própria."""
    if pedidos is None:
        pedidos = validacao.loc[validacao["lancar_frete"], "PED_MKP"]
    pedidos = pd.Series(pedidos).reset_index(drop=True)

    ab_pp = ab_pp.copy()
    ab_pp["_chave"] = ab_pp["EXTERNAL_REFERENCE"].astype(str).str.strip()
    primeira = ab_pp.drop_duplicates("_chave", keep="first").set_index("_chave")
    lookup = primeira.reindex(pedidos)

    custo_frete = validacao.set_index("PED_MKP")["custo_frete"]
    soma_frete = pedidos.map(lambda p: custo_frete.get(p, 0))

    df = pd.DataFrame({
        "DATE": lookup["DATE"].values,
        "SOURCE_ID": lookup["SOURCE_ID"].values,
        "EXTERNAL_REFERENCE": pedidos.values,
        "RECORD_TYPE": lookup["RECORD_TYPE"].values,
        "DESCRIPTION": "Frete",
        "GROSS_AMOUNT": 0,
        "MP_FEE_AMOUNT": 0,
        "FINANCING_FEE_AMOUNT": 0,
        "SHIPPING_FEE_AMOUNT": soma_frete.values,
        "TAXES_AMOUNT": 0,
        "COUPON_AMOUNT": 0,
        "INSTALLMENTS": lookup["INSTALLMENTS"].values,
        "PAYMENT_METHOD": lookup["PAYMENT_METHOD"].values,
    })
    soma_hk = df[["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT"]].sum(axis=1)
    df["NET_CREDIT_AMOUNT"] = soma_hk.where(soma_hk >= 0, 0)
    df["NET_DEBIT_AMOUNT"] = -soma_hk.where(soma_hk < 0, 0)
    return df[COLUNAS_PAGAMENTO]


def montar_para_processar(validacao, pagamento, frete, ab_pp):
    """Guia Para Processar: Pagamento + Frete colados, mais uma linha de
    AJUSTE por pedido cujo residual (`ProvaReal para Procesar` = Validação!AB
    = `SOMA(U:W) - líquido já lançado em Pagamento+Frete`) não fechou em
    zero — sobra sobretudo quando `Lançar Frete`="Não" mas o Custo Frete
    é ≠0 (fica de fora do Frete, sobra pra cá). ✅ Confirmado pelo Everton:
    o ajuste lança o valor residual em NET_CREDIT_AMOUNT **e**
    SHIPPING_FEE_AMOUNT (GROSS/MP_FEE/FINANCING_FEE/NET_DEBIT ficam 0);
    DESCRIPTION fixo em "mediation"; DATE/SOURCE_ID/RECORD_TYPE/INSTALLMENTS/
    PAYMENT_METHOD copiados da linha real do pedido em `ab_pp` (mesmo PROCX
    de primeira ocorrência usado em `montar_pagamento`/`montar_frete` — não
    inventar/deixar em branco)."""
    partes = pd.concat([pagamento, frete], ignore_index=True)
    liquido = (
        partes.groupby("EXTERNAL_REFERENCE")[["NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT"]]
        .sum()
        .assign(net=lambda d: d["NET_CREDIT_AMOUNT"] - d["NET_DEBIT_AMOUNT"])["net"]
    )

    alvo = validacao.set_index("PED_MKP")[["novo_gross_amount", "nova_comissao", "custo_frete"]].sum(axis=1)
    residual = round(alvo - liquido.reindex(alvo.index).fillna(0), 2)
    residual = residual[residual != 0]

    if residual.empty:
        return partes[COLUNAS_PAGAMENTO]

    ab_pp = ab_pp.copy()
    ab_pp["_chave"] = ab_pp["EXTERNAL_REFERENCE"].astype(str).str.strip()
    primeira = ab_pp.drop_duplicates("_chave", keep="first").set_index("_chave")
    lookup = primeira.reindex(residual.index)

    ajustes = pd.DataFrame({
        "DATE": lookup["DATE"].values, "SOURCE_ID": lookup["SOURCE_ID"].values,
        "EXTERNAL_REFERENCE": residual.index, "RECORD_TYPE": lookup["RECORD_TYPE"].values,
        "DESCRIPTION": "mediation",
        "NET_CREDIT_AMOUNT": residual.values, "NET_DEBIT_AMOUNT": 0,
        "GROSS_AMOUNT": 0, "MP_FEE_AMOUNT": 0, "FINANCING_FEE_AMOUNT": 0,
        "SHIPPING_FEE_AMOUNT": residual.values, "TAXES_AMOUNT": 0, "COUPON_AMOUNT": 0,
        "INSTALLMENTS": lookup["INSTALLMENTS"].values, "PAYMENT_METHOD": lookup["PAYMENT_METHOD"].values,
    })
    return pd.concat([partes, ajustes], ignore_index=True)[COLUNAS_PAGAMENTO]


def resumo_acao_manual(validacao):
    return validacao[validacao["categoria"].isin(CATEGORIAS_ACAO_MANUAL)]
