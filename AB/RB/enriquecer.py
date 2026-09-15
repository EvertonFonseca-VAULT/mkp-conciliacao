"""Enriquece o resultado do cruzamento (cruzamento.py) com o que falta pra
resolver um pedido `VERIFICAR` sem abrir TOTVS/ML varias vezes: o histórico
sugerido pra cada baixa e uma ação sugerida com os valores certos (via
scraper_ml, quando disponivel). Ver
Manuais/Passo a passo - Resolução manual de pendências (Baixas no TOTVS).md
pros 6 casos que essa lógica tenta cobrir.

NAO SUBSTITUI o julgamento do analista -- casos de cancelamento/ajuste
ainda exigem conferir o Saldo do título no TOTVS (esse programa não lê o
TOTVS diretamente). A ação sugerida assume o cenário mais comum (pedido
nunca lançado) quando não há nada no RB ainda.
"""

BANCO_TRANSITORIA = ("237", "3708", "X103925")


def enriquecer(resultado, raiz, ml_dados=None):
    """`ml_dados`: {PEDIDO_MKP: resumo_pedido_ou_None} vindo de
    scraper_ml.resumir_pedidos -- so precisa ser passado pra pedidos
    VERIFICAR (economiza tempo de scraping nos que ja bateram)."""
    ml_dados = ml_dados or {}
    df = resultado.copy()

    df["HIST_COMISSAO"] = f"{raiz}_COMISSAO"
    df["HIST_REPASSE"] = f"{raiz}_REPASSE"
    df["HIST_FRETE"] = f"{raiz}_FRETE"

    df["ML_COMISSAO"] = None
    df["ML_ENVIO"] = None
    df["ML_REBATE"] = None
    df["ML_TOTAL"] = None
    df["ACAO_SUGERIDA"] = ""

    for i, row in df.iterrows():
        status = row["STATUS"]

        if status == "SEM DIFF":
            df.at[i, "ACAO_SUGERIDA"] = "Nada a fazer."
            continue
        if status == "REBATE":
            df.at[i, "ACAO_SUGERIDA"] = (
                f"Resolvido pela formula -- incluir na planilha CMM (lote) o rebate de {row['REBATE']:.2f}."
            )
            continue
        if status == "MEDIACAO":
            df.at[i, "ACAO_SUGERIDA"] = "Resolvido pela formula -- diferenca e de mediacao, nao precisa de baixa manual."
            continue
        if status == "EST COM":
            df.at[i, "ACAO_SUGERIDA"] = "Resolvido pela formula -- estorno de comissao."
            continue

        # status == VERIFICAR: aqui e onde o scraper ajuda de verdade
        dados = ml_dados.get(row["PEDIDO_MKP"])

        if not dados or "erro" in dados:
            df.at[i, "ACAO_SUGERIDA"] = "Nao encontrado no ML/MP (ou erro na busca) -- confirmar numero do pedido manualmente."
            continue

        comissao = dados.get("comissao")
        envio = dados.get("envio")
        rebate = dados.get("rebate")
        total = dados.get("total")

        df.at[i, "ML_COMISSAO"] = comissao
        df.at[i, "ML_ENVIO"] = envio
        df.at[i, "ML_REBATE"] = rebate
        df.at[i, "ML_TOTAL"] = total

        if total is None:
            df.at[i, "ACAO_SUGERIDA"] = "Achou o pedido mas nao consegui ler o Total -- conferir manualmente no ML/MP."
            continue

        hist_com, hist_rep = row["HIST_COMISSAO"], row["HIST_REPASSE"]
        comissao_abs = abs(comissao) if comissao is not None else 0.0
        banco = "/".join(BANCO_TRANSITORIA)
        repasse_ab = row["REPASSE"]

        # O rebate (quando existe) vai pro CMM em lote, NUNCA na baixa do
        # titulo -- tem que sair do valor de repasse a baixar, senao conta o
        # mesmo valor duas vezes (achado real no pedido 2000017590233144).
        rebate_abs = abs(rebate) if rebate else 0.0
        repasse_a_baixar = round(total - rebate_abs, 2)

        if repasse_ab < 0:
            # Sinal de estorno/cancelamento -- o Total do ML/MP e so o valor
            # "de face" da venda, pode nao refletir um ajuste/estorno feito
            # DEPOIS (visto no pedido 47735704864: ML ainda mostrava 1.107,25
            # mesmo apos o estorno que a AB ja capturou como -921,84). Nao dar
            # uma ordem confiante de "baixar X" aqui -- e o Caso 2, exige
            # conferencia manual do Saldo do titulo.
            df.at[i, "ACAO_SUGERIDA"] = (
                f"Sinal de estorno/cancelamento (repasse esperado pela AB e negativo: {repasse_ab:.2f}) -- "
                f"NAO confiar direto no Total do ML/MP ({total:.2f}), pode ser o valor de antes de um ajuste "
                f"posterior. Conferir o Saldo do titulo no TOTVS e seguir o processo de cancelamento "
                f"(Caso 2 da documentacao): cancelar o que ja estiver lancado e relancar so a diferenca real "
                f"(nunca fracionar -- baixar o valor certo de uma vez)."
            )
        elif row["TOTAL_RB"] == 0:
            if total == 0:
                df.at[i, "ACAO_SUGERIDA"] = (
                    "Reembolso total no ML/MP (Total = 0,00) -- conferir titulo no TOTVS: "
                    "se ja tiver algo lancado, cancelar/excluir; se nao tiver nada, nao lancar."
                )
            else:
                df.at[i, "ACAO_SUGERIDA"] = (
                    f"Pedido nunca lancado -- baixar Comissao {comissao_abs:.2f} "
                    f"({hist_com}, MERC PAGO, sem banco) e Repasse {repasse_a_baixar:.2f} "
                    f"({hist_rep}, NORMAL, banco {banco})."
                    + (f" Rebate de {rebate_abs:.2f} (ja EXCLUIDO do repasse acima) vai pro CMM (lote)." if rebate_abs else "")
                )
        else:
            df.at[i, "ACAO_SUGERIDA"] = (
                f"Ja tem lancamento no banco mas nao bate -- conferir Saldo do titulo no TOTVS. "
                f"Valor correto segundo o ML/MP: Comissao {comissao_abs:.2f} ({hist_com}) e "
                f"Repasse {repasse_a_baixar:.2f} ({hist_rep}, banco {banco}). "
                f"Cancelar o que estiver errado e relancar com esses valores."
                + (f" Rebate de {rebate_abs:.2f} (ja EXCLUIDO do repasse acima) vai pro CMM (lote)." if rebate_abs else "")
            )

    return df
