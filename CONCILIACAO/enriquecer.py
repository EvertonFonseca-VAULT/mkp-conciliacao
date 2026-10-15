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


def _rebate_com_sinal_do_dif(dif, rebate, detalhe):
    """O Rebate mostrado no dash segue o sinal do DIF (DIF positivo ->
    rebate positivo, DIF negativo -> rebate negativo) -- pedido do Everton,
    16/09/2026. O Mercado mostra o valor com o sinal dele proprio, que nao
    necessariamente bate com a direcao do ajuste que a AB espera. Aplica o
    mesmo sinal nas linhas do detalhe (tooltip), senao a soma para de bater
    com o total exibido."""
    if rebate is None:
        return None, detalhe
    sinal = 1 if dif >= 0 else -1
    rebate_sinalizado = abs(rebate) * sinal
    if not detalhe:
        return rebate_sinalizado, detalhe
    linhas = []
    for linha in detalhe.split("\n"):
        rotulo, _, valor_txt = linha.rpartition(": ")
        try:
            valor = float(valor_txt.replace(",", "."))
        except ValueError:
            linhas.append(linha)
            continue
        linhas.append(f"{rotulo}: {abs(valor) * sinal:.2f}".replace(".", ","))
    return rebate_sinalizado, "\n".join(linhas)


def enriquecer(resultado, raiz, ml_dados=None, rebate_lancado_rbx=None):
    """`ml_dados`: {PEDIDO_MKP: resumo_pedido_ou_None} vindo de
    scraper_ml.resumir_pedidos -- so precisa ser passado pra pedidos
    VERIFICAR (economiza tempo de scraping nos que ja bateram).

    `rebate_lancado_rbx`: {PEDIDO_MKP: valor} vindo de
    conferir_rbx.rebate_lancado_por_pedido -- pedidos VERIFICAR cujo Rebate
    (achado pelo scraper na pagina do ML/MP) JA tem movimento na RBX hoje
    (credito normal ou cancelamento). A pagina do ML/MP nao sabe que o banco
    ja reverteu o cupom -- sem avisar, o Everton pode lancar/cancelar o
    rebate de novo no CMM, contando o mesmo ajuste 2x (achado real, FULL
    14/09/2026, pedido 2000018111220588: Rebate 136,16 ainda visivel no ML,
    mas ja CANCELADO na RBX). NAO esconde o valor -- so avisa. Pedido do
    Everton, 22/09/2026."""
    ml_dados = ml_dados or {}
    rebate_lancado_rbx = rebate_lancado_rbx or {}
    df = resultado.copy()

    df["HIST_COMISSAO"] = f"{raiz}_COMISSAO"
    df["HIST_REPASSE"] = f"{raiz}_REPASSE"
    df["HIST_FRETE"] = f"{raiz}_FRETE"

    df["ML_COMISSAO"] = None
    df["ML_ENVIO"] = None
    df["ML_CUSTO_ENVIO"] = None
    df["ML_REBATE"] = None
    df["ML_DESCONTO_BONUS"] = None
    df["ML_COMISSAO_LIQUIDA"] = None
    df["ML_TOTAL"] = None
    # texto (rotulo literal do Mercado + valor) pra tooltip no dashboard --
    # ver scraper_ml.py._detalhe / dashboard.py._cel_detalhe.
    df["ML_COMISSAO_DETALHE"] = None
    df["ML_ENVIO_DETALHE"] = None
    df["ML_REBATE_DETALHE"] = None
    df["ML_DESCONTO_BONUS_DETALHE"] = None
    # Motivo/Motivo detalhe: so em atividades sem ticket de compra (cashback
    # etc, ver classificar_ml.py) -- vira tooltip na Descricao do dashboard.
    # Pedido do Everton, 23/09/2026.
    df["ML_MOTIVO"] = None
    df["ML_MOTIVO_DETALHE"] = None
    df["ACAO_SUGERIDA"] = ""

    for i, row in df.iterrows():
        status = row["STATUS"]

        if status == "SEM DIFF":
            df.at[i, "ACAO_SUGERIDA"] = "Nada a fazer."
            continue
        if status == "REBATE" and row["DIF"] == 0:
            df.at[i, "ACAO_SUGERIDA"] = (
                f"Resolvido pela formula -- incluir na planilha CMM (lote) o rebate de {row['REBATE']:.2f}."
            )
            continue
        if status == "EST COM":
            df.at[i, "ACAO_SUGERIDA"] = "Resolvido pela formula -- estorno de comissao."
            continue

        # status == VERIFICAR, ou REBATE com DIF != 0 (cupom/CMM resolvido
        # pela formula, MAS tem um gap de banco sem relacao com o cupom que
        # precisa do scraper que nem qualquer VERIFICAR -- ver
        # _candidatos_verificar em rodar_conciliacao_rb.py e docstring deste
        # modulo). nota_cmm so pra lembrar que o rebate em si ja esta
        # resolvido, separado do que a sugestao abaixo resolve.
        nota_cmm = (
            f"[Rebate de {row['REBATE']:.2f} ja resolvido pela formula, vai pro CMM separado -- "
            "a sugestao abaixo e so pro gap de banco, sem relacao com o cupom] "
        ) if status == "REBATE" else ""
        dados = ml_dados.get(row["PEDIDO_MKP"])

        ja_na_rbx = rebate_lancado_rbx.get(row["PEDIDO_MKP"])
        aviso_rebate = nota_cmm + (
            f"⚠ Rebate desse pedido ja tem movimento na RBX hoje ({ja_na_rbx:.2f}) -- "
            "conferir antes de lancar/cancelar rebate de novo no CMM (pode ja estar la, ver 'Cancelamento de rebate' "
            "ou credito direto do ML na RBX). "
        ) if ja_na_rbx else nota_cmm

        if not dados or "erro" in dados:
            df.at[i, "ACAO_SUGERIDA"] = aviso_rebate + "Nao encontrado no ML/MP (ou erro na busca) -- confirmar numero do pedido manualmente."
            continue

        comissao = dados.get("comissao")
        envio = dados.get("envio")
        custo_envio = dados.get("custo_envio")
        rebate = dados.get("rebate")
        desconto_bonus = dados.get("desconto_bonus")
        total = dados.get("total")

        # Comissao "real" a baixar = tarifa - desconto/bonus (ambos em
        # valor absoluto) -- desconto/bonus e um credito que o MP da sobre
        # a propria tarifa, NAO e rebate (nao sai do repasse, nao vai pro
        # CMM em lote) -- pedido real 47818962330: tarifa 79,68, desconto
        # 11,95, comissao liquida 67,73. Ver scraper_ml.py.
        comissao_abs = abs(comissao) if comissao is not None else 0.0
        desconto_bonus_abs = abs(desconto_bonus) if desconto_bonus else 0.0
        comissao_liquida = round(comissao_abs - desconto_bonus_abs, 2)

        rebate_sinalizado, rebate_detalhe = _rebate_com_sinal_do_dif(
            row["DIF"], rebate, dados.get("rebate_detalhe"))

        df.at[i, "ML_COMISSAO"] = comissao
        df.at[i, "ML_ENVIO"] = envio
        df.at[i, "ML_CUSTO_ENVIO"] = custo_envio
        df.at[i, "ML_REBATE"] = rebate_sinalizado
        df.at[i, "ML_DESCONTO_BONUS"] = desconto_bonus
        df.at[i, "ML_COMISSAO_LIQUIDA"] = comissao_liquida
        df.at[i, "ML_TOTAL"] = total
        df.at[i, "ML_COMISSAO_DETALHE"] = dados.get("comissao_detalhe")
        df.at[i, "ML_ENVIO_DETALHE"] = dados.get("envio_detalhe")
        df.at[i, "ML_REBATE_DETALHE"] = rebate_detalhe
        df.at[i, "ML_DESCONTO_BONUS_DETALHE"] = dados.get("desconto_bonus_detalhe")
        df.at[i, "ML_MOTIVO"] = dados.get("motivo")
        df.at[i, "ML_MOTIVO_DETALHE"] = dados.get("motivo_detalhe")

        if total is None:
            df.at[i, "ACAO_SUGERIDA"] = aviso_rebate + "Achou o pedido mas nao consegui ler o Total -- conferir manualmente no ML/MP."
            continue

        hist_com, hist_rep = row["HIST_COMISSAO"], row["HIST_REPASSE"]
        nota_desconto = f" (comissao ja liquida do desconto/bonus de {desconto_bonus_abs:.2f})" if desconto_bonus_abs else ""
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
            df.at[i, "ACAO_SUGERIDA"] = aviso_rebate + (
                f"Sinal de estorno/cancelamento (repasse esperado pela AB e negativo: {repasse_ab:.2f}) -- "
                f"NAO confiar direto no Total do ML/MP ({total:.2f}), pode ser o valor de antes de um ajuste "
                f"posterior. Conferir o Saldo do titulo no TOTVS e seguir o processo de cancelamento "
                f"(Caso 2 da documentacao): cancelar o que ja estiver lancado e relancar so a diferenca real "
                f"(nunca fracionar -- baixar o valor certo de uma vez)."
            )
        elif row["TOTAL_RB"] == 0:
            if total == 0:
                df.at[i, "ACAO_SUGERIDA"] = aviso_rebate + (
                    "Reembolso total no ML/MP (Total = 0,00) -- conferir titulo no TOTVS: "
                    "se ja tiver algo lancado, cancelar/excluir; se nao tiver nada, nao lancar."
                )
            else:
                df.at[i, "ACAO_SUGERIDA"] = aviso_rebate + (
                    f"Pedido nunca lancado -- baixar Comissao {comissao_liquida:.2f}{nota_desconto} "
                    f"({hist_com}, MERC PAGO, sem banco) e Repasse {repasse_a_baixar:.2f} "
                    f"({hist_rep}, NORMAL, banco {banco})."
                    + (f" Rebate de {rebate_abs:.2f} (ja EXCLUIDO do repasse acima) vai pro CMM (lote)." if rebate_abs else "")
                )
        else:
            df.at[i, "ACAO_SUGERIDA"] = aviso_rebate + (
                f"Ja tem lancamento no banco mas nao bate -- conferir Saldo do titulo no TOTVS. "
                f"Valor correto segundo o ML/MP: Comissao {comissao_liquida:.2f}{nota_desconto} ({hist_com}) e "
                f"Repasse {repasse_a_baixar:.2f} ({hist_rep}, banco {banco}). "
                f"Cancelar o que estiver errado e relancar com esses valores."
                + (f" Rebate de {rebate_abs:.2f} (ja EXCLUIDO do repasse acima) vai pro CMM (lote)." if rebate_abs else "")
            )

    return df


if __name__ == "__main__":
    import pandas as pd

    # Pedido real 47818962330: Tarifa de venda total -79.68, dentro de
    # "Estorno" um "- Descontos e bonus" de 11.95 -- comissao liquida
    # esperada: 79.68 - 11.95 = 67.73, e NAO deve contar como rebate. DIF
    # positivo -> rebate (2a linha) tem que sair positivo tambem.
    resultado = pd.DataFrame([
        {"STATUS": "VERIFICAR", "PEDIDO_MKP": "47818962330",
         "REPASSE": 374.27, "TOTAL_RB": 0.0, "DIF": 374.27},
        {"STATUS": "VERIFICAR", "PEDIDO_MKP": "999",
         "REPASSE": -50.0, "TOTAL_RB": 0.0, "DIF": -50.0},
        # Achado real (FULL 14/09/2026, pedido 2000018111220588): Rebate
        # ainda aparece na pagina do ML, mas ja foi CANCELADO na RBX hoje --
        # tem que avisar, sem esconder o valor.
        {"STATUS": "VERIFICAR", "PEDIDO_MKP": "2000018111220588",
         "REPASSE": -737.61, "TOTAL_RB": 0.0, "DIF": -737.61},
    ])
    ml_dados = {
        "47818962330": {
            "comissao": -79.68, "envio": -44.05, "rebate": None,
            "desconto_bonus": 11.95, "total": 386.22,
            "comissao_detalhe": "Tarifa de 16%: 79,68\nDescontos e bonus: -11,95",
            "envio_detalhe": "Tarifa do Mercado Envios (por sua conta): -44,05",
        },
        "999": {
            "comissao": -10.0, "envio": None, "desconto_bonus": None, "total": 40.0,
            "rebate": 50.0,  # Mercado mostra positivo, mas DIF e negativo
            "rebate_detalhe": "Recebimento pelo desconto da sua contraparte: 50,00",
        },
        "2000018111220588": {
            "comissao": -130.44, "envio": -212.85, "desconto_bonus": None, "total": 595.71,
            "rebate": 136.16,
            "rebate_detalhe": "Recebimento pelo desconto da sua contraparte: 136,16",
        },
    }
    rebate_lancado_rbx = {"2000018111220588": -136.16}  # CAN- na RBX (ja cancelado)
    df = enriquecer(resultado, "MKP_MELI_FU", ml_dados, rebate_lancado_rbx)
    linha = df.iloc[0]
    assert linha["ML_DESCONTO_BONUS"] == 11.95, linha["ML_DESCONTO_BONUS"]
    assert linha["ML_COMISSAO_LIQUIDA"] == 67.73, linha["ML_COMISSAO_LIQUIDA"]
    assert "67.73" in linha["ACAO_SUGERIDA"], linha["ACAO_SUGERIDA"]
    assert "desconto/bonus de 11.95" in linha["ACAO_SUGERIDA"], linha["ACAO_SUGERIDA"]
    assert linha["ML_ENVIO_DETALHE"].startswith("Tarifa do Mercado Envios")
    soma_detalhe_comissao = sum(
        float(l.rpartition(": ")[2].replace(",", ".")) for l in linha["ML_COMISSAO_DETALHE"].split("\n")
    )
    assert abs(soma_detalhe_comissao - linha["ML_COMISSAO_LIQUIDA"]) < 0.001, soma_detalhe_comissao

    linha2 = df.iloc[1]
    assert linha2["ML_REBATE"] == -50.0, linha2["ML_REBATE"]  # DIF negativo -> rebate negativo
    assert linha2["ML_REBATE_DETALHE"] == "Recebimento pelo desconto da sua contraparte: -50,00", linha2["ML_REBATE_DETALHE"]
    assert "⚠" not in linha2["ACAO_SUGERIDA"], linha2["ACAO_SUGERIDA"]  # sem entrada em rebate_lancado_rbx -- sem aviso

    linha3 = df.iloc[2]
    assert linha3["ML_REBATE"] == -136.16, linha3["ML_REBATE"]  # DIF negativo -> rebate negativo
    assert "⚠" in linha3["ACAO_SUGERIDA"], linha3["ACAO_SUGERIDA"]  # ja cancelado na RBX -- avisa
    assert "-136.16" in linha3["ACAO_SUGERIDA"], linha3["ACAO_SUGERIDA"]
    assert "Sinal de estorno/cancelamento" in linha3["ACAO_SUGERIDA"], linha3["ACAO_SUGERIDA"]  # aviso NAO troca a acao, so avisa antes

    # REBATE com DIF == 0 (cupom resolvido, banco ja bateu certo): continua
    # so a mensagem da formula, NUNCA entra no scraper.
    resultado_rebate = pd.DataFrame([
        {"STATUS": "REBATE", "PEDIDO_MKP": "1", "REPASSE": 100.0, "TOTAL_RB": 100.0, "DIF": 0.0, "REBATE": -40.40},
        {"STATUS": "REBATE", "PEDIDO_MKP": "2", "REPASSE": 686.12, "TOTAL_RB": 551.01, "DIF": 135.11, "REBATE": -40.40},
    ])
    ml_dados_rebate = {"2": {"comissao": -10.0, "envio": None, "desconto_bonus": None, "total": 686.12, "rebate": None}}
    df_rebate = enriquecer(resultado_rebate, "MKP_MELI_WE", ml_dados_rebate, None)
    linha_a = df_rebate.iloc[0]
    assert linha_a["ACAO_SUGERIDA"] == "Resolvido pela formula -- incluir na planilha CMM (lote) o rebate de -40.40."
    assert linha_a["ML_TOTAL"] is None, linha_a["ML_TOTAL"]  # nunca passou pelo scraper
    # REBATE com DIF != 0 (cupom estornado, mas tem gap de banco sem
    # relacao): usa o scraper que nem VERIFICAR, so avisando que o rebate em
    # si ja esta resolvido.
    linha_b = df_rebate.iloc[1]
    assert "[Rebate de -40.40 ja resolvido pela formula" in linha_b["ACAO_SUGERIDA"], linha_b["ACAO_SUGERIDA"]
    assert linha_b["ML_TOTAL"] == 686.12, linha_b["ML_TOTAL"]  # passou pelo scraper de verdade
    print("enriquecer.py: ok")
