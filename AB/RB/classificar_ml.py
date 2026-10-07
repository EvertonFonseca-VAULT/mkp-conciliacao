"""Classificacao das linhas do detalhe de um pedido no ML/MP (comissao,
envio, rebate, desconto/bonus...) -- separado de scraper_ml.py de proposito
pra NAO precisar de playwright/Screaper instalado. scraper_ml.py usa isso
pra classificar o que acabou de raspar; rodar_conciliacao_rb.py tambem usa
(via carregar_descricoes_salvas) pra reaproveitar uma raspagem anterior
salva em descricoes_ml_<canal>_<data>.xlsx, sem precisar abrir o Chrome de
novo -- pedido do Everton, 16/09/2026.

NAO CONFIRMADO: os rotulos exatos variam entre pedidos (ex.: "Envios" vs
"Custo de envio"; "Recebimento pelo desconto da sua contraparte" vs
"Cancelamentos" vs "Estorno"). Casamos por substring, case-insensitive,
ignorando sublinhas de detalhamento ("  - ..."). Ajustar conforme mais
variacoes reais aparecerem.
"""

_ROTULOS = {
    "produto": ["preço do produto", "produto"],
    "comissao": ["tarifa de venda total"],
    "envio": ["custo de envio", "envio"],
    "total": ["total"],
}


def _eh_desconto_bonus(texto):
    """"Desconto" sozinho tambem aparece em rotulos que sao REBATE de
    verdade (ex. "Recebimento pelo desconto da sua contraparte" -- "estorno
    de desconto" nao e a mesma coisa que "desconto/bonus de comissao"). So
    conta como desconto/bonus de comissao quando "desconto" E "bonus"
    aparecem juntos (ex. "Descontos e bonus", pedido 47818962330). Pedido
    do Everton, 16/09/2026."""
    texto = texto.lower()
    return "desconto" in texto and ("bônus" in texto or "bonus" in texto)


def _eh_rebate(texto):
    """"contraparte" identifica o rebate de verdade (Recebimento pelo
    desconto da sua/pela contraparte) -- mas SO conta o lancamento
    original ("pai"), nao a "Cancelamento do recebimento pelo desconto da
    sua contraparte" que aparece dentro de Reembolsos quando o pedido e
    estornado (pedido 2000017834651356: conta so o +47,60, ignora o
    -47,60 -- pedido do Everton, 16/09/2026: "quando ele estiver dentro de
    algum pai voce nao calcula ele, so calcula quando ele for um pai")."""
    texto = texto.lower()
    if "contraparte" not in texto:
        return False
    return "cancelamento" not in texto and "estorno" not in texto


def _eh_custo_de_envio(texto):
    """A "Tarifa de cobranca" (ou "Tarifa de devolucao", mesmo tratamento --
    pedido devolvido/reembolsado) fica dentro do grupo Envios junto do custo
    de envio de verdade -- pro calculo do Repasse (dashboard.py) so conta o
    custo de envio, o resto fica de fora (residuo -- continua contando no
    total mostrado na coluna Frete, que usa o grupo Envios inteiro). O
    rotulo do Mercado varia: "Custo de envio" ou "Tarifa do Mercado Envios
    (por sua conta)", mesmo conceito -- confirmado com pedido real
    47833718202, 17/09/2026 (rotulo diferente do pedido 2000017834651356
    usado no self-test mais abaixo, mesmos valores). Pedido do Everton,
    16-17/09/2026."""
    texto = texto.lower()
    return "custo de envio" in texto or "tarifa do mercado envios" in texto


def _agrupar(linhas):
    """[(rotulo, valor)] pra [{label, valor, subs}] -- sublinha de
    detalhamento ("- ...") vira sub- do grupo anterior. Serve pro tooltip
    (breakdown de um total especifico, ex. Envios); a classificacao
    (desconto/bonus, rebate) busca na lista TODA (grupos + subs), ver
    _todas_linhas."""
    grupos = []
    for label, valor in linhas:
        if not label:
            continue
        texto = label.strip()
        if texto.startswith("-"):
            if grupos:
                grupos[-1]["subs"].append((texto[1:].strip(), valor))
        else:
            grupos.append({"label": texto, "valor": valor, "subs": []})
    return grupos


def _achar_grupo(grupos, chaves, inicio=False):
    for g in grupos:
        texto = g["label"].lower()
        bateu = any(texto.startswith(c) for c in chaves) if inicio else any(c in texto for c in chaves)
        if bateu:
            return g
    return None


def _todas_linhas(grupos):
    """Achata grupos + sublinhas numa lista flat (rotulo, valor) -- rebate e
    desconto/bonus podem aparecer em mais de uma linha (ex. recebimento +
    cancelamento do recebimento) espalhadas por grupos diferentes, entao a
    classificacao busca aqui em vez de olhar so o topo de cada grupo."""
    saida = []
    for g in grupos:
        saida.append((g["label"], g["valor"]))
        saida.extend(g["subs"])
    return saida


def _linhas_casadas(linhas_flat, predicado):
    return [(lbl, v) for lbl, v in linhas_flat if predicado(lbl)]


def _texto_por_rotulo(linhas, rotulo):
    """Valor (texto, nao numero) de uma linha crua pelo rotulo exato --
    "Motivo"/"Motivo detalhe" (ver Screaper.extrair_dados_detalhe_mp),
    presentes so em atividades sem ticket de compra (cashback etc).
    Pedido do Everton, 23/09/2026."""
    for lbl, v in linhas:
        if lbl and lbl.strip().lower() == rotulo:
            return v
    return None


def _somar(linhas_casadas):
    """None (nao "achou e deu zero") se a lista estiver vazia -- distingue
    "nao encontrado" de "encontrado, liquido zero" (ex. pedido reembolsado
    onde o rebate foi cancelado de volta)."""
    return round(sum(v for _, v in linhas_casadas), 2) if linhas_casadas else None


def _detalhe(grupo):
    """"Rotulo: valor" de cada sublinha do grupo, uma por linha -- com o
    mesmo texto literal que o Mercado mostrou (ex. "Tarifa do Mercado
    Envios (por sua conta)"). Vira tooltip (title) no dashboard ao passar
    o mouse no total. None se o grupo nao tiver sublinhas."""
    if not grupo or not grupo["subs"]:
        return None
    return _detalhe_linhas(grupo["subs"])


def _detalhe_linhas(linhas_casadas):
    if not linhas_casadas:
        return None
    return "\n".join(f"{lbl}: {v:.2f}".replace(".", ",") for lbl, v in linhas_casadas)


def _detalhe_comissao_liquida(grupo_comissao, linhas_desconto):
    """Comissao (rotulo literal, em valor absoluto) - Desconto/Bonus (idem,
    negativo, pode ser mais de 1 linha) -- a soma bate com
    ML_COMISSAO_LIQUIDA, que e sempre exibida positiva (ver enriquecer.py).
    Sem isso o tooltip mostraria so a tarifa bruta, que nao bate com o
    valor liquido que a coluna Comissao mostra."""
    linhas = []
    if grupo_comissao and grupo_comissao["subs"]:
        linhas.extend((lbl, abs(v)) for lbl, v in grupo_comissao["subs"])
    elif grupo_comissao:
        linhas.append((grupo_comissao["label"], abs(grupo_comissao["valor"])))
    linhas.extend((lbl, -abs(v)) for lbl, v in linhas_desconto)
    if not linhas:
        return None
    return "\n".join(f"{lbl}: {v:.2f}".replace(".", ",") for lbl, v in linhas)


def resumir_linhas(linhas, plataforma=None, pedido_ml=None):
    """Classifica uma lista crua [(rotulo, valor)] (ja com sublinhas
    prefixadas "- ") no mesmo formato que scraper_ml.resumo_pedido()
    retorna. Usado tanto direto na raspagem quanto reconstruindo a partir
    de descricoes_ml_<canal>_<data>.xlsx (ver carregar_descricoes_salvas)."""
    grupos = _agrupar(linhas)
    # Reembolsos e descartado ANTES de qualquer classificacao (nao so
    # filtrado predicado a predicado): uma sublinha dentro dele pode conter
    # a mesma substring de um rotulo valido (ex. "Cancelamento do custo de
    # envio" contem "custo de envio" e seria somada junto do Custo de envio
    # de verdade, inflando o valor -- bug real flagrado num print do
    # Everton, pedido com reembolso total, 17/09/2026).
    grupos = [g for g in grupos if "reembolso" not in g["label"].lower()]
    linhas_flat = _todas_linhas(grupos)

    grupo_comissao = _achar_grupo(grupos, _ROTULOS["comissao"])
    grupo_envio = _achar_grupo(grupos, _ROTULOS["envio"])
    grupo_produto = _achar_grupo(grupos, _ROTULOS["produto"])
    # "total" precisa bater no INICIO do rotulo -- senao "Total" colide com
    # o "l" final de "Tarifa de venda total", que aparece antes na lista e
    # tambem contem a substring "total".
    grupo_total = _achar_grupo(grupos, _ROTULOS["total"], inicio=True)

    # desconto/bonus e rebate podem aparecer em mais de 1 linha espalhada
    # (grupo proprio, sublinha de outro grupo, e a cancelamento dela num
    # pedido reembolsado) -- soma tudo que casar em qualquer lugar do
    # pedido, nao so a 1a ocorrencia.
    linhas_desconto = _linhas_casadas(linhas_flat, _eh_desconto_bonus)
    desconto_bonus = _somar(linhas_desconto)
    linhas_rebate = _linhas_casadas(linhas_flat, _eh_rebate)
    rebate = _somar(linhas_rebate)

    # "custo de envio" (sem a tarifa de cobranca) pro calculo do Repasse --
    # se nao achar rotulo especifico (pedido so tem 1 filho em Envios, sem
    # cobranca pra separar), cai pro valor do grupo Envios inteiro mesmo.
    linhas_custo_envio = _linhas_casadas(linhas_flat, _eh_custo_de_envio)
    custo_envio = _somar(linhas_custo_envio)
    if custo_envio is None:
        custo_envio = grupo_envio["valor"] if grupo_envio else None

    return {
        "plataforma": plataforma,
        "pedido_ml": pedido_ml,
        "produto": grupo_produto["valor"] if grupo_produto else None,
        "comissao": grupo_comissao["valor"] if grupo_comissao else None,
        "envio": grupo_envio["valor"] if grupo_envio else None,
        "custo_envio": custo_envio,
        "rebate": rebate,
        "desconto_bonus": desconto_bonus,
        "total": grupo_total["valor"] if grupo_total else None,
        # texto pronto (rotulo literal do Mercado + valor) das sublinhas de
        # cada grupo -- vira tooltip no dashboard, ver dashboard.py.
        "comissao_detalhe": _detalhe_comissao_liquida(grupo_comissao, linhas_desconto),
        "envio_detalhe": _detalhe(grupo_envio),
        "rebate_detalhe": _detalhe_linhas(linhas_rebate),
        "desconto_bonus_detalhe": _detalhe_linhas(linhas_desconto),
        # Motivo/Motivo detalhe: so em atividades sem ticket de compra
        # (cashback etc) -- "detalhe" traz a descricao completa, com o
        # pedido/venda relacionado embutido. Pedido do Everton, 23/09/2026.
        "motivo": _texto_por_rotulo(linhas, "motivo"),
        "motivo_detalhe": _texto_por_rotulo(linhas, "motivo detalhe"),
        # linhas cruas (inclui sublinhas de detalhamento "- ...", que os
        # buckets acima ignoram) -- ver descricoes_ml.py.
        "linhas_raw": linhas,
    }


def carregar_descricoes_salvas(caminho_xlsx):
    """Reconstroi {PEDIDO_MKP: resumo} a partir de um
    descricoes_ml_<canal>_<data>.xlsx ja salvo (colunas PEDIDO_MKP, ROTULO,
    VALOR, DETALHE -- ver descricoes_ml.py), reclassificando com as regras
    ATUAIS (_eh_rebate etc.) em vez de reusar numeros antigos. Deixa
    reaproveitar uma raspagem anterior sem abrir o Chrome de novo -- pedido
    do Everton, 16/09/2026. None se o arquivo nao existir."""
    import os

    import pandas as pd

    if not os.path.exists(caminho_xlsx):
        return None

    df = pd.read_excel(caminho_xlsx)
    resultado = {}
    for pedido, grupo in df.groupby("PEDIDO_MKP", sort=False):
        linhas = list(zip(grupo["ROTULO"], grupo["VALOR"]))
        resultado[str(pedido)] = resumir_linhas(linhas, pedido_ml=str(pedido))
    return resultado


if __name__ == "__main__":
    # Pedido real 2000017834651356 (reembolsado) -- ver conversa com o
    # Everton, 16/09/2026.
    linhas = [
        ("Tarifa de venda total (R$)", -38.68),
        ("- Custo por vender no Mercado Livre (R$)", -38.34),
        ("- Custo por cobrar no Mercado Pago (R$)", -0.34),
        ("Envios (R$)", -132.15),
        ("- Tarifa de devolução (R$)", -88.1),
        ("- Custo de envio (R$)", -44.05),
        ("Descontos (R$)", 47.6),
        ("- Recebimento pelo desconto da sua contraparte (R$)", 47.6),
        ("Reembolsos (R$)", -556.32),
        ("- Cancelamento do pagamento recebido (R$)", -547.4),
        ("- Cancelamento do recebimento pelo desconto da sua contraparte (R$)", -47.6),
    ]
    r = resumir_linhas(linhas)
    assert r["comissao"] == -38.68
    assert r["custo_envio"] == -44.05, r["custo_envio"]
    assert r["rebate"] == 47.6, r["rebate"]  # so o "pai", nao soma com a cancelamento
    assert r["desconto_bonus"] is None

    # Pedido real com reembolso TOTAL (print do Everton, 17/09/2026) -- o
    # grupo Reembolsos tem "Cancelamento do custo de envio: -184,99", que
    # SEM o filtro de grupo seria somado ao Custo de envio de verdade
    # (-106,95), dando -291,94 (bug real em producao, coluna Frete do
    # pedido 2000018014989912). Com o filtro, Reembolsos inteiro e
    # descartado antes de qualquer classificacao.
    linhas_reembolso_total = [
        ("Preço do produto (R$)", 1509.0),
        ("Custo de envio (R$)", -106.95),
        ("Tarifa de venda total (R$)", -64.68),
        ("- Custo por vender no Mercado Livre (R$)", -48.42),
        ("- Custo por cobrar no Mercado Pago (R$)", -14.52),
        ("- Taxa de parcelamento (R$)", -1.74),
        ("Recebimento pelo desconto da sua contraparte (R$)", 10.0),
        ("Reembolsos (R$)", -1639.31),
        ("- Cancelamento do pagamento recebido (R$)", -1509.0),
        ("- Cancelamento do recebimento pelo desconto da sua contraparte (R$)", -10.0),
        ("- Cancelamento de custo por vender no Mercado Livre (R$)", 48.42),
        ("- Cancelamento de custo por cobrar no Mercado Pago (R$)", 14.52),
        ("- Cancelamento da taxa de parcelamento (R$)", 1.74),
        ("- Cancelamento do custo de envio (R$)", -184.99),
    ]
    r2 = resumir_linhas(linhas_reembolso_total)
    assert r2["custo_envio"] == -106.95, r2["custo_envio"]  # sem o fix, dava -291.94
    assert r2["comissao"] == -64.68
    assert r2["rebate"] == 10.0
    assert r2["desconto_bonus"] is None

    # Pedido real 47833718202 (descricoes_ml_full_08092026.xlsx) -- rotulo
    # "Tarifa do Mercado Envios (por sua conta)" em vez de "Custo de envio",
    # mesmo conceito. Sem o fix nenhuma sublinha casava, caia no fallback
    # (custo_envio = grupo Envios inteiro, -132.15) e o residuo (tarifa de
    # devolucao) sumia -- so 1 checkbox no dashboard em vez de 2.
    linhas_tarifa_mercado_envios = [
        ("Envios (R$)", -132.15),
        ("- Tarifa do Mercado Envios (por sua conta) (R$)", -44.05),
        ("- Tarifa de devolução (R$)", -88.10),
    ]
    r3 = resumir_linhas(linhas_tarifa_mercado_envios)
    assert r3["custo_envio"] == -44.05, r3["custo_envio"]  # sem o fix, dava -132.15
    assert round(r3["envio"] - r3["custo_envio"], 2) == -88.10  # residuo = Tarifa de devolucao

    # Cashback (pedido real 3299121906, MP, 23/09/2026) -- sem ticket de
    # compra, so Motivo/Motivo detalhe (ver Screaper.extrair_dados_detalhe_mp).
    linhas_cashback = [
        ("Pedido", "3299121906"), ("Status", "approved"),
        ("Motivo", "Entrada de dinheiro por seu envio"),
        ("Motivo detalhe", "Reembolsamos R$BRL2.872,45 pelo que aconteceu na devolução de sua venda 2000017769032090"),
    ]
    r4 = resumir_linhas(linhas_cashback)
    assert r4["motivo"] == "Entrada de dinheiro por seu envio"
    assert "2000017769032090" in r4["motivo_detalhe"]
    assert r4["comissao"] is None and r4["rebate"] is None  # nao e comissao/rebate, so info

    print("classificar_ml.py: ok")
