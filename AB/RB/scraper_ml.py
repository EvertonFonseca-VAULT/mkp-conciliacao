"""Busca no Mercado Livre/Mercado Pago (reaproveitando o Screaper Mercado ja
existente, Playwright + Chrome com debug remoto) os valores reais de um
pedido: produto, tarifa (comissao), envio, rebate/bonus e total recebido.

Precisa do Chrome aberto com debug remoto (porta 9222), ja logado em
mercadolivre.com.br e mercadopago.com.br -- ver
Manuais/Documentação Screaper MP e ML.md.

Roda com o Python do venv do Screaper Mercado (tem playwright instalado):
    "C:\\Users\\everton.fonseca\\Desktop\\Projetos\\Screaper Mercado\\.venv\\Scripts\\python.exe" ...
"""

import os
import sys
import time

CAMINHO_SCRAPER = os.path.join(os.path.expanduser("~"), "Desktop", "Projetos", "Screaper Mercado")
if CAMINHO_SCRAPER not in sys.path:
    sys.path.insert(0, CAMINHO_SCRAPER)

from playwright.sync_api import sync_playwright  # noqa: E402

from Screaper import (  # noqa: E402
    buscar_e_abrir_ultimo_resultado_ml,
    buscar_e_abrir_ultimo_resultado_mp,
    conectar_chrome_existente,
    encontrar_aba_mercadolivre,
    encontrar_aba_mercadopago,
    extrair_dados_detalhe_ml,
    extrair_dados_detalhe_mp,
)

# NAO CONFIRMADO: os rotulos exatos variam entre pedidos (ex.: "Envios" vs
# "Custo de envio"; "Recebimento pelo desconto da sua contraparte" vs
# "Cancelamentos" vs "Estorno"). Casamos por substring, case-insensitive,
# ignorando sublinhas de detalhamento ("  - ..."). Ajustar conforme mais
# variacoes reais aparecerem.
_ROTULOS = {
    "produto": ["preço do produto", "produto"],
    "comissao": ["tarifa de venda total"],
    "envio": ["custo de envio", "envio"],
    # cancelamento/estorno de VERDADE (credito de pedido cancelado) -- vai
    # pro CMM em lote e sai do repasse a baixar (enriquecer.py). NAO inclui
    # desconto/bonus -- ver "desconto_bonus" abaixo (pedido 47818962330:
    # "Estorno" so existia pra embrulhar um "Descontos e bonus", que e outra
    # coisa -- abate a comissao, nao e rebate).
    "rebate": ["cancelamento", "estorno"],
    "desconto_bonus": ["desconto", "bônus", "bonus"],
    "total": ["total"],
}


def _agrupar(linhas):
    """[(rotulo, valor)] pra [{label, valor, subs}] -- sublinha de
    detalhamento ("- ...") vira sub- do grupo anterior. Precisa disso pra
    saber se um grupo tipo "Estorno" embrulha um "Descontos e bonus" (deixa
    de ser rebate, vira abatimento de comissao) em vez de so olhar o rotulo
    do topo."""
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


def _detalhe(grupo):
    """"Rotulo: valor" de cada sublinha do grupo, uma por linha -- com o
    mesmo texto literal que o Mercado mostrou (ex. "Tarifa do Mercado
    Envios (por sua conta)"). Vira tooltip (title) no dashboard ao passar
    o mouse no total. None se o grupo nao tiver sublinhas."""
    if not grupo or not grupo["subs"]:
        return None
    return "\n".join(f"{lbl}: {v:.2f}".replace(".", ",") for lbl, v in grupo["subs"])


def resumo_pedido(numero, context):
    """`numero` = SHIPPING_ID (ML, comeca com "4") ou EXTERNAL_REFERENCE
    (Mercado Pago, os demais). Retorna None se nao encontrar nada."""
    numero = str(numero).strip()

    if numero.startswith("4"):
        page = encontrar_aba_mercadolivre(context)
        if page is None:
            page = context.new_page()
        html = buscar_e_abrir_ultimo_resultado_ml(page, numero)
        dados = extrair_dados_detalhe_ml(html) if html else None
        plataforma = "ML"
    else:
        page = encontrar_aba_mercadopago(context)
        if page is None:
            return None
        html = buscar_e_abrir_ultimo_resultado_mp(page, numero)
        dados = extrair_dados_detalhe_mp(html) if html else None
        plataforma = "MP"

    if not dados:
        return None

    linhas = dados["linhas"]
    grupos = _agrupar(linhas)

    grupo_comissao = _achar_grupo(grupos, _ROTULOS["comissao"])
    grupo_envio = _achar_grupo(grupos, _ROTULOS["envio"])
    grupo_produto = _achar_grupo(grupos, _ROTULOS["produto"])
    # "total" precisa bater no INICIO do rotulo -- senao "Total" colide com
    # o "l" final de "Tarifa de venda total", que aparece antes na lista e
    # tambem contem a substring "total".
    grupo_total = _achar_grupo(grupos, _ROTULOS["total"], inicio=True)

    # desconto/bonus pode vir como grupo proprio OU como sublinha de um
    # grupo estorno/cancelamento (pedido 47818962330: "Estorno" > "-
    # Descontos e bonus") -- nesse 2o caso NAO conta como rebate.
    grupo_desconto = _achar_grupo(grupos, _ROTULOS["desconto_bonus"])
    desconto_bonus = grupo_desconto["valor"] if grupo_desconto else None
    grupo_desconto_fonte = grupo_desconto

    grupo_rebate = _achar_grupo(grupos, _ROTULOS["rebate"])
    rebate = None
    if grupo_rebate:
        sub_desconto = next(
            (v for lbl, v in grupo_rebate["subs"] if any(c in lbl.lower() for c in _ROTULOS["desconto_bonus"])),
            None,
        )
        if sub_desconto is not None:
            if desconto_bonus is None:
                desconto_bonus = sub_desconto
                grupo_desconto_fonte = grupo_rebate
        else:
            rebate = grupo_rebate["valor"]

    return {
        "plataforma": plataforma,
        "pedido_ml": dados["pedido"],
        "produto": grupo_produto["valor"] if grupo_produto else None,
        "comissao": grupo_comissao["valor"] if grupo_comissao else None,
        "envio": grupo_envio["valor"] if grupo_envio else None,
        "rebate": rebate,
        "desconto_bonus": desconto_bonus,
        "total": grupo_total["valor"] if grupo_total else None,
        # texto pronto (rotulo literal do Mercado + valor) das sublinhas de
        # cada grupo -- vira tooltip no dashboard, ver dashboard.py.
        "comissao_detalhe": _detalhe(grupo_comissao),
        "envio_detalhe": _detalhe(grupo_envio),
        "rebate_detalhe": _detalhe(grupo_rebate) if rebate is not None else None,
        "desconto_bonus_detalhe": _detalhe(grupo_desconto_fonte),
        # linhas cruas (inclui sublinhas de detalhamento "- ...", que os
        # buckets acima ignoram) -- ver descricoes_ml.py.
        "linhas_raw": linhas,
    }


def resumir_pedidos(numeros, pausa=1.5, tentativas=2):
    """Busca uma lista de pedidos em sequencia, reaproveitando a mesma
    conexao com o Chrome. Retorna {numero: resumo_ou_None}.

    O site as vezes falha numa busca isolada (mesmo padrao intermitente ja
    visto no scraper do SRI) -- tenta de novo antes de desistir."""
    resultados = {}
    with sync_playwright() as pw:
        context = conectar_chrome_existente(pw)
        for numero in numeros:
            resumo = None
            for tentativa in range(tentativas):
                try:
                    resumo = resumo_pedido(numero, context)
                except Exception as exc:  # pagina mudou / timeout / etc
                    resumo = {"erro": str(exc)}
                if resumo is not None and "erro" not in resumo:
                    break
                time.sleep(pausa)
            resultados[numero] = resumo
            time.sleep(pausa)
    return resultados
