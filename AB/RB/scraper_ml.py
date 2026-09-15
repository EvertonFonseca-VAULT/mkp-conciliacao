"""Busca no Mercado Livre/Mercado Pago (reaproveitando o Screaper Mercado ja
existente, Playwright + Chrome com debug remoto) os valores reais de um
pedido: produto, tarifa (comissao), envio, rebate/bonus e total recebido.

Precisa do Chrome aberto com debug remoto (porta 9222), ja logado em
mercadolivre.com.br e mercadopago.com.br -- ver
Manuais/Documentação Screaper MP e ML.md.

Roda com o Python do venv do Screaper Mercado (tem playwright instalado):
    "C:\\Users\\everton.fonseca\\Desktop\\Projetos\\Screaper Mercado\\.venv\\Scripts\\python.exe" ...
"""

import sys
import time

CAMINHO_SCRAPER = r"C:\Users\everton.fonseca\Desktop\Projetos\Screaper Mercado"
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
    "rebate": ["desconto", "cancelamento", "estorno", "bônus", "bonus"],
    "total": ["total"],
}


def _valor(linhas, chaves, inicio=False):
    for label, valor in linhas:
        if not label:
            continue
        texto = label.strip().lower()
        if texto.startswith("-"):
            continue  # sublinha de detalhamento, nao o total do grupo
        bateu = any(texto.startswith(c) for c in chaves) if inicio else any(c in texto for c in chaves)
        if bateu:
            return valor
    return None


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
    return {
        "plataforma": plataforma,
        "pedido_ml": dados["pedido"],
        "produto": _valor(linhas, _ROTULOS["produto"]),
        "comissao": _valor(linhas, _ROTULOS["comissao"]),
        "envio": _valor(linhas, _ROTULOS["envio"]),
        "rebate": _valor(linhas, _ROTULOS["rebate"]),
        # "total" precisa bater no INICIO do rotulo -- senao "Total" colide
        # com o "l" final de "Tarifa de venda total", que aparece antes na
        # lista e tambem contem a substring "total".
        "total": _valor(linhas, _ROTULOS["total"], inicio=True),
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
