"""Busca no Mercado Livre/Mercado Pago (reaproveitando o Screaper Mercado ja
existente, Playwright + Chrome com debug remoto) os valores reais de um
pedido: produto, tarifa (comissao), envio, rebate/bonus e total recebido.
A classificacao das linhas (o que e comissao/rebate/desconto etc.) mora em
classificar_ml.py, que NAO depende de playwright -- so este arquivo aqui
(a raspagem em si) precisa do Chrome/Screaper.

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

from classificar_ml import resumir_linhas  # noqa: E402


def _buscar_ml(numero, context):
    page = encontrar_aba_mercadolivre(context)
    if page is None:
        page = context.new_page()
    html = buscar_e_abrir_ultimo_resultado_ml(page, numero)
    dados = extrair_dados_detalhe_ml(html) if html else None
    return resumir_linhas(dados["linhas"], plataforma="ML", pedido_ml=dados["pedido"]) if dados else None


def _buscar_mp(numero, context):
    page = encontrar_aba_mercadopago(context)
    if page is None:
        return None
    html = buscar_e_abrir_ultimo_resultado_mp(page, numero)
    dados = extrair_dados_detalhe_mp(html) if html else None
    return resumir_linhas(dados["linhas"], plataforma="MP", pedido_ml=dados["pedido"]) if dados else None


def resumo_pedido(numero, context):
    """`numero` = SHIPPING_ID (ML, comeca com "4") ou EXTERNAL_REFERENCE
    (Mercado Pago, os demais). Retorna None se nao encontrar nada.

    Fallback (22/09/2026, pedido do Everton): quando `numero` NAO comeca
    com "4" e a busca no MP nao acha nada, tenta de novo o MESMO numero
    como SHIPPING_ID na pagina do ML antes de desistir -- alguns pedidos
    resolvidos via SRI (ver processar_full.gerar_para_processar) chegam
    aqui com um SHIPPING_ID que nao bate no padrao "comeca com 4", entao a
    escolha ML/MP so pelo prefixo deixava esses sem tentar o ML nunca.
    Quando `numero` ja comeca com "4" nao ha fallback melhor (ja e a
    tentativa certa)."""
    numero = str(numero).strip()

    if numero.startswith("4"):
        return _buscar_ml(numero, context)

    resumo = _buscar_mp(numero, context)
    return resumo if resumo is not None else _buscar_ml(numero, context)


def _barra(feitos, total, pedido, largura=30):
    """Barra de progresso na mesma linha do terminal (\\r), com o pedido que
    esta sendo buscado -- o Chrome roda fora da tela, entao o terminal e o
    unico lugar onde o Everton acompanha. Pedido do Everton, 19/09/2026."""
    cheio = largura * feitos // total
    linha = f"[{'#' * cheio}{'-' * (largura - cheio)}] {feitos * 100 // total:3d}% ({feitos}/{total}) pedido {pedido}"
    print("\r" + linha.ljust(100), end="", flush=True)


def _aviso(texto):
    """Linha propria (nao some quando a barra redesenha)."""
    print("\r" + " " * 100 + "\r" + texto, flush=True)


def resumir_pedidos(numeros, pausa=1.5, tentativas=2):
    """Busca uma lista de pedidos em sequencia, reaproveitando a mesma
    conexao com o Chrome. Retorna {numero: resumo_ou_None}.

    O site as vezes falha numa busca isolada (mesmo padrao intermitente ja
    visto no scraper do SRI) -- tenta de novo antes de desistir. Mostra
    barra de progresso, e avisa em linha propria cada erro/nao encontrado
    com o numero do pedido, mais um resumo no final."""
    resultados = {}
    numeros = list(numeros)
    total = len(numeros)
    com_erro, nao_encontrados = [], []
    with sync_playwright() as pw:
        context = conectar_chrome_existente(pw)
        for i, numero in enumerate(numeros):
            _barra(i, total, numero)
            resumo = None
            for tentativa in range(tentativas):
                try:
                    resumo = resumo_pedido(numero, context)
                except Exception as exc:  # pagina mudou / timeout / etc
                    resumo = {"erro": str(exc)}
                if resumo is not None and "erro" not in resumo:
                    break
                if tentativa + 1 < tentativas:
                    motivo = resumo["erro"][:80] if resumo else "nao encontrado"
                    _aviso(f"  pedido {numero}: tentativa {tentativa + 1}/{tentativas} falhou ({motivo}), tentando de novo...")
                time.sleep(pausa)
            if resumo is None:
                nao_encontrados.append(numero)
                _aviso(f"  NAO ENCONTRADO pedido {numero}")
            elif "erro" in resumo:
                com_erro.append(numero)
                _aviso(f"  ERRO pedido {numero}: {resumo['erro'][:120]}")
            resultados[numero] = resumo
            time.sleep(pausa)
    _barra(total, total, "-")
    ok = total - len(com_erro) - len(nao_encontrados)
    print(f"\nScraper: {ok} ok, {len(nao_encontrados)} nao encontrado(s), {len(com_erro)} com erro (de {total}).")
    if nao_encontrados:
        print("  Nao encontrados:", ", ".join(map(str, nao_encontrados)))
    if com_erro:
        print("  Com erro:", ", ".join(map(str, com_erro)))
    return resultados
