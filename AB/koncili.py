"""Koncili (components.koncili.com): acha o repasse de um dia/conta do Mercado
Livre, le o RESULTADO, baixa o AB original e confere com o calendario.

Conferencia (validada em 19/09/2026, todos os canais abaixo): soma de
GROSS + MP_FEE + FINANCING_FEE + SHIPPING_FEE do AB (colunas J:M do export de
60 colunas) == RESULTADO do Koncili == VALOR_EB do calendario.

Precisa do Python do venv do Screaper Mercado (tem playwright) e das
variaveis de ambiente KONCILI_USUARIO / KONCILI_SENHA (nunca no codigo):
    $env:KONCILI_USUARIO="email"; $env:KONCILI_SENHA="senha"
    python koncili.py ES 18.09.2026                    # so confere (AB vai pra pasta temp)
    python koncili.py ELETRO 11.09.2026 18.09.2026 --salvar   # grava no X: (nao sobrescreve)
    python koncili.py --teste                           # autoteste offline

Headless funciona no Koncili (diferente do Mercado Pago). O export do Koncili
traz as 60 colunas em ordem diferente dos AB antigos do X: -- o processar
(rodar_es.py) tem checagem de layout estrita na ordem, entao reordene por
LAYOUT_ES antes de processar um AB baixado daqui.
"""

import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "RB"))

import pandas as pd

from calendario_alvo import buscar_alvo_loja

URL = "https://components.koncili.com"
API = "https://api.koncili.com"
PASTA_AB = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB"
COLUNAS_REPASSE = ["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT"]

# conta = nome da CONTA no Koncili; pasta/prefixo = onde o AB mora no X:;
# subpasta = "antigo" so no ES (la o layout de 60 colunas em ingles fica
# separado do layout novo em portugues); loja = nome da LOJA no calendario.
CANAIS_KONCILI = {
    "ES": {"conta": "ECONTINENTAL ES", "pasta": "COLETAS ES", "prefixo": "MERCADO COLETAS ES", "subpasta": "antigo", "loja": "COLETAS ES"},
    "ELETRO": {"conta": "WEBCO ELETRO", "pasta": "ELETRODOMESTICO", "prefixo": "MERCADO ELETRODOMESTICO", "subpasta": "", "loja": "ELETRO"},
    "MOVEIS": {"conta": "WEBCO MÓVEIS", "pasta": "MOVEIS", "prefixo": "MERCADO MOVEIS", "subpasta": "", "loja": "MOVEIS"},
    "PRIME": {"conta": "WEBCOPRIME", "pasta": "PRIME", "prefixo": "MERCADO PRIME", "subpasta": "", "loja": "PRIME"},
    # Contas descobertas casando RESULTADO x calendario (valor E dia, 11-12 dias
    # cada, 19/09/2026) -- nao pelo nome. FULL: o export do Koncili (40 colunas)
    # e identico ao do X:, so a aba vira "Relatorio". CENTER/WEB/CO05 so
    # conferem: o AB do X: deles esta em outro layout (35/48 colunas), e salvar
    # o export de 60 colunas la misturaria layouts e quebraria o processar.
    "FULL": {"conta": "E CONTINENTAL", "pasta": "FULL", "prefixo": "MERCADO FULL", "subpasta": "", "loja": "FULL", "aba": "Relatorio"},
    "CENTER": {"conta": "CONTINENTALCENTER", "pasta": "CENTER", "prefixo": "MERCADO CENTER", "subpasta": "", "loja": "CENTER", "salvar": False},
    "WEB": {"conta": "WEB CONTINENTAL", "pasta": "WEB", "prefixo": "MERCADO WEB", "subpasta": "", "loja": "WEB", "salvar": False},
    "CO05": {"conta": "LTWE20240219130512", "pasta": "COLETAS CO05", "prefixo": "MERCADO COLETAS CO05", "subpasta": "", "loja": "COLETAS 05", "salvar": False},
}


def ler_credenciais():
    usuario, senha = os.environ.get("KONCILI_USUARIO"), os.environ.get("KONCILI_SENHA")
    faltando = [n for n, v in (("KONCILI_USUARIO", usuario), ("KONCILI_SENHA", senha)) if not v]
    if faltando:
        sys.exit(f"ERRO: defina a(s) variavel(is) de ambiente: {', '.join(faltando)}")
    return usuario, senha


def login(page, usuario, senha):
    """Nao usa networkidle: a pagina mantem requisicoes abertas e nunca assenta."""
    from playwright.sync_api import TimeoutError as PlaywrightTimeout

    page.goto(f"{URL}/auth", wait_until="domcontentloaded")
    page.wait_for_selector("input#user", timeout=20000)
    page.fill("input#user", usuario)
    page.fill("input#password", senha)
    page.click("button[type=submit]")
    try:
        page.wait_for_url(lambda url: "/auth" not in url, timeout=20000)
    except PlaywrightTimeout:
        raise RuntimeError("login nao confirmado: continua em /auth (usuario/senha incorretos?)")


def cabecalhos_api(page):
    """A API do Koncili autentica pelo header `gumgatoken` que a propria pagina
    manda -- captura o da primeira busca da tela de conciliacoes."""
    with page.expect_request(lambda r: "view/conciliation/search" in r.url, timeout=30000) as info:
        page.goto(f"{URL}/conciliation", wait_until="domcontentloaded")
    h = info.value.headers
    return {k: v for k, v in h.items() if k.lower() in ("gumgatoken", "content-type", "accept", "accept-language")}


def achar_repasse(req, hdr, conta, data_iso, max_paginas=10):
    """Repasse mais recente da `conta` cujo periodo comeca em `data_iso`
    (AAAA-MM-DD). A lista vem da mais nova pra mais antiga."""
    for pg in range(1, max_paginas + 1):
        corpo = {"analyzed": None, "status": None, "channelName": ["MERCADOLIVRE"], "page": pg, "pageSize": 100,
                 "orderByField": "PROCESS_CREATED_DATE", "orderByDir": "desc"}
        d = req.post(f"{API}/reconciliation/api/view/conciliation/search", data=json.dumps(corpo), headers=hdr).json()
        for i in d["values"]:
            if i["accountName"] == conta and i["initDate"][:10] == data_iso:
                return i
        if pg * 100 >= d["count"]:
            break
    return None


def resultado(req, hdr, id_repasse):
    r = req.get(f"{API}/reconciliation/api/reconciliation/view/{id_repasse}/summary", headers=hdr).json()
    return r["kpis"]["result"]


def baixar_ab(page, id_repasse, destino):
    """Botao 'Visualizar planilha' (#btn-view-spreadsheet) da tela do repasse."""
    page.goto(f"{URL}/reconciliation/{id_repasse}", wait_until="domcontentloaded")
    botao = page.locator('xpath=//*[@id="btn-view-spreadsheet"]')
    botao.wait_for(state="visible", timeout=30000)
    with page.expect_download(timeout=60000) as dl:
        botao.click()
    dl.value.save_as(destino)


def soma_ab(caminho):
    df = pd.read_excel(caminho, sheet_name=0)
    return sum(pd.to_numeric(df[c], errors="coerce").sum() for c in COLUNAS_REPASSE)


def caminho_destino(canal, data):
    """`data` = DD.MM.AAAA. Mesma convencao de pasta/nome do X:."""
    c = CANAIS_KONCILI[canal]
    dia, mes, ano = data.split(".")
    partes = [PASTA_AB, c["pasta"], ano, f"{mes}.{ano}"] + ([c["subpasta"]] if c["subpasta"] else [])
    return os.path.join(*partes, f"{c['prefixo']} {data}.xlsx")


def conferir(page, req, hdr, canal, data, salvar=False):
    c = CANAIS_KONCILI[canal]
    dia, mes, ano = data.split(".")
    rep = achar_repasse(req, hdr, c["conta"], f"{ano}-{mes}-{dia}")
    if rep is None:
        print(f"  {canal} {data}: sem repasse no Koncili (conta {c['conta']})")
        return None

    destino = caminho_destino(canal, data)
    ja_existe = os.path.exists(destino)
    tmp = os.path.join(tempfile.gettempdir(), f"koncili_{rep['id']}.xlsx")
    baixar_ab(page, rep["id"], tmp)

    ab, kon = soma_ab(tmp), resultado(req, hdr, rep["id"])
    cal = buscar_alvo_loja(c["loja"], f"{dia}/{mes}/{ano}")
    fmt = lambda v: "sem linha" if v is None else f"{v:,.2f}"
    ok_kon = abs(ab - kon) < 0.01
    ok_cal = None if cal is None else abs(ab - cal) < 0.01
    print(f"  {canal} {data}: repasse {rep['id']} | AB {fmt(ab)} | Koncili {fmt(kon)} {'OK' if ok_kon else 'DIFERE'} | calendario {fmt(cal)} {'' if ok_cal is None else ('OK' if ok_cal else 'DIFERE')}")

    if salvar:
        if not c.get("salvar", True):
            print(f"    {canal} so confere: layout do X: e diferente do export do Koncili, nao gravei")
        elif ja_existe:
            print(f"    ja existe no X:, nao sobrescrevi: {destino}")
        else:
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            if c.get("aba"):
                import openpyxl

                wb = openpyxl.load_workbook(tmp)
                wb.active.title = c["aba"]
                wb.save(tmp)
            shutil.move(tmp, destino)  # os.replace nao atravessa unidades (C: -> X:)
            print(f"    salvo: {destino}")
    return {"repasse": rep["id"], "ab": ab, "koncili": kon, "calendario": cal, "ok_koncili": ok_kon, "ok_calendario": ok_cal}


def main(argv):
    salvar = "--salvar" in argv
    args = [a for a in argv if not a.startswith("--")]
    if len(args) < 2 or args[0].upper() not in CANAIS_KONCILI:
        sys.exit(f"uso: python koncili.py <{'|'.join(CANAIS_KONCILI)}> DD.MM.AAAA [DD.MM.AAAA ...] [--salvar]")
    canal, datas = args[0].upper(), args[1:]
    usuario, senha = ler_credenciais()

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_context(accept_downloads=True).new_page()
        try:
            login(page, usuario, senha)
            hdr = cabecalhos_api(page)
            for i, data in enumerate(datas, 1):
                print(f"[{i}/{len(datas)}] {canal} {data}", flush=True)
                try:
                    conferir(page, page.context.request, hdr, canal, data, salvar)
                except Exception as exc:  # um dia com erro nao derruba os outros
                    print(f"  ERRO {canal} {data}: {str(exc)[:150]}", flush=True)
        finally:
            browser.close()


def _teste():
    """Offline: caminhos do X: e a soma do AB (nada de rede/login)."""
    assert caminho_destino("ES", "18.09.2026").endswith(r"COLETAS ES\2026\09.2026\antigo\MERCADO COLETAS ES 18.09.2026.xlsx")
    assert caminho_destino("MOVEIS", "11.09.2026").endswith(r"MOVEIS\2026\09.2026\MERCADO MOVEIS 11.09.2026.xlsx")
    tmp = os.path.join(tempfile.gettempdir(), "koncili_teste.xlsx")
    pd.DataFrame({"GROSS_AMOUNT": [100.0, 50.0], "MP_FEE_AMOUNT": [-10.0, -5.0], "FINANCING_FEE_AMOUNT": [-2.0, 0.0],
                  "SHIPPING_FEE_AMOUNT": [-8.0, -1.0], "COUPON_AMOUNT": [3.0, 3.0]}).to_excel(tmp, index=False)
    assert abs(soma_ab(tmp) - 124.0) < 0.001, "cupom nao entra na soma"
    os.remove(tmp)
    print("koncili.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
