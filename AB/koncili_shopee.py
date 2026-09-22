"""Extrai e empilha os lancamentos (repasses) do Shopee no Koncili --
processo a parte do Mercado (koncili.py): aqui ainda nao ha calendario/TOTVS
pra cruzar, so extrai o que o Koncili tem disponivel no periodo pedido e
empilha por conta.

Cada conta vai pra sua propria pasta em
X:\\...\\TEMP\\CONCILIACAO\\SHOPEE\\SHOPEE_<CANAL>\\: um arquivo por
lancamento (SHOPEE_<CANAL>_DDMMAA.xlsx, DDMMAA = inicio do periodo, sem
pontos -- convencao diferente do resto do projeto, pedido do Everton,
22/09/2026) + o empilhado (SHOPEE_<CANAL>_EMPILHADO.xlsx, todos os
lancamentos concatenados, sem mexer nas colunas).

Nomes de conta: o que o Everton copiou do seletor do Koncili (22/09/2026)
nao batia exatamente com o que a API retorna em accountName -- confirmado
com --descobrir no mesmo dia: ES e WEB usam travessao "-" (nao hifen "-")
e/ou espacos diferentes dos que foram passados. ES e WEB ja corrigidos com
o texto exato visto na API; LG, CENTER e PRIME ainda usam o texto original
do seletor e podem estar com a mesma pegadinha -- conferir com --descobrir
antes de confiar neles. O channelName NAO importa o case ("Shopee"/
"SHOPEE"/"shopee" dao o mesmo resultado, confirmado com dado real) -- entao
"nenhum lancamento achado" e sinal de nome de conta errado, nao de channel.

Uso (mesmas variaveis de ambiente e venv do koncili.py -- KONCILI_USUARIO/
KONCILI_SENHA, Python do venv do Screaper Mercado):
    python koncili_shopee.py                          # as 5 contas, julho e agosto/2026
    python koncili_shopee.py --conta=ES                # so uma conta
    python koncili_shopee.py --desde=2026-07-01 --ate=2026-08-31
    python koncili_shopee.py --descobrir               # lista as accountName reais da API
    python koncili_shopee.py --teste                   # autoteste offline
"""

import glob
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "RB"))

import pandas as pd

from koncili import API, baixar_ab, cabecalhos_api, ler_credenciais, login

PASTA_SHOPEE = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\TEMP\CONCILIACAO\SHOPEE"

# canal -> nome da CONTA no Koncili. ES e WEB confirmados com o texto exato
# da API (--descobrir, 22/09/2026); LG, CENTER e PRIME ainda sao o texto que
# o Everton copiou do seletor -- ver docstring do modulo.
CANAIS_SHOPEE = {
    "LG": "SHOPEE-LG_3P",
    "ES": "Shopee – 1059412251_ES_1P",
    "CENTER": "Shopee - 889973435_SC_1P",
    "PRIME": "SHOPEE-WEBCO PRIME_3P",
    "WEB": "SHOPEE - WEBCONTINENTAL MARKETPLACE_3P",
}

# Pedido do Everton, 22/09/2026: "extrai julho e agosto" (hoje e 22/09/2026).
DESDE_PADRAO = "2026-07-01"
ATE_PADRAO = "2026-08-31"


def pasta_conta(canal):
    return os.path.join(PASTA_SHOPEE, f"SHOPEE_{canal}")


def caminho_periodo(canal, data_iso):
    """data_iso = AAAA-MM-DD (initDate do repasse) -> nome DDMMAA."""
    ano, mes, dia = data_iso.split("-")
    return os.path.join(pasta_conta(canal), f"SHOPEE_{canal}_{dia}{mes}{ano[2:]}.xlsx")


def caminho_empilhado(canal):
    return os.path.join(pasta_conta(canal), f"SHOPEE_{canal}_EMPILHADO.xlsx")


def achar_repasses(req, hdr, conta, desde_iso, ate_iso, channel="Shopee", max_paginas=30):
    """Todos os repasses da `conta` com initDate dentro de [desde_iso,
    ate_iso] (AAAA-MM-DD). Sem early-stop por data (a ordenacao e por
    PROCESS_CREATED_DATE, nao por initDate -- nao da pra confiar que parar
    na 1a data fora do periodo nao perde nada); so para quando acaba a
    paginacao ou bate max_paginas."""
    achados = []
    for pg in range(1, max_paginas + 1):
        corpo = {"analyzed": None, "status": None, "channelName": [channel], "page": pg, "pageSize": 100,
                 "orderByField": "PROCESS_CREATED_DATE", "orderByDir": "desc"}
        d = req.post(f"{API}/reconciliation/api/view/conciliation/search", data=json.dumps(corpo), headers=hdr).json()
        for i in d["values"]:
            if i["accountName"] == conta and desde_iso <= i["initDate"][:10] <= ate_iso:
                achados.append(i)
        if pg * 100 >= d["count"]:
            break
    return achados


def extrair_conta(page, req, hdr, canal, desde_iso, ate_iso, channel="Shopee"):
    """Baixa (sem sobrescrever) os periodos ainda nao salvos da conta.
    Devolve a lista dos arquivos baixados agora."""
    conta = CANAIS_SHOPEE[canal]
    repasses = achar_repasses(req, hdr, conta, desde_iso, ate_iso, channel)
    if not repasses:
        print(f"  {canal} ({conta}): nenhum lancamento achado no periodo -- conferir nome da conta / --channel")
        return []
    pasta = pasta_conta(canal)
    os.makedirs(pasta, exist_ok=True)
    baixados = []
    for rep in repasses:
        data_iso = rep["initDate"][:10]
        destino = caminho_periodo(canal, data_iso)
        if os.path.exists(destino):
            print(f"  {canal} {data_iso}: ja existe, nao sobrescrevi ({destino})")
            continue
        baixar_ab(page, rep["id"], destino)
        print(f"  {canal} {data_iso}: baixado -> {destino}")
        baixados.append(destino)
    return baixados


def empilhar_conta(canal):
    """Concatena todos os SHOPEE_<canal>_DDMMAA.xlsx da pasta da conta num
    unico arquivo (linha a linha, sem cruzar com TOTVS/calendario -- so
    empilha o que o Koncili deu). Coluna ORIGEM (nome do arquivo) em cada
    linha, pra rastrear de qual lancamento veio. None se a pasta nao tem
    nenhum periodo baixado ainda."""
    pasta = pasta_conta(canal)
    arquivos = sorted(
        a for a in glob.glob(os.path.join(pasta, f"SHOPEE_{canal}_*.xlsx"))
        if not a.endswith("_EMPILHADO.xlsx")
    )
    if not arquivos:
        return None
    partes = []
    for a in arquivos:
        df = pd.read_excel(a, sheet_name=0)
        df.insert(0, "ORIGEM", os.path.basename(a))
        partes.append(df)
    destino = caminho_empilhado(canal)
    pd.concat(partes, ignore_index=True).to_excel(destino, index=False)
    return destino


def descobrir_contas(req, hdr, channel="Shopee", max_paginas=30):
    """Lista todas as accountName distintas que o channelName acha (case
    nao importa -- confirmado com dado real 22/09/2026: "Shopee"/"SHOPEE"/
    "shopee" deram os mesmos 610 registros). O problema real que fez as 5
    contas darem "nenhum lancamento" nao era o channel, era o nome exato da
    conta (traco "-" x "-" en-dash, espacos) -- por isso aqui pega tudo que
    existe de uma vez, pra copiar o texto exato certo pra CANAIS_SHOPEE."""
    vistas = set()
    for pg in range(1, max_paginas + 1):
        corpo = {"analyzed": None, "status": None, "channelName": [channel], "page": pg, "pageSize": 100,
                 "orderByField": "PROCESS_CREATED_DATE", "orderByDir": "desc"}
        d = req.post(f"{API}/reconciliation/api/view/conciliation/search", data=json.dumps(corpo), headers=hdr).json()
        vistas.update(i["accountName"] for i in d["values"])
        if pg * 100 >= d["count"]:
            break
    for nome in sorted(vistas):
        print(repr(nome))


def main(argv):
    contas = [a.split("=", 1)[1].upper() for a in argv if a.startswith("--conta=")] or list(CANAIS_SHOPEE)
    desde = next((a.split("=", 1)[1] for a in argv if a.startswith("--desde=")), DESDE_PADRAO)
    ate = next((a.split("=", 1)[1] for a in argv if a.startswith("--ate=")), ATE_PADRAO)
    channel = next((a.split("=", 1)[1] for a in argv if a.startswith("--channel=")), "Shopee")
    usuario, senha = ler_credenciais()

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_context(accept_downloads=True).new_page()
        try:
            login(page, usuario, senha)
            hdr = cabecalhos_api(page)
            if "--descobrir" in argv:
                descobrir_contas(page.context.request, hdr, channel)
                return
            for canal in contas:
                print(f"=== {canal} ({CANAIS_SHOPEE[canal]}) ===", flush=True)
                try:
                    extrair_conta(page, page.context.request, hdr, canal, desde, ate, channel)
                except Exception as exc:  # 1 conta com erro nao derruba as outras
                    print(f"  ERRO {canal}: {str(exc)[:200]}", flush=True)
                caminho = empilhar_conta(canal)
                print(f"  empilhado: {caminho or '(nada pra empilhar ainda)'}")
        finally:
            browser.close()


def _teste():
    """Offline: nomes de pasta/arquivo e o empilhamento (sem rede/login)."""
    assert pasta_conta("LG").endswith(r"SHOPEE\SHOPEE_LG")
    assert caminho_periodo("ES", "2026-07-15").endswith(r"SHOPEE_ES\SHOPEE_ES_150726.xlsx")
    assert caminho_empilhado("WEB").endswith(r"SHOPEE_WEB\SHOPEE_WEB_EMPILHADO.xlsx")

    import tempfile

    global PASTA_SHOPEE
    anterior = PASTA_SHOPEE
    with tempfile.TemporaryDirectory() as tmp:
        PASTA_SHOPEE = tmp
        assert empilhar_conta("LG") is None  # pasta nem existe ainda
        os.makedirs(pasta_conta("LG"))
        pd.DataFrame({"A": [1, 2]}).to_excel(caminho_periodo("LG", "2026-07-01"), index=False)
        pd.DataFrame({"A": [3]}).to_excel(caminho_periodo("LG", "2026-08-01"), index=False)
        destino = empilhar_conta("LG")
        assert destino == caminho_empilhado("LG")
        df = pd.read_excel(destino)
        assert list(df["A"]) == [1, 2, 3], df
        assert df["ORIGEM"].nunique() == 2, df
        # rodar nao deve incluir o proprio EMPILHADO na proxima empilhada
        destino2 = empilhar_conta("LG")
        assert len(pd.read_excel(destino2)) == 3
    PASTA_SHOPEE = anterior
    print("koncili_shopee.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
