"""Koncili (components.koncili.com): acha o repasse de um dia/conta do Mercado
Livre, le o RESULTADO, baixa o AB original e confere com o calendario.

Conferencia (validada em 19/09/2026, todos os canais abaixo): soma de
GROSS + MP_FEE + FINANCING_FEE + SHIPPING_FEE + EFFECTIVE_COUPON do AB (colunas
J:M + EFFECTIVE_COUPON_AMOUNT do export de 60 colunas) == RESULTADO do Koncili
== VALOR_EB do calendario. EFFECTIVE_COUPON_AMOUNT entrou em 13/10/2026: sem
ela, PRIME e WEB (lojas com cupom proprio) ficavam DIFERE -- o valor do cupom
(negativo) e exatamente a diferenca (confirmado em PRIME 01,02,05,06.10 e
WEB 01,02.10).

Precisa do Python do venv do Screaper Mercado (tem playwright) e das
credenciais KONCILI_USUARIO / KONCILI_SENHA -- por variavel de ambiente ou
gravadas no cofre do Windows (ver COMUM/cofre_senhas.py; ambas nunca no codigo):
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

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "COMUM"))

import pandas as pd
from datetime import datetime
from calendario_alvo import buscar_alvo_loja, LOJA_POR_CANAL

URL = "https://components.koncili.com"
API = "https://api.koncili.com"
PASTA_AB = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB"
COLUNAS_REPASSE = ["GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "EFFECTIVE_COUPON_AMOUNT"]

# Layouts de 35 colunas (em ingles) que o X: usa em CENTER e WEB: selecao/ordem
# das 60 do export do Koncili. WEB tem REFUND_ID, que o Koncili nao exporta
# (a coluna sai vazia).
_BASE_35 = ["RELEASE_DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "RECORD_TYPE", "DESCRIPTION", "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT",
            "GROSS_AMOUNT", "SELLER_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT",
            "COUPON_AMOUNT", "INSTALLMENTS", "PAYMENT_METHOD", "TAX_DETAIL", "TAX_AMOUNT_TELCO", "APPROVAL_DATE", "POS_ID",
            "POS_NAME", "EXTERNAL_POS_ID", "STORE_ID", "STORE_NAME", "EXTERNAL_STORE_ID"]
COLUNAS_CENTER = _BASE_35 + ["TAXES_DISAGGREGATED", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID", "METADATA",
                             "EFFECTIVE_COUPON_AMOUNT", "POI_ID", "CARD_INITIAL_NUMBER", "PRODUCT_SKU"]
COLUNAS_WEB = _BASE_35 + ["CURRENCY", "TAXES_DISAGGREGATED", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID", "METADATA",
                          "REFUND_ID", "EFFECTIVE_COUPON_AMOUNT", "POI_ID"]

# conta = nome da CONTA no Koncili; pasta/prefixo = onde o AB mora no X:;
# subpasta = "antigo" so no ES (la o layout de 60 colunas em ingles fica
# separado do layout novo em portugues); loja = nome da LOJA no calendario
# (preenchida abaixo a partir de canais.py -- CONCILIACAO/canais.py e a
# FONTE UNICA, ver revisao de arquitetura de 14/10/2026).
CANAIS_KONCILI = {
    "ES": {"conta": "ECONTINENTAL ES", "pasta": "COLETAS ES", "prefixo": "MERCADO COLETAS ES", "subpasta": "antigo"},
    "ELETRO": {"conta": "WEBCO ELETRO", "pasta": "ELETRODOMESTICO", "prefixo": "MERCADO ELETRODOMESTICO", "subpasta": ""},
    "MOVEIS": {"conta": "WEBCO MÓVEIS", "pasta": "MOVEIS", "prefixo": "MERCADO MOVEIS", "subpasta": ""},
    "PRIME": {"conta": "WEBCOPRIME", "pasta": "PRIME", "prefixo": "MERCADO PRIME", "subpasta": ""},
    # Contas descobertas casando RESULTADO x calendario (valor E dia, 11-12 dias
    # cada, 19/09/2026) -- nao pelo nome. FULL: o export do Koncili (40 colunas)
    # e identico ao do X:, so a aba vira "Relatorio". CENTER/WEB: gravam no
    # layout de 35 colunas do X:.
    "FULL": {"conta": "E CONTINENTAL", "pasta": "FULL", "prefixo": "MERCADO FULL", "subpasta": "", "aba": "Relatorio"},
    "CENTER": {"conta": "CONTINENTALCENTER", "pasta": "CENTER", "prefixo": "MERCADO CENTER", "subpasta": "", "aba": "Plan1", "colunas": COLUNAS_CENTER},
    "WEB": {"conta": "WEB CONTINENTAL", "pasta": "WEB", "prefixo": "MERCADO WEB", "subpasta": "", "aba": "Planilha1", "colunas": COLUNAS_WEB},
    # CO05: achado 23/09/2026 -- o X: usa o MESMO layout de 60 colunas do ES
    # (so ordem diferente, aba "Sheet0"), nao "48 colunas em portugues" como
    # se pensava antes. Sem "aba"/"colunas" (igual ES): salva a "Sheet0" do
    # jeito que o Koncili exporta, sem reajuste -- processar_es.carregar_ab
    # ja reordena as 60 colunas sozinho. "salvar" removido (era False).
    "COLETAS_CO05": {"conta": "LTWE20240219130512", "pasta": "COLETAS CO05", "prefixo": "MERCADO COLETAS CO05", "subpasta": ""},
}

# chave aqui -> chave de LOJA_POR_CANAL (COMUM/calendario_alvo.py), so onde
# o nome diverge.
_ALIAS_CANAL_CANONICO = {"COLETAS_CO05": "CO05"}

def _loja_canonica(canal):
    canonico = _ALIAS_CANAL_CANONICO.get(canal, canal)
    return LOJA_POR_CANAL.get(canonico)

for _canal, _cfg in CANAIS_KONCILI.items():
    _cfg["loja"] = _loja_canonica(_canal)
del _canal, _cfg


def _para_iso(data):
    """DD.MM.AAAA -> AAAA-MM-DD (valida a data)."""
    return datetime.strptime(data, "%d.%m.%Y").strftime("%Y-%m-%d")

def _de_iso(iso):
    a, m, d = iso.split("-")
    return f"{d}.{m}.{a}"


def ler_credenciais():
    """Credenciais do Koncili -- nunca no codigo. Prioridade: variavel de
    ambiente (mesmo padrao do SRI em Screaper Mercado/Pagina.py), senao
    Gerenciador de Credenciais do Windows (keyring, servico "MARKETPLACE_KONCILI").
    Para gravar no cofre: COMUM/cofre_senhas.py (ou cofre_senhas.bat).
        $env:KONCILI_USUARIO="email"; $env:KONCILI_SENHA="senha"
    """
    usuario, senha = os.environ.get("KONCILI_USUARIO"), os.environ.get("KONCILI_SENHA")
    if not usuario or not senha:
        import keyring
        usuario = usuario or keyring.get_password("MARKETPLACE_KONCILI", "usuario")
        senha = senha or keyring.get_password("MARKETPLACE_KONCILI", "senha")
    faltando = [n for n, v in (("KONCILI_USUARIO", usuario), ("KONCILI_SENHA", senha)) if not v]
    if faltando:
        sys.exit(f"ERRO: defina a(s) variavel(is) de ambiente ou grave no cofre (MARKETPLACE_KONCILI): {', '.join(faltando)}")
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


def listar_repasses(req, hdr, conta, ini_iso, fim_iso, max_paginas=30):
    """Repasses da `conta` cujo periodo comeca entre ini_iso e fim_iso
    (AAAA-MM-DD, inclusive). Um por dia: o mais recente, se houve reprocesso."""
    achados = {}
    for pg in range(1, max_paginas + 1):
        corpo = {"analyzed": None, "status": None, "channelName": ["MERCADOLIVRE"], "page": pg, "pageSize": 100,
                 "orderByField": "PROCESS_CREATED_DATE", "orderByDir": "desc"}
        d = req.post(f"{API}/reconciliation/api/view/conciliation/search", data=json.dumps(corpo), headers=hdr).json()
        for i in d["values"]:
            dia = i["initDate"][:10]
            if i["accountName"] == conta and ini_iso <= dia <= fim_iso:
                achados.setdefault(dia, i)  # lista vem da mais nova pra mais antiga
        if pg * 100 >= d["count"]:
            break
        # pagina inteira ja anterior ao inicio: os proximos so ficam mais velhos
        if d["values"] and all(i["initDate"][:10] < ini_iso for i in d["values"]):
            break
    return [achados[k] for k in sorted(achados)]

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


def _ajustar_layout(caminho, c):
    """Deixa o export do Koncili no layout do X:: renomeia a aba e, se o canal
    tem `colunas`, mantem so essas (nessa ordem; ausente = coluna vazia).
    Copia os valores como vieram (tipos e IDs intactos), sem passar por pandas.

    O export do WEB/CENTER vem com 2 abas (achado 13/10/2026, WEB 07.10.2026:
    aba "7" com os dados certos + aba "8", sobra do dia seguinte bundlada no
    mesmo arquivo) -- sem remover a sobra, ela fica na posicao 0 depois que a
    aba original e removida, e soma_ab/leitura posterior (sheet_name=0) le a
    aba errada. Por isso remove TODAS as outras abas, nao so `ws`."""
    import openpyxl

    wb = openpyxl.load_workbook(caminho)
    ws = wb.active
    if c.get("colunas"):
        linhas = ws.iter_rows(values_only=True)
        cab = list(next(linhas))
        pos = [cab.index(n) if n in cab else None for n in c["colunas"]]
        novo = wb.create_sheet(c["aba"])
        novo.append(c["colunas"])
        for lin in linhas:
            novo.append([None if p is None else lin[p] for p in pos])
        for extra in wb.sheetnames:
            if extra != c["aba"]:
                wb.remove(wb[extra])
    else:
        ws.title = c["aba"]
        for extra in list(wb.sheetnames):
            if extra != c["aba"]:
                wb.remove(wb[extra])
    wb.save(caminho)


def conferir(page, req, hdr, canal, data, salvar=False, rep=None):
    c = CANAIS_KONCILI[canal]
    dia, mes, ano = data.split(".")
    if rep is None:
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
                _ajustar_layout(tmp, c)
            shutil.move(tmp, destino)  # os.replace nao atravessa unidades (C: -> X:)
            print(f"    salvo: {destino}")
    return {"repasse": rep["id"], "ab": ab, "koncili": kon, "calendario": cal, "ok_koncili": ok_kon, "ok_calendario": ok_cal}


def main(argv):
    salvar = "--salvar" in argv
    intervalo = "--intervalo" in argv
    args = [a for a in argv if not a.startswith("--")]
    ok_args = len(args) == 3 if intervalo else len(args) >= 2
    if not ok_args or args[0].upper() not in CANAIS_KONCILI:
        sys.exit(f"uso: python koncili.py <{'|'.join(CANAIS_KONCILI)}> DD.MM.AAAA [DD.MM.AAAA ...] [--salvar]\n"
                 f"     python koncili.py <canal> DATA_INICIO DATA_FIM --intervalo [--salvar]")
    canal = args[0].upper()
    if intervalo:
        try:
            ini_iso, fim_iso = _para_iso(args[1]), _para_iso(args[2])
        except ValueError:
            sys.exit("ERRO: datas devem estar no formato DD.MM.AAAA")
        if ini_iso > fim_iso:
            sys.exit("ERRO: data inicial maior que a final")
    usuario, senha = ler_credenciais()

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_context(accept_downloads=True).new_page()
        try:
            login(page, usuario, senha)
            hdr = cabecalhos_api(page)
            if intervalo:
                reps = listar_repasses(page.context.request, hdr, CANAIS_KONCILI[canal]["conta"], ini_iso, fim_iso)
                print(f"{canal}: {len(reps)} repasse(s) entre {args[1]} e {args[2]}", flush=True)
                trabalho = [(_de_iso(r["initDate"][:10]), r) for r in reps]
            else:
                trabalho = [(d, None) for d in args[1:]]
            for i, (data, rep) in enumerate(trabalho, 1):
                print(f"[{i}/{len(trabalho)}] {canal} {data}", flush=True)
                try:
                    conferir(page, page.context.request, hdr, canal, data, salvar, rep)
                except Exception as exc:  # um dia com erro nao derruba os outros
                    print(f"  ERRO {canal} {data}: {str(exc)[:150]}", flush=True)
        finally:
            browser.close()


def _teste():
    """Offline: caminhos do X: e a soma do AB (nada de rede/login)."""
    assert caminho_destino("ES", "18.09.2026").endswith(r"COLETAS ES\2026\09.2026\antigo\MERCADO COLETAS ES 18.09.2026.xlsx")
    import openpyxl

    t = os.path.join(tempfile.gettempdir(), "koncili_layout.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "7"
    wb.active.append(["ZZ", "SOURCE_ID", "GROSS_AMOUNT"] + [c for c in COLUNAS_WEB if c not in ("SOURCE_ID", "GROSS_AMOUNT")][:0])
    wb.active.append(["x", 123, 9.5])
    wb.create_sheet("8").append(["lixo do dia seguinte, bundlado no mesmo export"])  # ver docstring de _ajustar_layout
    wb.save(t)
    _ajustar_layout(t, CANAIS_KONCILI["WEB"])
    wb2 = openpyxl.load_workbook(t)
    assert wb2.sheetnames == ["Planilha1"], "aba extra do dia seguinte nao foi removida"
    ws = wb2.active
    assert ws.title == "Planilha1" and [c.value for c in ws[1]] == COLUNAS_WEB and len(COLUNAS_WEB) == 35
    assert ws.cell(2, 2).value == 123 and ws.cell(2, 8).value == 9.5 and ws.cell(2, 33).value is None  # REFUND_ID vazio
    assert soma_ab(t) == 9.5  # sheet_name=0 tem que ler a Planilha1, nao a aba extra
    os.remove(t)
    assert len(COLUNAS_CENTER) == 35
    assert caminho_destino("MOVEIS", "11.09.2026").endswith(r"MOVEIS\2026\09.2026\MERCADO MOVEIS 11.09.2026.xlsx")
    tmp = os.path.join(tempfile.gettempdir(), "koncili_teste.xlsx")
    pd.DataFrame({"GROSS_AMOUNT": [100.0, 50.0], "MP_FEE_AMOUNT": [-10.0, -5.0], "FINANCING_FEE_AMOUNT": [-2.0, 0.0],
                  "SHIPPING_FEE_AMOUNT": [-8.0, -1.0], "COUPON_AMOUNT": [3.0, 3.0],
                  "EFFECTIVE_COUPON_AMOUNT": [-1.0, -1.0]}).to_excel(tmp, index=False)
    assert abs(soma_ab(tmp) - 122.0) < 0.001, "COUPON_AMOUNT bruto entrou na soma, ou EFFECTIVE_COUPON_AMOUNT nao entrou"
    os.remove(tmp)

    # Credenciais: env var tem prioridade; sem env, cai no cofre (keyring)
    import types
    senha_no_cofre = {"usuario": "fulano", "senha": "123"}
    fake_keyring = types.SimpleNamespace(
        get_password=lambda servico, chave: senha_no_cofre.get(chave) if servico == "MARKETPLACE_KONCILI" else None
    )
    sys.modules["keyring"] = fake_keyring
    env_antigo = {k: os.environ.pop(k, None) for k in ("KONCILI_USUARIO", "KONCILI_SENHA")}
    try:
        assert ler_credenciais() == ("fulano", "123"), "nao caiu no cofre quando faltou variavel de ambiente"
        os.environ["KONCILI_USUARIO"] = "var-ambiente"
        os.environ["KONCILI_SENHA"] = "senha-ambiente"
        assert ler_credenciais() == ("var-ambiente", "senha-ambiente"), "variavel de ambiente devia ter prioridade sobre o cofre"
    finally:
        del sys.modules["keyring"]
        for k, v in env_antigo.items():
            if v is not None:
                os.environ[k] = v
            else:
                os.environ.pop(k, None)

    print("koncili.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
