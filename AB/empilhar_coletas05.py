"""Empilhado do COLETAS 05 (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_web.py:
Nao carrega o TOTVS. O "Nº QUE ESTA NO TOTVS" (coluna U) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\COLETAS_CO05\\PROCESSAR_MERCADO COLETAS_CO05 DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * Mesmo layout de 25 colunas do WEB (A..Y). O que muda, e o script LE DO PROPRIO ARQUIVO (ver _convencoes):
      - VALOR (W): formula SUM(I,J,L,K) (GROSS + MP_FEE + SHIPPING_FEE + FINANCING_FEE), aceita tambem SUM(I:K) e SUM(I+J+K);
      - SEQUENCIA (X): NUMERO FIXO = numero da linha (X2 = 2), nao formula; se a ultima linha do arquivo tiver formula,
        segue formula (X<anterior>+1); o deslocamento (linha x sequencia) tambem e lido da ultima linha;
      - CHAVE_PRIMARIA (Y) = CONCATENATE(X,"_COLETAS_CO05_MERCADO").
  * INSTALLMENTS (O) e gravado como TEXTO, igual ao exemplo do Coletas 05 (no WEB e numero).
  * nao mexe em calendario/tabela a nao ser que existam no arquivo.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_coletas05.py 11.09.2026                        # monta e confere; NAO grava nada
    python empilhar_coletas05.py 11.09.2026 14.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_coletas05.py --teste                           # teste offline
Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).
"""

import datetime
import glob
import os
import re
import sys
import xml.etree.ElementTree as ET
import zlib
import zipfile
from xml.sax.saxutils import escape

import pandas as pd

_AQUI = os.path.dirname(os.path.abspath(__file__))

# ---- AJUSTAR ----------------------------------------------------------------------------
PASTA_EMPILHADO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\ARQUIVOS EMPILHADOS DE BAIXAS\ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS"
ARQUIVO = "COLETAS_CO05_MERCADO_2025.xlsx"
CANAL = "COLETAS_CO05"            # nome do canal no koncili.CANAIS_KONCILI e no arquivo Para Processar
# (ou passe o caminho completo na linha de comando: --arquivo="X:\...\arquivo.xlsx")
ABA = "Base"                # nome da aba de dados do empilhado (se o arquivo so tiver 1 aba, usa ela)
PASTA_PROCESSAR = r"C:\RODAR TOTVS\PROCESSAR\COLETAS CO05"   # onde ficam os PROCESSAR_MERCADO COLETAS_CO05 DD.MM.AAAA.xlsx
SUFIXO = "COLETAS_CO05_MERCADO"    # o FINAL da CHAVE_PRIMARIA (ex.: "137_COLETAS_CO05_MERCADO").
                                   # NAO e o nome do arquivo Para Processar. So serve pra conferencia
                                   # (a formula de verdade e lida do arquivo do X:, ver _convencoes).
# -----------------------------------------------------------------------------------------

LETRAS = "ABCDEFGHIJKLMNOPQRSTUVWXY"      # as 25 colunas de empilhar.COLUNAS, na ordem
COLUNAS_TEXTO = set("BCDEFOPQRSTUV")      # A, G-N e X sao numero; W e Y sao formula (O = INSTALLMENTS e texto aqui)
COLUNAS = [
    "DATA_REPASSE", "DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "DESCRIÇÃO 1", "DESCRIÇÃO 2", "VALOR REPASSE", "ESTORNO",
    "GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT",
    "INSTALLMENTS", "PAYMENT_METHOD", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID",
    "Nº QUE ESTÁ NO TOTVS", "OPERAÇÃO", "VALOR", "SEQUENCIA", "CHAVE_PRIMÁRIA",
]


# ---------------------------------------------------------------- celulas / linhas
def _serial(d):
    return (d - datetime.date(1899, 12, 30)).days


def _texto(v):
    """Texto de uma celula (IDs sem '.0'); vazio = None."""
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    t = str(v)
    return t if t != "" and t.lower() != "nan" else None


def _numero(x):
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    if f != f:
        return None
    return str(int(f)) if f.is_integer() else repr(f)


def _linha_xml(r, v, est, attrs, conv):
    """XML de uma linha do empilhado + o valor da formula W. `v` = tupla das 25 colunas.
    `conv` = convencoes lidas do proprio arquivo (ver _convencoes): colunas somadas, formula, sequencia e sufixo."""
    cel = []
    for i, col in enumerate(LETRAS[:22]):                       # A..V (dados)
        s = ' s="%s"' % est[col] if est[col] else ""
        if col == "A":
            val = str(_serial(v[i].date()))
        elif col in COLUNAS_TEXTO:
            t = _texto(v[i])
            if t is not None:
                cel.append('<c r="%s%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (col, r, s, escape(t)))
            continue
        else:
            val = _numero(v[i])
            if val is None:
                continue
        cel.append('<c r="%s%d"%s><v>%s</v></c>' % (col, r, s, val))
    w = sum(float(v[LETRAS.index(c)]) for c in conv["cols"])    # ex.: I+J+L+K (varia por arquivo)
    seq = r + conv["offset"]                                    # SEQUENCIA = linha + deslocamento lido do arquivo
    sw, sx_, sy = (' s="%s"' % est[c] if est[c] else "" for c in "WXY")
    cel.append('<c r="W%d"%s><f>%s</f><v>%s</v></c>' % (r, sw, conv["formula"](r), repr(w)))
    if conv["seq_formula"]:
        cel.append('<c r="X%d"%s><f>X%d+1</f><v>%d</v></c>' % (r, sx_, r - 1, seq))
    else:
        cel.append('<c r="X%d"%s><v>%d</v></c>' % (r, sx_, seq))
    cel.append('<c r="Y%d"%s t="str"><f>CONCATENATE(X%d,"_%s")</f><v>%d_%s</v></c>' % (r, sy, r, conv["sufixo"], seq, conv["sufixo"]))
    return '<row r="%d"%s>%s</row>' % (r, attrs, "".join(cel)), w


# ---------------------------------------------------------------- leitura do xlsx
def _atributos(el):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', el))


def _aba_dados(z):
    """Caminho interno da aba de dados (ABA, ou a primeira)."""
    rels = {a["Id"]: a["Target"] for a in map(_atributos, re.findall(r"<Relationship [^>]*>", z.read("xl/_rels/workbook.xml.rels").decode("utf-8")))}
    abas = {a["name"]: a["r:id"] for a in map(_atributos, re.findall(r"<sheet [^>]*>", z.read("xl/workbook.xml").decode("utf-8")))}
    if ABA not in abas:
        if len(abas) != 1:
            sys.exit('nao achei a aba "%s" no empilhado. Abas: %s -- ajuste ABA no topo do script' % (ABA, list(abas)))
        print('  (aba "%s" nao existe; usando a unica do arquivo: "%s")' % (ABA, next(iter(abas))))
    t = rels[abas.get(ABA) or next(iter(abas.values()))]
    return t.lstrip("/") if t.startswith("/") else "xl/" + t


def _dias_gravados(sx):
    """DATA_REPASSE (serial) que ja estao na coluna A."""
    return {int(v) for v in re.findall(r'<c r="A\d+"(?: s="\d+")?><v>(\d+)</v>', sx)}


def _convencoes(sx, linha, ult):
    """Le no arquivo real (nao no codigo): (1) a formula de VALOR (coluna W): SUM(I2,J2,L2,K2), SUM(I2:K2) ou SUM(I2+J2+K2);
    (2) se SEQUENCIA (X) e numero fixo ou formula X+1 e quanto vale em relacao ao numero da linha (X2 = 2 -> +0; X2 = 1 -> -1),
    olhando a ultima linha de dados; (3) o sufixo da CHAVE_PRIMARIA (CONCATENATE(X,"_?")).
    Evita hardcode por canal -- o script segue exatamente o que ja esta gravado no empilhado.
    Devolve dict(cols, formula(r), seq_formula, offset, sufixo)."""
    modelos = {}
    for m in re.finditer(r'<c r="W(\d+)"[^>]*><f[^>]*>(SUM\([^<]*\))</f>', sx):
        modelo = re.sub(r"([A-Z]+)%s\b" % m.group(1), r"\1{r}", m.group(2))
        modelos[modelo] = modelos.get(modelo, 0) + 1
    if not modelos:
        sys.exit('nao achei a formula "SUM(...)" da coluna W no arquivo -- confirme que ha pelo menos 1 linha de dados')
    modelo = max(modelos, key=modelos.get)
    if len(modelos) > 1:
        print("  ATENCAO: o arquivo tem mais de uma formula na coluna W: %s -- usando a mais frequente (%s)" % (modelos, modelo))
    tok = r"[A-Z]+\{r\}(?::[A-Z]+\{r\})?"
    if not re.fullmatch(r"SUM\(%s(?:[,+]%s)*\)" % (tok, tok), modelo):
        sys.exit("formula de VALOR (coluna W) nao suportada: %s" % modelo)
    cols = []
    for t in re.split(r"[,+]", modelo[4:-1]):
        ab = [x.replace("{r}", "") for x in t.split(":")]
        cols += ab if len(ab) == 1 else LETRAS[LETRAS.index(ab[0]):LETRAS.index(ab[1]) + 1]
    formula = lambda r: modelo.replace("{r}", str(r))

    m = re.search(r'<c r="X%d"[^>]*>(<f[^>]*/>|<f[^>]*>[^<]*</f>)?<v>(-?\d+)</v>' % ult, linha)
    if not m:
        sys.exit("nao achei a SEQUENCIA (coluna X) numerica na ultima linha de dados (linha %d)" % ult)
    seq_formula, offset = m.group(1) is not None, int(m.group(2)) - ult

    m = re.search(r'CONCATENATE\(X\d+,"_([^"]+)"\)', sx)
    if not m:
        sys.exit('nao achei a formula CONCATENATE da coluna Y (CHAVE_PRIMARIA) no arquivo')
    sufixo = m.group(1)
    print('  (convencoes do arquivo: VALOR = %s  |  SEQUENCIA = %s  |  CHAVE = <SEQUENCIA>_%s)' % (
        modelo.replace("{r}", ""), "formula X+1" if seq_formula else "numero fixo" + (" (linha %+d)" % offset if offset else " = linha"), sufixo))
    return dict(cols=cols, formula=formula, seq_formula=seq_formula, offset=offset, sufixo=sufixo)


def _ultima_linha(sx, fim):
    """(numero, inicio, fim) da ultima linha com dado na coluna A; se so tiver cabecalho, devolve a linha 1."""
    ini = max(0, fim - 300000)
    for m in reversed(list(re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sx[ini:fim], flags=re.S))):
        n = m.group(1)
        if re.search(r'<c r="A%s"(?: s="\d+")?><v>' % n, m.group(2)) or n == "1":
            return int(n), ini + m.start(), ini + m.end()
    raise ValueError("nao achei a ultima linha de dados")


def inserir(sx, dias, sufixo_esperado=None):
    """Devolve (sx novo, info). `dias` = lista de DataFrames no formato de empilhar.montar_linhas.
    A formula de VALOR e o sufixo da CHAVE sao lidos do proprio arquivo (_convencoes), nao do codigo.
    Se `sufixo_esperado` for passado e nao bater com o do arquivo, para (protege contra apontar
    sem querer para o empilhado de outro canal)."""
    fim = sx.rindex("</sheetData>")
    ult, ini_ult, fim_ult = _ultima_linha(sx, fim)
    assert not re.search(r"<v>|<f", sx[fim_ult:fim]), "ha dados depois da ultima linha detectada -- conferir a mao"
    linha = sx[ini_ult:fim_ult]
    conv = _convencoes(sx, linha, ult)
    if sufixo_esperado and conv["sufixo"] != sufixo_esperado:
        sys.exit('o arquivo usa sufixo "_%s" na chave, mas o script esperava "_%s" (SUFIXO no topo do script) -- arquivo errado?' % (conv["sufixo"], sufixo_esperado))
    attrs = re.search(r'<row r="\d+"([^>]*)>', linha).group(1)
    est = {c: None for c in LETRAS}
    if ult > 1:                                                  # estilos copiados da ultima linha de dados
        for c in LETRAS:
            m = re.search(r'<c r="%s%d" s="(\d+)"' % (c, ult), linha)
            est[c] = m.group(1) if m else None
    novas, r, somas = [], ult, []
    for df in dias:
        soma = 0.0
        for v in df[COLUNAS].itertuples(index=False):
            r += 1
            x, w = _linha_xml(r, v, est, attrs, conv)
            novas.append(x)
            soma += w
        somas.append(soma)
    sx2 = sx[:fim_ult] + "".join(novas) + sx[fim:]
    sx2 = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: '<dimension ref="A1:%s%d"/>' % (m.group(1), r), sx2, count=1)
    sx2 = re.sub(r'<autoFilter ref="A1:([A-Z]+)\d+"', lambda m: '<autoFilter ref="A1:%s%d"' % (m.group(1), r), sx2, count=1)
    return sx2, dict(ult=ult, primeira_nova=ult + 1, nova_ult=r, somas=somas, novas="".join(novas), linhas=r - ult, sufixo=conv["sufixo"], offset=conv["offset"])


def conferir_xml(sx2, info):
    """Confere o XML novo ainda em memoria: bem formado, linhas em ordem, SEQUENCIA/CHAVE continuas."""
    ET.fromstring(sx2.encode("utf-8"))
    nums = [int(n) for n in re.findall(r'<row r="(\d+)"', sx2)]
    assert nums == sorted(set(nums)) and nums[-1] == info["nova_ult"], "linhas duplicadas/fora de ordem"
    xs = re.findall(r'<c r="X(\d+)"[^>]*>(?:<f>[^<]*</f>)?<v>(\d+)</v>', info["novas"])
    assert len(xs) == info["linhas"] and all(int(v) == int(r) + info["offset"] for r, v in xs), "SEQUENCIA das linhas novas errada"
    ys = re.findall(r'<c r="Y(\d+)"[^>]*><f>[^<]*</f><v>([^<]*)</v>', info["novas"])
    assert len(ys) == info["linhas"] and all(v == "%d_%s" % (int(r) + info["offset"], info["sufixo"]) for r, v in ys), "CHAVE das linhas novas errada"


# ---------------------------------------------------------------- gravacao
def _ajustes(z, nova_ult):
    """Partes que acompanham a Base: tabela/filtro (se existirem) e calcChain (o Excel refaz)."""
    troca = {}
    for n in z.namelist():
        if n.startswith("xl/tables/") and n.endswith(".xml"):
            troca[n] = re.sub(r'ref="A1:Y\d+"', 'ref="A1:Y%d"' % nova_ult, z.read(n).decode("utf-8"))
    troca["xl/workbook.xml"] = re.sub(r"(\$A\$1:\$Y\$)\d+", r"\g<1>%d" % nova_ult, z.read("xl/workbook.xml").decode("utf-8"))
    remover = set()
    if "xl/calcChain.xml" in z.namelist():
        remover.add("xl/calcChain.xml")
        troca["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", z.read("[Content_Types].xml").decode("utf-8"))
        troca["xl/_rels/workbook.xml.rels"] = re.sub(r"<Relationship [^>]*calcChain[^>]*/>", "", z.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
    return troca, remover


def escrever_zip(z, saida, base_p, sx2, nova_ult):
    """Grava `saida` = o xlsx de `z` com a Base trocada por `sx2`. Confere o zip gravado (so o indice, sem reler tudo)."""
    troca, remover = _ajustes(z, nova_ult)
    troca[base_p] = sx2
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as out:
        for item in z.infolist():
            if item.filename not in remover:
                out.writestr(item.filename, troca[item.filename].encode("utf-8") if item.filename in troca else z.read(item.filename))
    with zipfile.ZipFile(saida) as chk:
        assert chk.namelist() == [n for n in z.namelist() if n not in remover], "partes do zip novo diferentes do esperado"
        assert chk.getinfo(base_p).CRC == zlib.crc32(sx2.encode("utf-8")), "a Base gravada nao confere"


def _assinatura(p):
    s = os.stat(p)
    return s.st_size, int(s.st_mtime)


# ---------------------------------------------------------------- dias (AB + TOTVS/SRI)
_IDS = ["SOURCE_ID", "EXTERNAL_REFERENCE", "SHIPPING_ID", "ORDER_ID", "PACK_ID"]
_VALORES = ["NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT",
            "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT"]
_OBRIGATORIAS = ["RELEASE_DATE", "RECORD_TYPE", "DESCRIPTION", "INSTALLMENTS", "PAYMENT_METHOD", "SHIPMENT_MODE"] + _IDS + _VALORES


def carregar_ab_center(caminho):
    """AB do Coletas 05: confere as colunas pelo NOME (o layout do FULL, de 40 colunas exatas, nao vale aqui;
    o relatorio tem 60 colunas, em outra ordem, numa unica aba "Relatorio"). Ids como texto, valores como numero."""
    df = pd.read_excel(caminho, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    faltam = [c for c in _OBRIGATORIAS if c not in df.columns]
    if faltam:
        raise ValueError("AB do Coletas 05 sem as colunas %s\nEncontrado: %s" % (faltam, list(df.columns)))
    for c in _IDS:
        df[c] = df[c].str.strip().str.replace(r"\.0$", "", regex=True)
    for c in _VALORES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df.dropna(subset=["RELEASE_DATE"]).reset_index(drop=True)


def carregar_processar(data):
    """Arquivo "Para Processar" do dia (PROCESSAR_MERCADO COLETAS_CO05 DD.MM.AAAA.xlsx): SOURCE_ID -> numero do TOTVS
    (a EXTERNAL_REFERENCE do arquivo). Ele vem sem duplicados, entao nem todo SOURCE_ID do AB esta nele."""
    d, m, a = data.split(".")
    achados = []
    nome_canal = re.sub(r"[_ ]+", "*", CANAL)                   # COLETAS_CO05 casa com "COLETAS CO05", "COLETAS_CO05"...
    for sep in (".", "_", "-"):                                  # PROCESSAR_MERCADO COLETAS_CO05 11.09.2026.xlsx (ou 11_09_2026)
        achados += glob.glob(os.path.join(PASTA_PROCESSAR, "PROCESSAR*%s*%s%s%s%s%s*.xlsx" % (nome_canal, d, sep, m, sep, a)))
    achados = sorted({x for x in achados if not os.path.basename(x).startswith("~$")})
    if len(achados) != 1:
        sys.exit("esperava 1 arquivo Para Processar do dia %s em %s, achei %d" % (data, PASTA_PROCESSAR, len(achados)))
    pp = pd.read_excel(achados[0], dtype=str)
    pp.columns = [str(c).strip() for c in pp.columns]
    for c in ("SOURCE_ID", "EXTERNAL_REFERENCE"):
        pp[c] = pp[c].str.strip().str.replace(r"\.0$", "", regex=True)
    pp = pp.dropna(subset=["SOURCE_ID", "EXTERNAL_REFERENCE"])
    return pp.drop_duplicates("SOURCE_ID").set_index("SOURCE_ID")["EXTERNAL_REFERENCE"]


def numero_totvs(ab, por_source):
    """Coluna "Nº QUE ESTA NO TOTVS" de cada linha do AB. Devolve (serie, origem, ambiguas).
    1) o SOURCE_ID esta no Para Processar -> usa o numero dele ("PP");
    2) senao (o arquivo vem sem duplicados) -> entre SHIPPING_ID, ORDER_ID e EXTERNAL_REFERENCE da propria linha,
       o que existe no conjunto de numeros do Para Processar ("cand"). Nao achou nenhum -> fica vazio ("vazio"),
       igual ao que o processo faz quando o titulo nao esta no Para Processar. Achou mais de um -> ambigua ("?"),
       porque ai nao da pra saber qual dos dois vale sem checar o TOTVS, e isso para a execucao (ver carregar_dias)."""
    nums = set(por_source.values)
    num, origem, ambiguas = [], [], []
    for r in ab.itertuples():
        n = por_source.get(r.SOURCE_ID)
        if n is not None:
            num.append(n)
            origem.append("PP")
            continue
        cand = {x for x in (r.SHIPPING_ID, r.ORDER_ID, r.EXTERNAL_REFERENCE) if isinstance(x, str) and x in nums}
        if len(cand) <= 1:
            num.append(next(iter(cand), None))
            origem.append("cand" if cand else "vazio")
        else:
            num.append(None)
            origem.append("?")
            ambiguas.append((r.Index, sorted(cand)))
    return pd.Series(num, index=ab.index), pd.Series(origem, index=ab.index), ambiguas


def montar_linhas(ab, data, num_totvs):
    """As 25 colunas do empilhado (SEQUENCIA/CHAVE ficam vazias: sao formulas, escritas na gravacao)."""
    dia, mes, ano = (int(p) for p in data.split("."))
    n = lambda c: pd.to_numeric(ab[c], errors="coerce").fillna(0)
    df = pd.DataFrame({
        "DATA_REPASSE": pd.Timestamp(ano, mes, dia), "DATE": ab["RELEASE_DATE"], "SOURCE_ID": ab["SOURCE_ID"],
        "EXTERNAL_REFERENCE": ab["EXTERNAL_REFERENCE"], "DESCRIÇÃO 1": ab["RECORD_TYPE"], "DESCRIÇÃO 2": ab["DESCRIPTION"],
        "VALOR REPASSE": n("NET_CREDIT_AMOUNT"), "ESTORNO": n("NET_DEBIT_AMOUNT"), "GROSS_AMOUNT": n("GROSS_AMOUNT"),
        "MP_FEE_AMOUNT": n("MP_FEE_AMOUNT"), "FINANCING_FEE_AMOUNT": n("FINANCING_FEE_AMOUNT"),
        "SHIPPING_FEE_AMOUNT": n("SHIPPING_FEE_AMOUNT"), "TAXES_AMOUNT": n("TAXES_AMOUNT"), "COUPON_AMOUNT": n("COUPON_AMOUNT"),
        "INSTALLMENTS": ab["INSTALLMENTS"], "PAYMENT_METHOD": ab["PAYMENT_METHOD"], "SHIPPING_ID": ab["SHIPPING_ID"],
        "SHIPMENT_MODE": ab["SHIPMENT_MODE"], "ORDER_ID": ab["ORDER_ID"], "PACK_ID": ab["PACK_ID"],
        "Nº QUE ESTÁ NO TOTVS": num_totvs, "OPERAÇÃO": ab["RECORD_TYPE"].astype(str) + "_" + ab["DESCRIPTION"].astype(str),
        "VALOR": n("NET_CREDIT_AMOUNT") - n("NET_DEBIT_AMOUNT"), "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None,
    })
    return df[COLUNAS].reset_index(drop=True)


def carregar_dias(datas):
    """Um DataFrame por dia (AB + Para Processar), ja conferido; aborta se faltar numero ou se a soma diferir do calendario."""
    for sub in ("FULL", "ES", "MOVEIS", "RB"):
        sys.path.insert(0, os.path.join(_AQUI, sub))
    sys.path.insert(0, _AQUI)
    from koncili import CANAIS_KONCILI, caminho_destino

    if CANAL not in CANAIS_KONCILI:
        sys.exit("registre o canal %s em koncili.CANAIS_KONCILI (loja + caminho do AB) antes de rodar" % CANAL)
    try:
        from calendario_alvo import buscar_alvo_loja
    except ImportError:
        buscar_alvo_loja = lambda loja, dia: None
    dias, erros = [], []
    for d in datas:
        ab = carregar_ab_center(caminho_destino(CANAL, d))
        por_source = carregar_processar(d)
        num, origem, ambiguas = numero_totvs(ab, por_source)
        if ambiguas:
            cols = ["SOURCE_ID", "RECORD_TYPE", "DESCRIPTION", "EXTERNAL_REFERENCE", "SHIPPING_ID", "ORDER_ID", "PACK_ID",
                    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT"]
            print(ab.loc[[i for i, _ in ambiguas[:15]], cols].to_string())
            print("   candidatos achados por linha:", [c for _, c in ambiguas[:15]])
            erros.append(d)
            print("  " + CANAL + " %s: %d linha(s) com mais de um candidato (SHIPPING_ID e ORDER_ID ambos no Para Processar, com numeros diferentes)" % (d, len(ambiguas)))
            continue
        df = montar_linhas(ab, d, num)
        alvo = buscar_alvo_loja(CANAIS_KONCILI[CANAL]["loja"], d.replace(".", "/"))
        soma = round(df["VALOR"].sum(), 2)
        ok = alvo is None or abs(soma - alvo) < 0.01
        vazias = (origem == "vazio").sum()
        print("  " + CANAL + " %s: AB %d linhas | Para Processar %d | numero: %d pelo PP + %d pelos candidatos + %d vazio | soma VALOR %s | calendario %s | %s" % (
            d, len(ab), len(por_source), (origem == "PP").sum(), (origem == "cand").sum(), vazias, format(soma, ",.2f"),
            "-" if alvo is None else format(alvo, ",.2f"), "sem linha no calendario" if alvo is None else ("OK" if ok else "DIFERE")), flush=True)
        if vazias:
            print(ab.loc[num[num.isna()].index, ["SOURCE_ID", "RECORD_TYPE", "DESCRIPTION", "EXTERNAL_REFERENCE", "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT"]].to_string())
        if not ok:
            erros.append(d)
        dias.append(df)
    if erros:
        sys.exit("problema em %s -- nada foi gravado" % ", ".join(erros))
    return dias


# ---------------------------------------------------------------- principal
def main(argv):
    publicar = "--publicar" in argv
    arquivo = next((a.split("=", 1)[1] for a in argv if a.startswith("--arquivo=")), os.path.join(PASTA_EMPILHADO, ARQUIVO))
    datas = sorted({a for a in argv if not a.startswith("--")}, key=lambda d: datetime.datetime.strptime(d, "%d.%m.%Y"))
    if not datas:
        sys.exit("uso: python empilhar_coletas05.py DD.MM.AAAA [...] [--publicar] [--arquivo=CAMINHO]")
    pasta, nome = os.path.split(arquivo)
    lock = os.path.join(pasta, "~$" + nome)
    if os.path.exists(lock):
        sys.exit("ARQUIVO ABERTO por alguem (~$ na pasta) -- feche e rode de novo")
    assin = _assinatura(arquivo)

    print("lendo o empilhado do X:...", flush=True)
    z = zipfile.ZipFile(arquivo)
    base_p = _aba_dados(z)
    sx = z.read(base_p).decode("utf-8")
    ja = _dias_gravados(sx)
    novos = []
    for d in datas:
        if _serial(datetime.datetime.strptime(d, "%d.%m.%Y").date()) in ja:
            print("  %s ja esta no empilhado -- pulado" % d)
        else:
            novos.append(d)
    if not novos:
        print("Nada novo pra gravar. Nada foi trocado.")
        return

    dias = carregar_dias(novos)
    sx2, info = inserir(sx, dias, sufixo_esperado=SUFIXO)
    conferir_xml(sx2, info)
    for d, df, w in zip(novos, dias, info["somas"]):
        aviso = "" if abs(w - df["VALOR"].sum()) < 0.01 else "  <-- ATENCAO: a formula da coluna W difere do VALOR (credito - debito)"
        print("  %s: soma da coluna W %s%s" % (d, format(w, ",.2f"), aviso))
    print("linhas novas: %d (%d a %d)" % (info["linhas"], info["primeira_nova"], info["nova_ult"]))
    if not publicar:
        print("OK. Nada foi gravado no X:. Para gravar: --publicar")
        return

    tmp = arquivo + ".novo"
    try:
        print("gravando...", flush=True)
        escrever_zip(z, tmp, base_p, sx2, info["nova_ult"])
        if os.path.exists(lock) or _assinatura(arquivo) != assin:
            sys.exit("o arquivo do X: mudou/abriu durante a montagem -- nao publiquei")
        z.close()                                                # no Windows nao troca com o original aberto
        os.replace(tmp, arquivo)
        print("PUBLICADO:", arquivo)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def _teste():
    """Offline: acrescenta 1 dia sintetico numa Base minima e confere linhas, formulas e dimensao.
    Caso 1 = como o exemplo do Coletas 05 (SEQUENCIA fixa = linha, SUM(I,J,L,K)); caso 2 = como o WEB (X+1, SUM(I:K))."""
    def base(w2, x2, y2, extra=""):
        return ('<worksheet><dimension ref="A1:Y%d"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>' % (3 if extra else 2)
                + '<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="I2" s="2"><v>10</v></c>'
                + w2 + x2 + y2 + '</row>' + extra + '</sheetData></worksheet>')
    v = dict.fromkeys(COLUNAS, 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "INSTALLMENTS": "12", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5,
              "PACK_ID": None, "VALOR": 84.5, "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[COLUNAS]

    # caso 1: SEQUENCIA numero fixo (X2 = 2), formula com virgulas
    sx = base('<c r="W2" s="2"><f>SUM(I2,J2,L2,K2)</f><v>10</v></c>', '<c r="X2" s="3"><v>2</v></c>',
              '<c r="Y2" s="3" t="str"><f t="shared" ref="Y2:Y2" si="0">CONCATENATE(X2,"_COLETAS_CO05_MERCADO")</f><v>2_COLETAS_CO05_MERCADO</v></c>')
    sx2, info = inserir(sx, [df], sufixo_esperado="COLETAS_CO05_MERCADO")
    conferir_xml(sx2, info)
    assert info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Y4"/>' in sx2 and "<f>SUM(I3,J3,L3,K3)</f>" in sx2 and "<f>SUM(I4,J4,L4,K4)</f>" in sx2
    assert '<c r="X3" s="3"><v>3</v></c>' in sx2 and "<f>X" not in sx2 and "<v>4_COLETAS_CO05_MERCADO</v>" in sx2
    assert 't="inlineStr"><is><t xml:space="preserve">12</t>' in sx2 and 'r="T3"' not in sx2 and 'r="C3" s="1"' not in sx2
    assert _dias_gravados(sx2) == {46275, 46276}

    # caso 2: SEQUENCIA com formula X+1 (X3 = 2 -> linha - 1), formula com intervalo
    extra = ('<row r="3" spans="1:25"><c r="A3" s="1"><v>46275</v></c><c r="W3" s="2"><f>SUM(I3:K3)</f><v>10</v></c>'
             '<c r="X3" s="3"><f>X2+1</f><v>2</v></c><c r="Y3" s="3" t="str"><f>CONCATENATE(X3,"_WEB_MERCADO")</f><v>2_WEB_MERCADO</v></c></row>')
    sx = base('<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c>', '<c r="X2" s="3"><v>1</v></c>',
              '<c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_WEB_MERCADO")</f><v>1_WEB_MERCADO</v></c>', extra)
    sx2, info = inserir(sx, [df])
    conferir_xml(sx2, info)
    assert info["primeira_nova"] == 4 and info["nova_ult"] == 5 and abs(info["somas"][0] - 169.0) < 1e-9
    assert "<f>X3+1</f><v>3</v>" in sx2 and "<f>SUM(I4:K4)</f>" in sx2 and "<v>4_WEB_MERCADO</v>" in sx2
    print("empilhar_coletas05.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
