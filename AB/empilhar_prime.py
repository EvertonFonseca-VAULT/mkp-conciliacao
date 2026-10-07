"""Empilhado do PRIME (Mercado Pago) -- monta as linhas do dia e grava direto no .xlsx do X:.

Mesmas premissas do empilhar_web.py:
Nao carrega o TOTVS. O "PEDIDO TOTVS" (coluna AJ) vem do arquivo Para Processar do dia
(C:\\RODAR TOTVS\\PROCESSAR\\PRIME\\PROCESSAR_MERCADO PRIME DD.MM.AAAA.xlsx), cruzado com o AB por SOURCE_ID.
Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam o SHIPPING_ID / ORDER_ID /
EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para Processar (se der 0 ou 2, para e avisa).
  * sem pasta de backup e sem copia local: le o .xlsx do X: e grava um ".novo" na mesma pasta;
    so no fim troca o original (os.replace, sem trafego). Se der erro antes, o original fica intacto.
  * VALOR LIQUIDO (coluna AL) e a formula SUM(I+J+K+L) (le do proprio arquivo, pode ser SUM(I:L) tambem);
    SEQUENCIA (AM) = AM<linha anterior>+1; CHAVE_PRIMARIA (AN) = CONCATENATE(AM,"_PRIME_MERCADO").
  * nao mexe em calendario/tabela a nao ser que existam no arquivo.
  * sem cache em pickle: monta, confere e grava na mesma execucao.

O que muda em relacao ao WEB:
  * o empilhado do PRIME tem 40 colunas (A..AN) em outra ordem: entra SELLER_AMOUNT, TAX_DETAIL, TAX_AMOUNT_TELCO,
    POS_*, STORE_*, CURRENCY, TAXES_DISAGGREGATED, METADATA, REFUND_ID, EFFECTIVE_COUPON_AMOUNT; nao tem
    "DESCRICAO 1/2" nem "ESTORNO" separados (sao RECORD_TYPE / DESCRIPTION / NET_DEBIT_AMOUNT).
  * o AB do PRIME vem com UMA ABA POR DIA DE REPASSE (nome da aba = dia, ex.: "11", "14"),
    e o arquivo pode ter mais de um dia. A aba lida e a do dia do repasse informado.
  * REFUND_ID nao existe no AB atual: a coluna AH fica vazia (se um dia o AB trouxer, e usada).

Uso (coloque este arquivo na mesma pasta do empilhar.py):
    python empilhar_prime.py 11.09.2026                        # monta e confere; NAO grava nada
    python empilhar_prime.py 11.09.2026 14.09.2026 --publicar  # idem e grava no empilhado do X:
    python empilhar_prime.py 14.09.2026 --ab="X:\\...\\MERCADO_PRIME_11_09_2026.xlsx"   # AB de outro dia (aba = dia)
    python empilhar_prime.py --teste                           # teste offline
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
ARQUIVO = "PRIME_MERCADO.xlsx"
CANAL = "PRIME"                   # nome do canal no koncili.CANAIS_KONCILI e no arquivo Para Processar
# (ou passe o caminho completo na linha de comando: --arquivo="X:\...\arquivo.xlsx")
ABA = "Fatura"                # nome da aba de dados do empilhado (se o arquivo so tiver 1 aba, usa ela)
PASTA_PROCESSAR = r"C:\RODAR TOTVS\PROCESSAR\PRIME"   # onde ficam os PROCESSAR_MERCADO PRIME DD.MM.AAAA.xlsx
SUFIXO = "PRIME_MERCADO"          # o FINAL da CHAVE_PRIMARIA (ex.: "137_PRIME_MERCADO").
                                   # NAO e o nome do arquivo Para Processar. So serve pra conferencia
                                   # (a formula de verdade e lida do arquivo do X:, ver _convencoes).
# -----------------------------------------------------------------------------------------


def _col(i):
    """0 -> A, 25 -> Z, 26 -> AA ..."""
    s, n = "", i + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


COLUNAS = [
    "DATA REPASSE", "DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "RECORD_TYPE", "DESCRIPTION",
    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "SELLER_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT",
    "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT", "INSTALLMENTS", "PAYMENT_METHOD", "TAX_DETAIL",
    "TAX_AMOUNT_TELCO", "TRANSACTION_APPROVAL_DATE", "POS_ID", "POS_NAME", "EXTERNAL_POS_ID", "STORE_ID", "STORE_NAME",
    "EXTERNAL_STORE_ID", "CURRENCY", "TAXES_DISAGGREGATED", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID",
    "METADATA", "REFUND_ID", "EFFECTIVE_COUPON_AMOUNT", "PEDIDO TOTVS", "OPERACAO",
    "VALOR LÍQUIDO", "SEQUENCIA", "CHAVE_PRIMÁRIA",
]
LETRAS = [_col(i) for i in range(len(COLUNAS))]          # A..AN, na ordem de COLUNAS
N_DADOS = 37                                             # A..AK sao dados; AL, AM, AN sao formulas
L_VALOR, L_SEQ, L_CHAVE = LETRAS[37], LETRAS[38], LETRAS[39]      # AL, AM, AN
L_ULT = LETRAS[-1]                                       # AN
# colunas gravadas como texto (IDs, datas em texto, descricoes); as demais dos dados (G..O, S, AI) sao numero; A e data
COLUNAS_TEXTO = {"B", "C", "D", "E", "F", "P", "Q", "R", "T", "U", "V", "W", "X", "Y", "Z",
                 "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AJ", "AK"}
assert len(COLUNAS) == 40 and L_VALOR == "AL" and L_SEQ == "AM" and L_CHAVE == "AN"


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
    """XML de uma linha do empilhado + o valor da formula AL. `v` = tupla das 40 colunas.
    `conv` = convencoes lidas do proprio arquivo (ver _convencoes): colunas somadas, formula e sufixo da chave."""
    cel = []
    for i, col in enumerate(LETRAS[:N_DADOS]):                   # A..AK (dados)
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
    w = sum(float(v[LETRAS.index(c)]) for c in conv["cols"])     # ex.: I+J+K+L (varia por arquivo)
    p = r - 1                                                   # SEQUENCIA = linha - 1
    sw, sx_, sy = (' s="%s"' % est[c] if est[c] else "" for c in (L_VALOR, L_SEQ, L_CHAVE))
    cel.append('<c r="%s%d"%s><f>%s</f><v>%s</v></c>' % (L_VALOR, r, sw, conv["formula"](r), repr(w)))
    if r == 2:                                                  # primeira linha de dados: AM1 e o cabecalho
        cel.append('<c r="%s%d"%s><v>1</v></c>' % (L_SEQ, r, sx_))
    else:
        cel.append('<c r="%s%d"%s><f>%s%d+1</f><v>%d</v></c>' % (L_SEQ, r, sx_, L_SEQ, r - 1, p))
    cel.append('<c r="%s%d"%s t="str"><f>CONCATENATE(%s%d,"_%s")</f><v>%d_%s</v></c>' % (
        L_CHAVE, r, sy, L_SEQ, r, conv["sufixo"], p, conv["sufixo"]))
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
    """DATA REPASSE (serial) que ja estao na coluna A."""
    return {int(v) for v in re.findall(r'<c r="A\d+"(?: s="\d+")?><v>(\d+)</v>', sx)}


def _convencoes(sx, ult):
    """Le no arquivo real (nao no codigo) a formula de VALOR LIQUIDO (SUM(I2+J2+K2+L2) ou SUM(I2:L2)) e qual o
    sufixo da CHAVE_PRIMARIA (CONCATENATE(AM,"_?")). Evita hardcode por canal e erro de digitar
    o sufixo errado -- o script segue exatamente o que ja esta gravado no empilhado.
    Devolve dict(cols=[colunas somadas], formula=f(linha), sufixo=...)."""
    m = re.search(r"SUM\(I\d+((?:\+[A-Z]+\d+)+)\)", sx)
    if m:                                                       # SUM(I2+J2+K2+L2)
        cols = ["I"] + re.findall(r"\+([A-Z]+)\d+", m.group(1))
        formula = lambda r: "SUM(" + "+".join("%s%d" % (c, r) for c in cols) + ")"
        desc = "SUM(" + "+".join(cols) + ")"
    else:
        m = re.search(r"SUM\(I\d+:([A-Z]+)\d+\)", sx)           # SUM(I2:L2)
        if not m:
            sys.exit('nao achei a formula "SUM(I..+J..)" ou "SUM(I..:..)" da coluna %s no arquivo -- confirme que ha pelo menos 1 linha de dados' % L_VALOR)
        fim = m.group(1)
        cols = LETRAS[LETRAS.index("I"):LETRAS.index(fim) + 1]
        formula = lambda r: "SUM(I%d:%s%d)" % (r, fim, r)
        desc = "SUM(I:%s)" % fim
    m = re.search(r'CONCATENATE\(%s\d+,"_([^"]+)"\)' % L_SEQ, sx)
    if not m:
        sys.exit('nao achei a formula CONCATENATE da coluna %s (CHAVE_PRIMARIA) no arquivo' % L_CHAVE)
    sufixo = m.group(1)
    print('  (convencoes do arquivo: VALOR LIQUIDO = %s  |  CHAVE = <SEQUENCIA>_%s)' % (desc, sufixo))
    return dict(cols=cols, formula=formula, sufixo=sufixo)


def _ultima_linha(sx, fim):
    """(numero, inicio, fim) da ultima linha com dado na coluna A; se so tiver cabecalho, devolve a linha 1."""
    ini = max(0, fim - 300000)
    for m in reversed(list(re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sx[ini:fim], flags=re.S))):
        n = m.group(1)
        if re.search(r'<c r="A%s"(?: s="\d+")?><v>' % n, m.group(2)) or n == "1":
            return int(n), ini + m.start(), ini + m.end()
    raise ValueError("nao achei a ultima linha de dados")


def inserir(sx, dias, sufixo_esperado=None):
    """Devolve (sx novo, info). `dias` = lista de DataFrames no formato de montar_linhas.
    A formula de VALOR LIQUIDO e o sufixo da CHAVE sao lidos do proprio arquivo (_convencoes), nao do codigo.
    Se `sufixo_esperado` for passado e nao bater com o do arquivo, para (protege contra apontar
    sem querer para o empilhado de outro canal)."""
    fim = sx.rindex("</sheetData>")
    ult, ini_ult, fim_ult = _ultima_linha(sx, fim)
    assert not re.search(r"<v>|<f", sx[fim_ult:fim]), "ha dados depois da ultima linha detectada -- conferir a mao"
    conv = _convencoes(sx, ult)
    if sufixo_esperado and conv["sufixo"] != sufixo_esperado:
        sys.exit('o arquivo usa sufixo "_%s" na chave, mas o script esperava "_%s" -- arquivo errado?' % (conv["sufixo"], sufixo_esperado))
    linha = sx[ini_ult:fim_ult]
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
    return sx2, dict(ult=ult, primeira_nova=ult + 1, nova_ult=r, somas=somas, novas="".join(novas), linhas=r - ult, sufixo=conv["sufixo"])


def conferir_xml(sx2, info):
    """Confere o XML novo ainda em memoria: bem formado, linhas em ordem, SEQUENCIA/CHAVE continuas."""
    ET.fromstring(sx2.encode("utf-8"))
    nums = [int(n) for n in re.findall(r'<row r="(\d+)"', sx2)]
    assert nums == sorted(set(nums)) and nums[-1] == info["nova_ult"], "linhas duplicadas/fora de ordem"
    xs = re.findall(r'<c r="%s(\d+)"[^>]*>(?:<f>[^<]*</f>)?<v>(\d+)</v>' % L_SEQ, info["novas"])
    assert len(xs) == info["linhas"] and all(int(v) == int(r) - 1 for r, v in xs), "SEQUENCIA das linhas novas errada"
    ys = re.findall(r'<c r="%s(\d+)"[^>]*><f>[^<]*</f><v>([^<]*)</v>' % L_CHAVE, info["novas"])
    assert len(ys) == info["linhas"] and all(v == "%d_%s" % (int(r) - 1, info["sufixo"]) for r, v in ys), "CHAVE das linhas novas errada"


# ---------------------------------------------------------------- gravacao
def _ajustes(z, nova_ult):
    """Partes que acompanham a Base: tabela/filtro (se existirem) e calcChain (o Excel refaz)."""
    troca = {}
    for n in z.namelist():
        if n.startswith("xl/tables/") and n.endswith(".xml"):
            troca[n] = re.sub(r'ref="A1:%s\d+"' % L_ULT, 'ref="A1:%s%d"' % (L_ULT, nova_ult), z.read(n).decode("utf-8"))
    troca["xl/workbook.xml"] = re.sub(r"(\$A\$1:\$%s\$)\d+" % L_ULT, r"\g<1>%d" % nova_ult, z.read("xl/workbook.xml").decode("utf-8"))
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
_VALORES = ["NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "SELLER_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT",
            "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT", "TAX_AMOUNT_TELCO", "EFFECTIVE_COUPON_AMOUNT"]
_TEXTOS = ["RELEASE_DATE", "APPROVAL_DATE", "RECORD_TYPE", "DESCRIPTION", "INSTALLMENTS", "PAYMENT_METHOD", "TAX_DETAIL",
           "POS_ID", "POS_NAME", "EXTERNAL_POS_ID", "STORE_ID", "STORE_NAME", "EXTERNAL_STORE_ID", "CURRENCY",
           "TAXES_DISAGGREGATED", "SHIPMENT_MODE", "METADATA"]
_OBRIGATORIAS = _TEXTOS + _IDS + _VALORES                      # REFUND_ID e opcional (o AB atual nao traz)


def carregar_ab_prime(caminho, data):
    """AB do PRIME: uma aba por dia de repasse (nome da aba = dia, ex.: "11", "14"); le a aba do dia `data`.
    Confere as colunas pelo NOME (o AB tem 60 colunas, em ordem propria). Ids como texto, valores como numero."""
    if not os.path.exists(caminho):
        sys.exit("AB do %s nao encontrado: %s\n(se o AB desse dia esta dentro do arquivo de outra data, passe --ab=CAMINHO)" % (CANAL, caminho))
    dia, mes, ano = (int(p) for p in data.split("."))
    xl = pd.ExcelFile(caminho)
    aba = next((n for n in xl.sheet_names if re.fullmatch(r"0*%d" % dia, str(n).strip())), None)
    if aba is None:
        if len(xl.sheet_names) != 1:
            sys.exit('nao achei a aba "%d" no AB %s. Abas: %s' % (dia, caminho, xl.sheet_names))
        aba = xl.sheet_names[0]
        print('  (aba "%d" nao existe; usando a unica do arquivo: "%s")' % (dia, aba))
    df = xl.parse(aba, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    faltam = [c for c in _OBRIGATORIAS if c not in df.columns]
    if faltam:
        raise ValueError("AB do %s sem as colunas %s\nEncontrado: %s" % (CANAL, faltam, list(df.columns)))
    for c in _IDS:
        df[c] = df[c].str.strip().str.replace(r"\.0$", "", regex=True)
    for c in _VALORES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["RELEASE_DATE"]).reset_index(drop=True)
    # conferencia leve: a liberacao nao pode ser depois do repasse nem muito antes (fim de semana/feriado = ate ~5 dias)
    rel = pd.to_datetime(df["RELEASE_DATE"].str[:10], errors="coerce")
    rep = pd.Timestamp(ano, mes, dia)
    fora = int(((rel > rep) | (rel < rep - pd.Timedelta(days=7))).sum())
    if fora:
        print("  ATENCAO: %d linha(s) da aba %s com RELEASE_DATE fora de [repasse-7d, repasse] -- confira se o arquivo/aba e do mes certo" % (fora, aba))
    return df


def carregar_processar(data):
    """Arquivo "Para Processar" do dia (PROCESSAR_MERCADO PRIME DD.MM.AAAA.xlsx): SOURCE_ID -> numero do TOTVS
    (a EXTERNAL_REFERENCE do arquivo). Ele vem sem duplicados, entao nem todo SOURCE_ID do AB esta nele."""
    d, m, a = data.split(".")
    achados = []
    for sep in (".", "_", "-"):                                  # PROCESSAR_MERCADO PRIME 11.09.2026.xlsx (ou 11_09_2026)
        achados += glob.glob(os.path.join(PASTA_PROCESSAR, "PROCESSAR*%s*%s%s%s%s%s*.xlsx" % (CANAL, d, sep, m, sep, a)))
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
    """Coluna "PEDIDO TOTVS" de cada linha do AB. Devolve (serie, origem, ambiguas).
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
    """As 40 colunas do empilhado (SEQUENCIA/CHAVE ficam vazias: sao formulas, escritas na gravacao;
    VALOR LIQUIDO aqui e so a conferencia = credito - debito, a formula SUM e escrita na gravacao)."""
    dia, mes, ano = (int(p) for p in data.split("."))
    n = lambda c: pd.to_numeric(ab[c], errors="coerce").fillna(0)
    refund = ab["REFUND_ID"] if "REFUND_ID" in ab.columns else None
    df = pd.DataFrame({
        "DATA REPASSE": pd.Timestamp(ano, mes, dia), "DATE": ab["RELEASE_DATE"], "SOURCE_ID": ab["SOURCE_ID"],
        "EXTERNAL_REFERENCE": ab["EXTERNAL_REFERENCE"], "RECORD_TYPE": ab["RECORD_TYPE"], "DESCRIPTION": ab["DESCRIPTION"],
        "NET_CREDIT_AMOUNT": n("NET_CREDIT_AMOUNT"), "NET_DEBIT_AMOUNT": n("NET_DEBIT_AMOUNT"), "GROSS_AMOUNT": n("GROSS_AMOUNT"),
        "SELLER_AMOUNT": n("SELLER_AMOUNT"), "MP_FEE_AMOUNT": n("MP_FEE_AMOUNT"), "FINANCING_FEE_AMOUNT": n("FINANCING_FEE_AMOUNT"),
        "SHIPPING_FEE_AMOUNT": n("SHIPPING_FEE_AMOUNT"), "TAXES_AMOUNT": n("TAXES_AMOUNT"), "COUPON_AMOUNT": n("COUPON_AMOUNT"),
        "INSTALLMENTS": ab["INSTALLMENTS"], "PAYMENT_METHOD": ab["PAYMENT_METHOD"], "TAX_DETAIL": ab["TAX_DETAIL"],
        "TAX_AMOUNT_TELCO": n("TAX_AMOUNT_TELCO"), "TRANSACTION_APPROVAL_DATE": ab["APPROVAL_DATE"],
        "POS_ID": ab["POS_ID"], "POS_NAME": ab["POS_NAME"], "EXTERNAL_POS_ID": ab["EXTERNAL_POS_ID"], "STORE_ID": ab["STORE_ID"],
        "STORE_NAME": ab["STORE_NAME"], "EXTERNAL_STORE_ID": ab["EXTERNAL_STORE_ID"], "CURRENCY": ab["CURRENCY"],
        "TAXES_DISAGGREGATED": ab["TAXES_DISAGGREGATED"], "SHIPPING_ID": ab["SHIPPING_ID"], "SHIPMENT_MODE": ab["SHIPMENT_MODE"],
        "ORDER_ID": ab["ORDER_ID"], "PACK_ID": ab["PACK_ID"], "METADATA": ab["METADATA"], "REFUND_ID": refund,
        "EFFECTIVE_COUPON_AMOUNT": n("EFFECTIVE_COUPON_AMOUNT"),
        "PEDIDO TOTVS": num_totvs, "OPERACAO": ab["RECORD_TYPE"].astype(str) + "_" + ab["DESCRIPTION"].astype(str),
        "VALOR LÍQUIDO": n("NET_CREDIT_AMOUNT") - n("NET_DEBIT_AMOUNT"), "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None,
    })
    return df[COLUNAS].reset_index(drop=True)


def carregar_dias(datas, ab_forcado=None):
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
        ab = carregar_ab_prime(ab_forcado or caminho_destino(CANAL, d), d)
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
        soma = round(df["VALOR LÍQUIDO"].sum(), 2)
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
    ab_forcado = next((a.split("=", 1)[1] for a in argv if a.startswith("--ab=")), None)
    datas = sorted({a for a in argv if not a.startswith("--")}, key=lambda d: datetime.datetime.strptime(d, "%d.%m.%Y"))
    if not datas:
        sys.exit("uso: python empilhar_prime.py DD.MM.AAAA [...] [--publicar] [--ab=CAMINHO] [--arquivo=CAMINHO]")
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

    dias = carregar_dias(novos, ab_forcado)
    sx2, info = inserir(sx, dias, sufixo_esperado=SUFIXO)
    conferir_xml(sx2, info)
    for d, df, w in zip(novos, dias, info["somas"]):
        aviso = "" if abs(w - df["VALOR LÍQUIDO"].sum()) < 0.01 else "  <-- ATENCAO: a soma das colunas da formula difere do VALOR (credito - debito)"
        print("  %s: soma da coluna %s %s%s" % (d, L_VALOR, format(w, ",.2f"), aviso))
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
    """Offline: acrescenta 1 dia sintetico numa Base minima (layout de 40 colunas) e confere linhas, formulas e dimensao."""
    linha2 = ('<row r="2" spans="1:40"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="I2" s="2"><v>10</v></c>'
              '<c r="AL2" s="2"><f>SUM(I2+J2+K2+L2)</f><v>10</v></c><c r="AM2" s="3"><v>1</v></c>'
              '<c r="AN2" s="3" t="str"><f>CONCATENATE(AM2,"_PRIME_MERCADO")</f><v>1_PRIME_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:AN2"/><sheetData><row r="1" spans="1:40"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(COLUNAS, 0.0)
    v.update({"DATA REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "RECORD_TYPE": "release",
              "DESCRIPTION": "payment", "GROSS_AMOUNT": 100.0, "SELLER_AMOUNT": 0.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5,
              "PACK_ID": None, "REFUND_ID": None, "VALOR LÍQUIDO": 84.5, "PEDIDO TOTVS": "47", "OPERACAO": "release_payment",
              "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[COLUNAS]
    sx2, info = inserir(sx, [df])
    conferir_xml(sx2, info)
    assert info["sufixo"] == "PRIME_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:AN4"/>' in sx2 and "<f>AM3+1</f><v>3</v>" in sx2 and "<v>3_PRIME_MERCADO</v>" in sx2
    assert "<f>SUM(I3+J3+K3+L3)</f>" in sx2 and '<c r="C3" s="1"' not in sx2 and 'r="AH3"' not in sx2
    assert 't="inlineStr"><is><t xml:space="preserve">123</t>' in sx2 and 'r="AF3"' not in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_prime.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
