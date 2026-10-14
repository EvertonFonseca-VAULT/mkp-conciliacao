"""Empilhado dos canais Mercado Pago (CENTER, COLETAS_CO05, MOVEIS, PRIME, WEB, FULL, ES, ELETRO) -- monta as
linhas do dia e grava direto no .xlsx do X:. Modulo profundo: a implementacao comum fica aqui, cada canal
e so o que varia de fato (COLUNAS/LETRAS, convencao de formula/sequencia, leitura do AB, leitura do Para
Processar).

Mesmas premissas em todos os canais (ver Canal em CONTEXT.md):
Nao carrega o TOTVS. O numero do TOTVS de cada linha vem do arquivo Para Processar do dia, cruzado com o
AB por SOURCE_ID. Como esse arquivo vem sem duplicados, as linhas do AB cujo SOURCE_ID nao esta nele usam
o SHIPPING_ID / ORDER_ID / EXTERNAL_REFERENCE da propria linha que exista no conjunto de numeros do Para
Processar (se der 0 ou 2, para e avisa). Sem pasta de backup e sem copia local: le o .xlsx do X: e grava
um ".novo" na mesma pasta; so no fim troca o original (os.replace, sem trafego). VALOR, SEQUENCIA e o
sufixo da CHAVE_PRIMARIA sao lidos do proprio arquivo (ver as funcoes _convencoes_*), nao do codigo -- o
script segue exatamente o que ja esta gravado no empilhado. Sem cache em pickle: monta, confere e grava
na mesma execucao. Rodar de novo com um dia que ja esta no empilhado nao duplica nada (o dia e pulado).

Uso:
    python empilhar_canais.py CANAL DD.MM.AAAA [...] [--publicar] [--arquivo=CAMINHO]
    python empilhar_canais.py CANAL --teste    # roda o teste offline de 1 canal
    python empilhar_canais.py --teste          # roda os testes offline dos 8 canais
"""

import datetime
import glob
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zlib
import zipfile
from xml.sax.saxutils import escape

import pandas as pd

_AQUI = os.path.dirname(os.path.abspath(__file__))

PASTA_EMPILHADO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\ARQUIVOS EMPILHADOS DE BAIXAS\ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS"


# ======================================================================================
# celulas / leitura do xlsx -- identico nos 5 canais
# ======================================================================================
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


def _atributos(el):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', el))


def _aba_dados(z, aba):
    """Caminho interno da aba de dados (`aba`, ou a primeira)."""
    rels = {a["Id"]: a["Target"] for a in map(_atributos, re.findall(r"<Relationship [^>]*>", z.read("xl/_rels/workbook.xml.rels").decode("utf-8")))}
    abas = {a["name"]: a["r:id"] for a in map(_atributos, re.findall(r"<sheet [^>]*>", z.read("xl/workbook.xml").decode("utf-8")))}
    if aba not in abas:
        if len(abas) != 1:
            sys.exit('nao achei a aba "%s" no empilhado. Abas: %s -- ajuste ABA no cadastro do canal' % (aba, list(abas)))
        print('  (aba "%s" nao existe; usando a unica do arquivo: "%s")' % (aba, next(iter(abas))))
    t = rels[abas.get(aba) or next(iter(abas.values()))]
    return t.lstrip("/") if t.startswith("/") else "xl/" + t


def _dias_gravados(sx):
    """DATA_REPASSE (serial) que ja estao na coluna A."""
    return {int(v) for v in re.findall(r'<c r="A\d+"(?: s="\d+")?><v>(\d+)</v>', sx)}


def _ultima_linha(sx, fim):
    """(numero, inicio, fim) da ultima linha com dado na coluna A; se so tiver cabecalho, devolve a linha 1."""
    ini = max(0, fim - 300000)
    for m in reversed(list(re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sx[ini:fim], flags=re.S))):
        n = m.group(1)
        if re.search(r'<c r="A%s"(?: s="\d+")?><v>' % n, m.group(2)) or n == "1":
            return int(n), ini + m.start(), ini + m.end()
    raise ValueError("nao achei a ultima linha de dados")


def _assinatura(p):
    s = os.stat(p)
    return s.st_size, int(s.st_mtime)


def numero_totvs(ab, por_source):
    """Coluna de numero do TOTVS de cada linha do AB. Devolve (serie, origem, ambiguas).
    1) o SOURCE_ID esta no Para Processar -> usa o numero dele ("PP");
    2) senao (o arquivo vem sem duplicados) -> entre SHIPPING_ID, ORDER_ID e EXTERNAL_REFERENCE da propria
       linha, o que existe no conjunto de numeros do Para Processar ("cand"). Nao achou nenhum -> vazio
       ("vazio"). Achou mais de um -> ambigua ("?"), e isso para a execucao (ver carregar_dias)."""
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


# ======================================================================================
# gravacao -- identico nos 5 canais (parametrizado pela ultima coluna do layout)
# ======================================================================================
def _ajustes(z, nova_ult, ultima_letra):
    """Partes que acompanham a Base: tabela/filtro (se existirem) e calcChain (o Excel refaz)."""
    troca = {}
    for n in z.namelist():
        if n.startswith("xl/tables/") and n.endswith(".xml"):
            troca[n] = re.sub(r'ref="A1:%s\d+"' % ultima_letra, 'ref="A1:%s%d"' % (ultima_letra, nova_ult), z.read(n).decode("utf-8"))
    troca["xl/workbook.xml"] = re.sub(r"(\$A\$1:\$%s\$)\d+" % ultima_letra, r"\g<1>%d" % nova_ult, z.read("xl/workbook.xml").decode("utf-8"))
    remover = set()
    if "xl/calcChain.xml" in z.namelist():
        remover.add("xl/calcChain.xml")
        troca["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", z.read("[Content_Types].xml").decode("utf-8"))
        troca["xl/_rels/workbook.xml.rels"] = re.sub(r"<Relationship [^>]*calcChain[^>]*/>", "", z.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
    return troca, remover


def escrever_zip(z, saida, base_p, sx2, nova_ult, ultima_letra):
    """Grava `saida` = o xlsx de `z` com a Base trocada por `sx2`. Confere o zip gravado (so o indice, sem reler tudo)."""
    troca, remover = _ajustes(z, nova_ult, ultima_letra)
    troca[base_p] = sx2
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as out:
        for item in z.infolist():
            if item.filename not in remover:
                out.writestr(item.filename, troca[item.filename].encode("utf-8") if item.filename in troca else z.read(item.filename))
    with zipfile.ZipFile(saida) as chk:
        assert chk.namelist() == [n for n in z.namelist() if n not in remover], "partes do zip novo diferentes do esperado"
        assert chk.getinfo(base_p).CRC == zlib.crc32(sx2.encode("utf-8")), "a Base gravada nao confere"


# ======================================================================================
# grupo "padrao" -- CENTER, COLETAS_CO05, MOVEIS (25 colunas A..Y, formula/sequencia lidas
# do proprio arquivo com offset generico: SEQUENCIA pode ser numero fixo ou formula X+1)
# ======================================================================================
def _linha_xml_padrao(r, v, est, attrs, conv, letras, colunas_texto):
    """XML de uma linha do empilhado + o valor da formula W. `v` = tupla das colunas do layout de 25."""
    cel = []
    for i, col in enumerate(letras[:22]):                       # A..V (dados)
        s = ' s="%s"' % est[col] if est[col] else ""
        if col == "A":
            val = str(_serial(v[i].date()))
        elif col in colunas_texto:
            t = _texto(v[i])
            if t is not None:
                cel.append('<c r="%s%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (col, r, s, escape(t)))
            continue
        else:
            val = _numero(v[i])
            if val is None:
                continue
        cel.append('<c r="%s%d"%s><v>%s</v></c>' % (col, r, s, val))
    w = sum(float(v[letras.index(c)]) for c in conv["cols"])    # ex.: I+J+K (varia por canal)
    seq = r + conv["offset"]                                    # SEQUENCIA = linha + deslocamento lido do arquivo
    sw, sx_, sy = (' s="%s"' % est[c] if est[c] else "" for c in "WXY")
    cel.append('<c r="W%d"%s><f>%s</f><v>%s</v></c>' % (r, sw, conv["formula"](r), repr(w)))
    if conv["seq_formula"]:
        cel.append('<c r="X%d"%s><f>X%d+1</f><v>%d</v></c>' % (r, sx_, r - 1, seq))
    else:
        cel.append('<c r="X%d"%s><v>%d</v></c>' % (r, sx_, seq))
    cel.append('<c r="Y%d"%s t="str"><f>CONCATENATE(X%d,"_%s")</f><v>%d_%s</v></c>' % (r, sy, r, conv["sufixo"], seq, conv["sufixo"]))
    return '<row r="%d"%s>%s</row>' % (r, attrs, "".join(cel)), w


def _convencoes_padrao(sx, linha, ult, letras):
    """Le no arquivo real (nao no codigo): (1) a formula de VALOR (coluna W): SUM(I2,J2,K2), SUM(I2:K2),
    SUM(I2+J2+K2) ou (sem SUM) I2+J2++K2; (2) se SEQUENCIA (X) e numero fixo ou formula X+1 e quanto vale
    em relacao ao numero da linha; (3) o sufixo da CHAVE_PRIMARIA (CONCATENATE(X,"_?")).
    Devolve dict(cols, formula(r), seq_formula, offset, sufixo)."""
    modelos = {}
    for m in re.finditer(r'<c r="W(\d+)"[^>]*><f[^>]*>([^<]+)</f>', sx):
        modelo = re.sub(r"([A-Z]+)%s\b" % m.group(1), r"\1{r}", m.group(2))
        modelos[modelo] = modelos.get(modelo, 0) + 1
    if not modelos:
        sys.exit('nao achei a formula da coluna W (VALOR) no arquivo -- confirme que ha pelo menos 1 linha de dados')
    modelo = max(modelos, key=modelos.get)
    if len(modelos) > 1:
        print("  ATENCAO: o arquivo tem mais de uma formula na coluna W: %s -- usando a mais frequente (%s)" % (modelos, modelo))
    tok = r"[A-Z]+\{r\}(?::[A-Z]+\{r\})?"
    if not re.fullmatch(r"SUM\(%s(?:[,+]+%s)*\)|%s(?:\++%s)*" % (tok, tok, tok, tok), modelo):
        sys.exit("formula de VALOR (coluna W) nao suportada: %s" % modelo)
    cols = []
    for t in re.split(r"[,+]+", modelo[4:-1] if modelo.startswith("SUM(") else modelo):
        ab = [x.replace("{r}", "") for x in t.split(":")]
        cols += ab if len(ab) == 1 else letras[letras.index(ab[0]):letras.index(ab[1]) + 1]
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


def _conferir_xml_padrao(sx2, info):
    """Confere o XML novo ainda em memoria: bem formado, linhas em ordem, SEQUENCIA/CHAVE continuas."""
    ET.fromstring(sx2.encode("utf-8"))
    nums = [int(n) for n in re.findall(r'<row r="(\d+)"', sx2)]
    assert nums == sorted(set(nums)) and nums[-1] == info["nova_ult"], "linhas duplicadas/fora de ordem"
    xs = re.findall(r'<c r="X(\d+)"[^>]*>(?:<f>[^<]*</f>)?<v>(\d+)</v>', info["novas"])
    assert len(xs) == info["linhas"] and all(int(v) == int(r) + info["offset"] for r, v in xs), "SEQUENCIA das linhas novas errada"
    ys = re.findall(r'<c r="Y(\d+)"[^>]*><f>[^<]*</f><v>([^<]*)</v>', info["novas"])
    assert len(ys) == info["linhas"] and all(v == "%d_%s" % (int(r) + info["offset"], info["sufixo"]) for r, v in ys), "CHAVE das linhas novas errada"


# ======================================================================================
# grupo "web" -- WEB (25 colunas A..Y, VALOR sempre SUM(I:fim), SEQUENCIA sempre X+1)
# ======================================================================================
def _linha_xml_web(r, v, est, attrs, fim_w, sufixo, letras, colunas_texto):
    cel = []
    for i, col in enumerate(letras[:22]):
        s = ' s="%s"' % est[col] if est[col] else ""
        if col == "A":
            val = str(_serial(v[i].date()))
        elif col in colunas_texto:
            t = _texto(v[i])
            if t is not None:
                cel.append('<c r="%s%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (col, r, s, escape(t)))
            continue
        else:
            val = _numero(v[i])
            if val is None:
                continue
        cel.append('<c r="%s%d"%s><v>%s</v></c>' % (col, r, s, val))
    fim_i = letras.index("I")
    w = sum(float(v[i]) for i in range(fim_i, letras.index(fim_w) + 1))   # I ate fim_w (lido do arquivo)
    p = r - 1
    sw, sx_, sy = (' s="%s"' % est[c] if est[c] else "" for c in "WXY")
    cel.append('<c r="W%d"%s><f>SUM(I%d:%s%d)</f><v>%s</v></c>' % (r, sw, r, fim_w, r, repr(w)))
    if r == 2:                                                  # primeira linha de dados: X1 e o cabecalho
        cel.append('<c r="X%d"%s><v>1</v></c>' % (r, sx_))
    else:
        cel.append('<c r="X%d"%s><f>X%d+1</f><v>%d</v></c>' % (r, sx_, r - 1, p))
    cel.append('<c r="Y%d"%s t="str"><f>CONCATENATE(X%d,"_%s")</f><v>%d_%s</v></c>' % (r, sy, r, sufixo, p, sufixo))
    return '<row r="%d"%s>%s</row>' % (r, attrs, "".join(cel)), w


def _convencoes_web(sx, ult):
    """Le no arquivo real ate onde vai a formula de VALOR (SUM(I:?)) e o sufixo da CHAVE_PRIMARIA."""
    m = re.search(r"SUM\(I\d+:([A-Z]+)\d+\)", sx)
    if not m:
        sys.exit('nao achei a formula "SUM(I..:..)" da coluna W no arquivo -- confirme que ha pelo menos 1 linha de dados')
    fim_w = m.group(1)
    m = re.search(r'CONCATENATE\(X\d+,"_([^"]+)"\)', sx)
    if not m:
        sys.exit('nao achei a formula CONCATENATE da coluna Y (CHAVE_PRIMARIA) no arquivo')
    sufixo = m.group(1)
    print('  (convencoes do arquivo: VALOR = SUM(I:%s)  |  CHAVE = <SEQUENCIA>_%s)' % (fim_w, sufixo))
    return fim_w, sufixo


def _conferir_xml_web(sx2, info):
    ET.fromstring(sx2.encode("utf-8"))
    nums = [int(n) for n in re.findall(r'<row r="(\d+)"', sx2)]
    assert nums == sorted(set(nums)) and nums[-1] == info["nova_ult"], "linhas duplicadas/fora de ordem"
    xs = re.findall(r'<c r="X(\d+)"[^>]*>(?:<f>[^<]*</f>)?<v>(\d+)</v>', info["novas"])
    assert len(xs) == info["linhas"] and all(int(v) == int(r) - 1 for r, v in xs), "SEQUENCIA das linhas novas errada"
    ys = re.findall(r'<c r="Y(\d+)"[^>]*><f>[^<]*</f><v>([^<]*)</v>', info["novas"])
    assert len(ys) == info["linhas"] and all(v == "%d_%s" % (int(r) - 1, info["sufixo"]) for r, v in ys), "CHAVE das linhas novas errada"


# ======================================================================================
# grupo "prime" -- PRIME (40 colunas A..AN, VALOR LIQUIDO em AL, SEQUENCIA em AM sempre AM+1)
# ======================================================================================
def _col(i):
    """0 -> A, 25 -> Z, 26 -> AA ..."""
    s, n = "", i + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _linha_xml_prime(r, v, est, attrs, conv, letras, colunas_texto, n_dados, l_valor, l_seq, l_chave):
    cel = []
    for i, col in enumerate(letras[:n_dados]):                  # A..AK (dados)
        s = ' s="%s"' % est[col] if est[col] else ""
        if col == "A":
            val = str(_serial(v[i].date()))
        elif col in colunas_texto:
            t = _texto(v[i])
            if t is not None:
                cel.append('<c r="%s%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (col, r, s, escape(t)))
            continue
        else:
            val = _numero(v[i])
            if val is None:
                continue
        cel.append('<c r="%s%d"%s><v>%s</v></c>' % (col, r, s, val))
    w = sum(float(v[letras.index(c)]) for c in conv["cols"])
    p = r - 1
    sw, sx_, sy = (' s="%s"' % est[c] if est[c] else "" for c in (l_valor, l_seq, l_chave))
    cel.append('<c r="%s%d"%s><f>%s</f><v>%s</v></c>' % (l_valor, r, sw, conv["formula"](r), repr(w)))
    if r == 2:                                                  # primeira linha de dados: AM1 e o cabecalho
        cel.append('<c r="%s%d"%s><v>1</v></c>' % (l_seq, r, sx_))
    else:
        cel.append('<c r="%s%d"%s><f>%s%d+1</f><v>%d</v></c>' % (l_seq, r, sx_, l_seq, r - 1, p))
    cel.append('<c r="%s%d"%s t="str"><f>CONCATENATE(%s%d,"_%s")</f><v>%d_%s</v></c>' % (
        l_chave, r, sy, l_seq, r, conv["sufixo"], p, conv["sufixo"]))
    return '<row r="%d"%s>%s</row>' % (r, attrs, "".join(cel)), w


def _convencoes_prime(sx, l_valor, l_seq, l_chave, letras):
    """Le no arquivo real a formula de VALOR LIQUIDO (SUM(I2+J2+K2+L2) ou SUM(I2:L2)) e o sufixo da
    CHAVE_PRIMARIA (CONCATENATE(AM,"_?")). Devolve dict(cols, formula(r), sufixo)."""
    m = re.search(r"SUM\(I\d+((?:\+[A-Z]+\d+)+)\)", sx)
    if m:                                                       # SUM(I2+J2+K2+L2)
        cols = ["I"] + re.findall(r"\+([A-Z]+)\d+", m.group(1))
        formula = lambda r: "SUM(" + "+".join("%s%d" % (c, r) for c in cols) + ")"
        desc = "SUM(" + "+".join(cols) + ")"
    else:
        m = re.search(r"SUM\(I\d+:([A-Z]+)\d+\)", sx)           # SUM(I2:L2)
        if not m:
            sys.exit('nao achei a formula "SUM(I..+J..)" ou "SUM(I..:..)" da coluna %s no arquivo -- confirme que ha pelo menos 1 linha de dados' % l_valor)
        fim = m.group(1)
        cols = letras[letras.index("I"):letras.index(fim) + 1]
        formula = lambda r: "SUM(I%d:%s%d)" % (r, fim, r)
        desc = "SUM(I:%s)" % fim
    m = re.search(r'CONCATENATE\(%s\d+,"_([^"]+)"\)' % l_seq, sx)
    if not m:
        sys.exit('nao achei a formula CONCATENATE da coluna %s (CHAVE_PRIMARIA) no arquivo' % l_chave)
    sufixo = m.group(1)
    print('  (convencoes do arquivo: VALOR LIQUIDO = %s  |  CHAVE = <SEQUENCIA>_%s)' % (desc, sufixo))
    return dict(cols=cols, formula=formula, sufixo=sufixo)


def _conferir_xml_prime(sx2, info, l_seq, l_chave):
    ET.fromstring(sx2.encode("utf-8"))
    nums = [int(n) for n in re.findall(r'<row r="(\d+)"', sx2)]
    assert nums == sorted(set(nums)) and nums[-1] == info["nova_ult"], "linhas duplicadas/fora de ordem"
    xs = re.findall(r'<c r="%s(\d+)"[^>]*>(?:<f>[^<]*</f>)?<v>(\d+)</v>' % l_seq, info["novas"])
    assert len(xs) == info["linhas"] and all(int(v) == int(r) - 1 for r, v in xs), "SEQUENCIA das linhas novas errada"
    ys = re.findall(r'<c r="%s(\d+)"[^>]*><f>[^<]*</f><v>([^<]*)</v>' % l_chave, info["novas"])
    assert len(ys) == info["linhas"] and all(v == "%d_%s" % (int(r) - 1, info["sufixo"]) for r, v in ys), "CHAVE das linhas novas errada"


_LETRAS_PRIME = [_col(i) for i in range(40)]
_N_DADOS_PRIME = 37
_L_VALOR_PRIME, _L_SEQ_PRIME, _L_CHAVE_PRIME = _LETRAS_PRIME[37], _LETRAS_PRIME[38], _LETRAS_PRIME[39]
assert _L_VALOR_PRIME == "AL" and _L_SEQ_PRIME == "AM" and _L_CHAVE_PRIME == "AN"


# ======================================================================================
# inserir / conferir_xml -- despacham pro grupo do canal (padrao/web/prime)
# ======================================================================================
def inserir(sx, dias, cfg, sufixo_esperado=None):
    """Devolve (sx novo, info). `dias` = lista de DataFrames no formato de montar_linhas do canal.
    A formula de VALOR e o sufixo da CHAVE sao lidos do proprio arquivo (ver _convencoes_*), nao do
    codigo. Se `sufixo_esperado` for passado e nao bater com o do arquivo, para (protege contra apontar
    sem querer para o empilhado de outro canal)."""
    letras, grupo = cfg["letras"], cfg["grupo"]
    fim = sx.rindex("</sheetData>")
    ult, ini_ult, fim_ult = _ultima_linha(sx, fim)
    assert not re.search(r"<v>|<f", sx[fim_ult:fim]), "ha dados depois da ultima linha detectada -- conferir a mao"
    linha = sx[ini_ult:fim_ult]

    fim_w = None
    if grupo == "prime":
        conv = _convencoes_prime(sx, cfg["l_valor"], cfg["l_seq"], cfg["l_chave"], letras)
        sufixo_atual = conv["sufixo"]
    elif grupo == "web":
        fim_w, sufixo_atual = _convencoes_web(sx, ult)
    else:
        conv = _convencoes_padrao(sx, linha, ult, letras)
        sufixo_atual = conv["sufixo"]
    if sufixo_esperado and sufixo_atual != sufixo_esperado:
        sys.exit('o arquivo usa sufixo "_%s" na chave, mas o script esperava "_%s" (SUFIXO no cadastro do canal) -- arquivo errado?' % (sufixo_atual, sufixo_esperado))

    attrs = re.search(r'<row r="\d+"([^>]*)>', linha).group(1)
    est = {c: None for c in letras}
    if ult > 1:                                                  # estilos copiados da ultima linha de dados
        for c in letras:
            m = re.search(r'<c r="%s%d" s="(\d+)"' % (c, ult), linha)
            est[c] = m.group(1) if m else None

    novas, r, somas = [], ult, []
    for df in dias:
        soma = 0.0
        for v in df[cfg["colunas"]].itertuples(index=False):
            r += 1
            if grupo == "prime":
                x, w = _linha_xml_prime(r, v, est, attrs, conv, letras, cfg["colunas_texto"], cfg["n_dados"], cfg["l_valor"], cfg["l_seq"], cfg["l_chave"])
            elif grupo == "web":
                x, w = _linha_xml_web(r, v, est, attrs, fim_w, sufixo_atual, letras, cfg["colunas_texto"])
            else:
                x, w = _linha_xml_padrao(r, v, est, attrs, conv, letras, cfg["colunas_texto"])
            novas.append(x)
            soma += w
        somas.append(soma)
    sx2 = sx[:fim_ult] + "".join(novas) + sx[fim:]
    sx2 = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: '<dimension ref="A1:%s%d"/>' % (m.group(1), r), sx2, count=1)
    sx2 = re.sub(r'<autoFilter ref="A1:([A-Z]+)\d+"', lambda m: '<autoFilter ref="A1:%s%d"' % (m.group(1), r), sx2, count=1)
    info = dict(ult=ult, primeira_nova=ult + 1, nova_ult=r, somas=somas, novas="".join(novas), linhas=r - ult, sufixo=sufixo_atual)
    if grupo == "padrao":
        info["offset"] = conv["offset"]
    return sx2, info


def conferir_xml(sx2, info, cfg):
    """Confere o XML novo ainda em memoria: bem formado, linhas em ordem, SEQUENCIA/CHAVE continuas."""
    grupo = cfg["grupo"]
    if grupo == "prime":
        _conferir_xml_prime(sx2, info, cfg["l_seq"], cfg["l_chave"])
    elif grupo == "web":
        _conferir_xml_web(sx2, info)
    else:
        _conferir_xml_padrao(sx2, info)


# ======================================================================================
# Para Processar -- 3 algoritmos de localizacao do arquivo do dia
# ======================================================================================
def _carregar_processar_glob(data, pasta_processar, canal):
    """PROCESSAR_MERCADO <CANAL> DD.MM.AAAA.xlsx (ou 11_09_2026): SOURCE_ID -> numero do TOTVS."""
    d, m, a = data.split(".")
    achados = []
    for sep in (".", "_", "-"):
        achados += glob.glob(os.path.join(pasta_processar, "PROCESSAR*%s*%s%s%s%s%s*.xlsx" % (canal, d, sep, m, sep, a)))
    achados = sorted({x for x in achados if not os.path.basename(x).startswith("~$")})
    if len(achados) != 1:
        sys.exit("esperava 1 arquivo Para Processar do dia %s em %s, achei %d" % (data, pasta_processar, len(achados)))
    pp = pd.read_excel(achados[0], dtype=str)
    pp.columns = [str(c).strip() for c in pp.columns]
    for c in ("SOURCE_ID", "EXTERNAL_REFERENCE"):
        pp[c] = pp[c].str.strip().str.replace(r"\.0$", "", regex=True)
    pp = pp.dropna(subset=["SOURCE_ID", "EXTERNAL_REFERENCE"])
    return pp.drop_duplicates("SOURCE_ID").set_index("SOURCE_ID")["EXTERNAL_REFERENCE"]


def _carregar_processar_glob_wildcard(data, pasta_processar, canal):
    """Igual a _carregar_processar_glob, mas casa "COLETAS_CO05" com "COLETAS CO05", "COLETAS_CO05" etc.
    (o nome do canal pode vir com espaco OU underscore no arquivo real, que nao e o CANAL cadastrado)."""
    d, m, a = data.split(".")
    nome_canal = re.sub(r"[_ ]+", "*", canal)
    achados = []
    for sep in (".", "_", "-"):
        achados += glob.glob(os.path.join(pasta_processar, "PROCESSAR*%s*%s%s%s%s%s*.xlsx" % (nome_canal, d, sep, m, sep, a)))
    achados = sorted({x for x in achados if not os.path.basename(x).startswith("~$")})
    if len(achados) != 1:
        sys.exit("esperava 1 arquivo Para Processar do dia %s em %s, achei %d" % (data, pasta_processar, len(achados)))
    pp = pd.read_excel(achados[0], dtype=str)
    pp.columns = [str(c).strip() for c in pp.columns]
    for c in ("SOURCE_ID", "EXTERNAL_REFERENCE"):
        pp[c] = pp[c].str.strip().str.replace(r"\.0$", "", regex=True)
    pp = pp.dropna(subset=["SOURCE_ID", "EXTERNAL_REFERENCE"])
    return pp.drop_duplicates("SOURCE_ID").set_index("SOURCE_ID")["EXTERNAL_REFERENCE"]


def _carregar_processar_glob_sem_ano(data, pasta_processar, canal):
    """Como _carregar_processar_glob, mas o nome do arquivo NAO leva o ano (ex.: COLETAS ES -- rodar_canal.py
    grava "PROCESSAR_MERCADO COLETAS ES DD.MM.xlsx", sem ano, ver CANAIS_PP["ES"]["ano_no_nome"]=False em
    rodar_canal.py): casa so por DD.MM. Se a pasta tiver 2 anos com o mesmo DD.MM, para e avisa (mesma
    protecao de sempre) -- na pratica a pasta guarda só o ciclo do ano corrente."""
    d, m, _a = data.split(".")
    achados = []
    for sep in (".", "_", "-"):
        achados += glob.glob(os.path.join(pasta_processar, "PROCESSAR*%s*%s%s%s*.xlsx" % (canal, d, sep, m)))
    achados = sorted({x for x in achados if not os.path.basename(x).startswith("~$")})
    if len(achados) != 1:
        sys.exit("esperava 1 arquivo Para Processar do dia %s em %s, achei %d" % (data, pasta_processar, len(achados)))
    pp = pd.read_excel(achados[0], dtype=str)
    pp.columns = [str(c).strip() for c in pp.columns]
    for c in ("SOURCE_ID", "EXTERNAL_REFERENCE"):
        pp[c] = pp[c].str.strip().str.replace(r"\.0$", "", regex=True)
    pp = pp.dropna(subset=["SOURCE_ID", "EXTERNAL_REFERENCE"])
    return pp.drop_duplicates("SOURCE_ID").set_index("SOURCE_ID")["EXTERNAL_REFERENCE"]


def _norm(t):
    """Sem acento e maiusculo (MÓVEIS == MOVEIS)."""
    return unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().upper()


def _carregar_processar_normalizado(data, pasta_processar, canal):
    """Como _carregar_processar_glob, mas casa o nome sem acento e sem diferenciar espaco/underscore,
    e a data pode vir com ponto, underscore ou hifen (usado quando o nome do arquivo real varia mais)."""
    d, m, a = data.split(".")
    if not os.path.isdir(pasta_processar):
        pai = os.path.dirname(pasta_processar)
        sys.exit("a pasta do Para Processar nao existe: %s\n  pastas encontradas em %s: %s" % (
            pasta_processar, pai, sorted(os.listdir(pai)) if os.path.isdir(pai) else "(nem essa existe)"))
    tokens = [t for t in re.split(r"[_ ]+", _norm(canal)) if t]
    datas = ["%s%s%s%s%s" % (d, sep, m, sep, a) for sep in ".-_"]
    achados = []
    for nome in os.listdir(pasta_processar):
        n = _norm(nome)
        if nome.startswith("~$") or not n.endswith(".XLSX") or not n.startswith("PROCESSAR"):
            continue
        if all(t in n for t in tokens) and any(x in n for x in datas):
            achados.append(os.path.join(pasta_processar, nome))
    achados.sort()
    if len(achados) != 1:
        sys.exit("esperava 1 arquivo Para Processar do dia %s em %s, achei %d%s" % (
            data, pasta_processar, len(achados), (": " + str([os.path.basename(x) for x in achados])) if achados else ""))
    pp = pd.read_excel(achados[0], dtype=str)
    pp.columns = [str(c).strip() for c in pp.columns]
    for c in ("SOURCE_ID", "EXTERNAL_REFERENCE"):
        pp[c] = pp[c].str.strip().str.replace(r"\.0$", "", regex=True)
    pp = pp.dropna(subset=["SOURCE_ID", "EXTERNAL_REFERENCE"])
    return pp.drop_duplicates("SOURCE_ID").set_index("SOURCE_ID")["EXTERNAL_REFERENCE"]


# ======================================================================================
# AB -- leitura do relatorio do canal
# ======================================================================================
_IDS = ["SOURCE_ID", "EXTERNAL_REFERENCE", "SHIPPING_ID", "ORDER_ID", "PACK_ID"]
_VALORES_SIMPLES = ["NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT",
                     "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT"]
_OBRIGATORIAS_SIMPLES = ["RELEASE_DATE", "RECORD_TYPE", "DESCRIPTION", "INSTALLMENTS", "PAYMENT_METHOD", "SHIPMENT_MODE"] + _IDS + _VALORES_SIMPLES


def _carregar_ab_simples(caminho, nome_exibicao):
    """AB simples (Center/Coletas 05/Web): confere as colunas pelo NOME (o layout do FULL, de 40 colunas
    exatas, nao vale aqui). Ids como texto, valores como numero."""
    df = pd.read_excel(caminho, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    faltam = [c for c in _OBRIGATORIAS_SIMPLES if c not in df.columns]
    if faltam:
        raise ValueError("AB do %s sem as colunas %s\nEncontrado: %s" % (nome_exibicao, faltam, list(df.columns)))
    for c in _IDS:
        df[c] = df[c].str.strip().str.replace(r"\.0$", "", regex=True)
    for c in _VALORES_SIMPLES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df.dropna(subset=["RELEASE_DATE"]).reset_index(drop=True)


_PT_EN_MOVEIS = {
    "DATA DE LIBERAÇÃO": "RELEASE_DATE", "ID DA OPERAÇÃO NO MERCADO PAGO": "SOURCE_ID", "CÓDIGO DE REFERÊNCIA": "EXTERNAL_REFERENCE",
    "TIPO DE REGISTRO": "RECORD_TYPE", "DESCRIÇÃO": "DESCRIPTION", "VALOR LÍQUIDO CREDITADO": "NET_CREDIT_AMOUNT",
    "VALOR LÍQUIDO DEBITADO": "NET_DEBIT_AMOUNT", "VALOR BRUTO DA OPERAÇÃO": "GROSS_AMOUNT",
    "TARIFA DO MERCADO PAGO OU DO MERCADO LIVRE": "MP_FEE_AMOUNT", "TARIFA POR OFERECER PARCELAMENTO SEM JUROS": "FINANCING_FEE_AMOUNT",
    "CUSTO DE ENVIO": "SHIPPING_FEE_AMOUNT", "IMPOSTOS COBRADOS POR RETENÇÕES DE IIBB": "TAXES_AMOUNT", "CUPOM DE DESCONTO": "COUPON_AMOUNT",
    "PARCELAS": "INSTALLMENTS", "MEIO DE PAGAMENTO": "PAYMENT_METHOD", "ID DO ENVIO": "SHIPPING_ID", "FORMA DE ENVIO": "SHIPMENT_MODE",
    "ID DO PEDIDO": "ORDER_ID", "ID DO PACOTE": "PACK_ID",
}


def _data_texto(v, date_sufixo):
    """RELEASE_DATE como texto AAAA-MM-DDTHH:MM:SS<date_sufixo>. O AB do Moveis traz data/hora do Excel
    (lida como "2026-08-28 11:13:17"); se ja vier como texto ISO com "T" (AB antigo), fica como esta; se a
    coluna de data nao tiver tipo de data no Excel (ELETRO as vezes traz so o numero, ex. "46262.51..."),
    converte do serial do Excel (dias desde 1899-12-30)."""
    if not isinstance(v, str) or not v.strip():
        return None
    v = v.strip()
    if re.match(r"\d{4}-\d{2}-\d{2}T", v):
        return v
    if re.match(r"\d+(\.\d+)?$", v):
        t = pd.to_datetime(float(v), unit="D", origin="1899-12-30")
    else:
        t = pd.to_datetime(v, errors="coerce")
    return None if pd.isna(t) else t.strftime("%Y-%m-%dT%H:%M:%S") + date_sufixo


def _carregar_ab_moveis(caminho, date_sufixo):
    """AB do Moveis: uma aba ("Plan1"), colunas em portugues (traduzidas por _PT_EN_MOVEIS; se vierem em
    ingles, tambem serve). Confere as colunas pelo NOME. Ids como texto, valores como numero, RELEASE_DATE
    como texto ISO."""
    df = pd.read_excel(caminho, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={k: v for k, v in _PT_EN_MOVEIS.items() if k in df.columns and v not in df.columns})
    faltam = [c for c in _OBRIGATORIAS_SIMPLES if c not in df.columns]
    if faltam:
        raise ValueError("AB do Moveis sem as colunas %s\nEncontrado: %s" % (faltam, list(df.columns)))
    for c in _IDS:
        df[c] = df[c].str.strip().str.replace(r"\.0$", "", regex=True)
    for c in _VALORES_SIMPLES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["RELEASE_DATE"]).reset_index(drop=True)
    df["RELEASE_DATE"] = df["RELEASE_DATE"].map(lambda v: _data_texto(v, date_sufixo))
    if df["RELEASE_DATE"].isna().any():
        sys.exit("AB do Moveis: %d linha(s) com DATA DE LIBERAÇÃO que nao consegui ler como data" % df["RELEASE_DATE"].isna().sum())
    return df


def _normalizar_ab_sem_dtype_str(df, nome_exibicao):
    """Para ABs lidos sem dtype=str (FULL/ES, via processar_full/processar_es.carregar_ab -- ja confere
    o layout exato/aba/traducao deles): ids convertidos pra texto (sem ".0" de float, ver _texto), valores
    pra numero. Confere so a PRESENCA das colunas que o empilhado usa (igual _carregar_ab_simples)."""
    df = df.copy()
    faltam = [c for c in _OBRIGATORIAS_SIMPLES if c not in df.columns]
    if faltam:
        raise ValueError("AB do %s sem as colunas %s\nEncontrado: %s" % (nome_exibicao, faltam, list(df.columns)))
    for c in _IDS:
        df[c] = df[c].map(_texto)
    for c in _VALORES_SIMPLES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df.dropna(subset=["RELEASE_DATE"]).reset_index(drop=True)


def _carregar_ab_full(caminho):
    """AB do FULL: aba 'Relatorio', layout fixo de 40 colunas (ver processar_full.carregar_ab, que ja
    confere o layout exato -- so normaliza ids/valores aqui, igual os demais canais)."""
    import processar_full
    return _normalizar_ab_sem_dtype_str(processar_full.carregar_ab(caminho), "FULL")


def _carregar_ab_es(caminho):
    """AB do COLETAS ES: aba 'Relatorio'/'Sheet0', varias variantes de layout (ver processar_es.carregar_ab,
    que ja resolve aba/traducao/reordenacao) -- so normaliza ids/valores aqui."""
    import processar_es
    return _normalizar_ab_sem_dtype_str(processar_es.carregar_ab(caminho), "ES")


_VALORES_PRIME = ["NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "SELLER_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT",
                   "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT", "TAX_AMOUNT_TELCO", "EFFECTIVE_COUPON_AMOUNT"]
_TEXTOS_PRIME = ["RELEASE_DATE", "APPROVAL_DATE", "RECORD_TYPE", "DESCRIPTION", "INSTALLMENTS", "PAYMENT_METHOD", "TAX_DETAIL",
                  "POS_ID", "POS_NAME", "EXTERNAL_POS_ID", "STORE_ID", "STORE_NAME", "EXTERNAL_STORE_ID", "CURRENCY",
                  "TAXES_DISAGGREGATED", "SHIPMENT_MODE", "METADATA"]
_OBRIGATORIAS_PRIME = _TEXTOS_PRIME + _IDS + _VALORES_PRIME          # REFUND_ID e opcional (o AB atual nao traz)


def _carregar_ab_prime(caminho, data, canal):
    """AB do PRIME: uma aba por dia de repasse (nome da aba = dia, ex.: "11", "14"); le a aba do dia
    `data`. Confere as colunas pelo NOME. Ids como texto, valores como numero."""
    if not os.path.exists(caminho):
        sys.exit("AB do %s nao encontrado: %s\n(se o AB desse dia esta dentro do arquivo de outra data, passe --ab=CAMINHO)" % (canal, caminho))
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
    faltam = [c for c in _OBRIGATORIAS_PRIME if c not in df.columns]
    if faltam:
        raise ValueError("AB do %s sem as colunas %s\nEncontrado: %s" % (canal, faltam, list(df.columns)))
    for c in _IDS:
        df[c] = df[c].str.strip().str.replace(r"\.0$", "", regex=True)
    for c in _VALORES_PRIME:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["RELEASE_DATE"]).reset_index(drop=True)
    rel = pd.to_datetime(df["RELEASE_DATE"].str[:10], errors="coerce")
    rep = pd.Timestamp(ano, mes, dia)
    fora = int(((rel > rep) | (rel < rep - pd.Timedelta(days=7))).sum())
    if fora:
        print("  ATENCAO: %d linha(s) da aba %s com RELEASE_DATE fora de [repasse-7d, repasse] -- confira se o arquivo/aba e do mes certo" % (fora, aba))
    return df


# ======================================================================================
# montar_linhas -- 2 formatos (25 colunas "padrao" partilhado por center/coletas05/moveis/web; prime)
# ======================================================================================
def _montar_linhas_padrao(ab, data, num_totvs, colunas):
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
    return df[colunas].reset_index(drop=True)


def _montar_linhas_prime(ab, data, num_totvs, colunas):
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
    return df[colunas].reset_index(drop=True)


# ======================================================================================
# cadastro dos canais -- o que varia de fato (ver Canal em CONTEXT.md)
# ======================================================================================
_LETRAS_25 = "ABCDEFGHIJKLMNOPQRSTUVWXY"
_COLUNAS_25 = [
    "DATA_REPASSE", "DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "DESCRIÇÃO 1", "DESCRIÇÃO 2", "VALOR REPASSE", "ESTORNO",
    "GROSS_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT", "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT",
    "INSTALLMENTS", "PAYMENT_METHOD", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID",
    "Nº QUE ESTÁ NO TOTVS", "OPERAÇÃO", "VALOR", "SEQUENCIA", "CHAVE_PRIMÁRIA",
]
_COLUNAS_PRIME = [
    "DATA REPASSE", "DATE", "SOURCE_ID", "EXTERNAL_REFERENCE", "RECORD_TYPE", "DESCRIPTION",
    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT", "GROSS_AMOUNT", "SELLER_AMOUNT", "MP_FEE_AMOUNT", "FINANCING_FEE_AMOUNT",
    "SHIPPING_FEE_AMOUNT", "TAXES_AMOUNT", "COUPON_AMOUNT", "INSTALLMENTS", "PAYMENT_METHOD", "TAX_DETAIL",
    "TAX_AMOUNT_TELCO", "TRANSACTION_APPROVAL_DATE", "POS_ID", "POS_NAME", "EXTERNAL_POS_ID", "STORE_ID", "STORE_NAME",
    "EXTERNAL_STORE_ID", "CURRENCY", "TAXES_DISAGGREGATED", "SHIPPING_ID", "SHIPMENT_MODE", "ORDER_ID", "PACK_ID",
    "METADATA", "REFUND_ID", "EFFECTIVE_COUPON_AMOUNT", "PEDIDO TOTVS", "OPERACAO",
    "VALOR LÍQUIDO", "SEQUENCIA", "CHAVE_PRIMÁRIA",
]
_COLUNAS_TEXTO_PRIME = {"B", "C", "D", "E", "F", "P", "Q", "R", "T", "U", "V", "W", "X", "Y", "Z",
                         "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AJ", "AK"}
assert len(_COLUNAS_PRIME) == 40

# So fatos de layout de Excel/saida desse front -- vendedor/raiz TOTVS moram
# em CONCILIACAO/canais.py (FONTE UNICA), loja do calendario mora em
# COMUM/calendario_alvo.LOJA_POR_CANAL; nao duplicar aqui (ver revisao de
# arquitetura de 14/10/2026).
CANAIS = {
    "CENTER": dict(
        script_sufixo="center", arquivo="CENTER_MERCADO.xlsx", aba="Base",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\CENTER", sufixo="CENTER_MERCADO",
        letras=_LETRAS_25, ultima_letra="Y", colunas=_COLUNAS_25, colunas_texto=set("BCDEFPQRSTUV"),
        grupo="padrao", valor_col="VALOR", letra_valor="W", aviso_formula="a formula da coluna W difere do VALOR (credito - debito)",
        carregar_ab=lambda caminho: _carregar_ab_simples(caminho, "Center"),
        carregar_processar=_carregar_processar_glob, montar_linhas=_montar_linhas_padrao,
    ),
    # layout 25 colunas igual WEB; SEQUENCIA (X) fixo por linha (X2=2), nao X+1 como
    # os demais; INSTALLMENTS (O) gravado como texto.
    "COLETAS_CO05": dict(
        script_sufixo="coletas05", arquivo="COLETAS_CO05_MERCADO_2025.xlsx", aba="Base",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\COLETAS CO05", sufixo="COLETAS_CO05_MERCADO",
        letras=_LETRAS_25, ultima_letra="Y", colunas=_COLUNAS_25, colunas_texto=set("BCDEFOPQRSTUV"),
        grupo="padrao", valor_col="VALOR", letra_valor="W", aviso_formula="a formula da coluna W difere do VALOR (credito - debito)",
        carregar_ab=lambda caminho: _carregar_ab_simples(caminho, "Coletas 05"),
        carregar_processar=_carregar_processar_glob_wildcard, montar_linhas=_montar_linhas_padrao,
    ),
    # AB vem com cabecalhos em PORTUGUES (traduzidos) e DATE como texto com virgula
    # (AAAA-MM-DDTHH:MM:SS,000-03:00) -- ver _carregar_ab_moveis. INSTALLMENTS (O)
    # gravado como texto.
    "MOVEIS": dict(
        script_sufixo="moveis", arquivo="MOVEIS_MERCADO.xlsx", aba="EMPILHADO",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\MOVEIS", sufixo="MÓVEIS_MERCADO",
        letras=_LETRAS_25, ultima_letra="Y", colunas=_COLUNAS_25, colunas_texto=set("BCDEFOPQRSTUV"),
        grupo="padrao", valor_col="VALOR", letra_valor="W", aviso_formula="a formula da coluna W difere do VALOR (credito - debito)",
        carregar_ab=lambda caminho: _carregar_ab_moveis(caminho, ",000-03:00"),
        carregar_processar=_carregar_processar_normalizado, montar_linhas=_montar_linhas_padrao,
    ),
    # 40 colunas (A..AN), AB com UMA ABA POR DIA DE REPASSE (nome = dia, ex. "11"):
    # le a aba do dia informado. --ab=CAMINHO forca outro arquivo (aba = mesmo dia).
    "PRIME": dict(
        script_sufixo="prime", arquivo="PRIME_MERCADO.xlsx", aba="Fatura",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\PRIME", sufixo="PRIME_MERCADO",
        letras=_LETRAS_PRIME, ultima_letra=_L_CHAVE_PRIME, colunas=_COLUNAS_PRIME, colunas_texto=_COLUNAS_TEXTO_PRIME,
        grupo="prime", n_dados=_N_DADOS_PRIME, l_valor=_L_VALOR_PRIME, l_seq=_L_SEQ_PRIME, l_chave=_L_CHAVE_PRIME,
        valor_col="VALOR LÍQUIDO", letra_valor=_L_VALOR_PRIME, aviso_formula="a soma das colunas da formula difere do VALOR (credito - debito)",
        aceita_ab_forcado=True,
        carregar_ab=lambda caminho, data: _carregar_ab_prime(caminho, data, "PRIME"),
        carregar_processar=_carregar_processar_glob, montar_linhas=_montar_linhas_prime,
    ),
    "WEB": dict(
        script_sufixo="web", arquivo="WEB_MERCADO.xlsx", aba="WEB",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\WEB", sufixo="WEB_MERCADO",
        letras=_LETRAS_25, ultima_letra="Y", colunas=_COLUNAS_25, colunas_texto=set("BCDEFPQRSTUV"),
        grupo="web", valor_col="VALOR", letra_valor="W", aviso_formula="SUM(I:K) difere do VALOR (credito - debito)",
        carregar_ab=lambda caminho: _carregar_ab_simples(caminho, "Web"),
        carregar_processar=_carregar_processar_glob, montar_linhas=_montar_linhas_padrao,
    ),
    "FULL": dict(
        script_sufixo="full", arquivo="FULL MERCADO_2024.xlsx", aba="Base",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\FULL", sufixo="FULL_MERCADO",
        letras=_LETRAS_25, ultima_letra="Y", colunas=_COLUNAS_25, colunas_texto=set("BCDEFOPQRSTUV"),
        grupo="padrao", valor_col="VALOR", letra_valor="W", aviso_formula="a formula da coluna W difere do VALOR (credito - debito)",
        carregar_ab=_carregar_ab_full,
        carregar_processar=_carregar_processar_glob, montar_linhas=_montar_linhas_padrao,
    ),
    "ES": dict(
        script_sufixo="es", arquivo="COLETAS_ES_MERCADO_2024.xlsx", aba="Base",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\COLETAS ES", sufixo="COLETAS_MERCADO",
        letras=_LETRAS_25, ultima_letra="Z", colunas=_COLUNAS_25, colunas_texto=set("BCDEFOPQRSTUV"),
        grupo="padrao", valor_col="VALOR", letra_valor="W", aviso_formula="a formula da coluna W difere do VALOR (credito - debito)",
        carregar_ab=_carregar_ab_es,
        carregar_processar=_carregar_processar_glob_sem_ano, montar_linhas=_montar_linhas_padrao,
    ),
    # AB vem com cabecalhos em PORTUGUES (traduzidos) e DATE como texto com virgula
    # -- ver _carregar_ab_moveis (mesma funcao do MOVEIS).
    "ELETRO": dict(
        script_sufixo="eletro", arquivo="ELETRODOMESTICO_MERCADO.xlsx", aba="EMPILHADO",
        pasta_processar=r"C:\RODAR TOTVS\PROCESSAR\ELETRO", sufixo="ELETRODOMÉSTICOS_MERCADO",
        letras=_LETRAS_25, ultima_letra="Y", colunas=_COLUNAS_25, colunas_texto=set("BCDEFOPQRSTUV"),
        grupo="padrao", valor_col="VALOR", letra_valor="W", aviso_formula="a formula da coluna W difere do VALOR (credito - debito)",
        carregar_ab=lambda caminho: _carregar_ab_moveis(caminho, ",000-03:00"),
        carregar_processar=_carregar_processar_normalizado, montar_linhas=_montar_linhas_padrao,
    ),
}


# ======================================================================================
# dias (AB + TOTVS/SRI) -- generico, despacha pro cadastro do canal
# ======================================================================================
def carregar_dias(canal_nome, datas, ab_forcado=None):
    """Um DataFrame por dia (AB + Para Processar), ja conferido; aborta se faltar numero ou se a soma
    diferir do calendario."""
    cfg = CANAIS[canal_nome]
    for sub in ("PARA_PROCESSAR/FULL", "PARA_PROCESSAR/ES", "COMUM", "KONCILI"):
        sys.path.insert(0, os.path.join(_AQUI, "..", sub))
    from koncili import CANAIS_KONCILI, caminho_destino

    if canal_nome not in CANAIS_KONCILI:
        sys.exit("registre o canal %s em koncili.CANAIS_KONCILI (loja + caminho do AB) antes de rodar" % canal_nome)
    # calendario_alvo.py (COMUM) e um modulo-folha de verdade desde a revisao
    # de arquitetura (card 2): nao depende mais de CONCILIACAO, so stdlib +
    # openpyxl (ja e dependencia obrigatoria do pipeline) -- sem motivo real
    # pro import falhar, nao precisa mais de fallback silencioso.
    from calendario_alvo import buscar_alvo_loja

    valor_col = cfg["valor_col"]
    dias, erros = [], []
    for d in datas:
        if cfg["grupo"] == "prime":
            ab = cfg["carregar_ab"](ab_forcado or caminho_destino(canal_nome, d), d)
        else:
            ab = cfg["carregar_ab"](caminho_destino(canal_nome, d))
        por_source = cfg["carregar_processar"](d, cfg["pasta_processar"], canal_nome)
        num, origem, ambiguas = numero_totvs(ab, por_source)
        if ambiguas:
            cols = ["SOURCE_ID", "RECORD_TYPE", "DESCRIPTION", "EXTERNAL_REFERENCE", "SHIPPING_ID", "ORDER_ID", "PACK_ID",
                    "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT"]
            print(ab.loc[[i for i, _ in ambiguas[:15]], cols].to_string())
            print("   candidatos achados por linha:", [c for _, c in ambiguas[:15]])
            erros.append(d)
            print("  " + canal_nome + " %s: %d linha(s) com mais de um candidato (SHIPPING_ID e ORDER_ID ambos no Para Processar, com numeros diferentes)" % (d, len(ambiguas)))
            continue
        df = cfg["montar_linhas"](ab, d, num, cfg["colunas"])
        alvo = buscar_alvo_loja(CANAIS_KONCILI[canal_nome]["loja"], d.replace(".", "/"))
        soma = round(df[valor_col].sum(), 2)
        ok = alvo is None or abs(soma - alvo) < 0.01
        vazias = (origem == "vazio").sum()
        print("  " + canal_nome + " %s: AB %d linhas | Para Processar %d | numero: %d pelo PP + %d pelos candidatos + %d vazio | soma VALOR %s | calendario %s | %s" % (
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


# ======================================================================================
# principal -- generico, despacha pro cadastro do canal
# ======================================================================================
def main(canal_nome, argv):
    cfg = CANAIS[canal_nome]
    publicar = "--publicar" in argv
    arquivo = next((a.split("=", 1)[1] for a in argv if a.startswith("--arquivo=")), os.path.join(PASTA_EMPILHADO, cfg["arquivo"]))
    ab_forcado = next((a.split("=", 1)[1] for a in argv if a.startswith("--ab=")), None) if cfg.get("aceita_ab_forcado") else None
    datas = sorted({a for a in argv if not a.startswith("--")}, key=lambda d: datetime.datetime.strptime(d, "%d.%m.%Y"))
    if not datas:
        flags = "[--publicar] [--arquivo=CAMINHO]" + (" [--ab=CAMINHO]" if cfg.get("aceita_ab_forcado") else "")
        sys.exit("uso: python empilhar_%s.py DD.MM.AAAA [...] %s" % (cfg["script_sufixo"], flags))
    pasta, nome = os.path.split(arquivo)
    lock = os.path.join(pasta, "~$" + nome)
    if os.path.exists(lock):
        sys.exit("ARQUIVO ABERTO por alguem (~$ na pasta) -- feche e rode de novo")
    assin = _assinatura(arquivo)

    print("lendo o empilhado do X:...", flush=True)
    z = zipfile.ZipFile(arquivo)
    base_p = _aba_dados(z, cfg["aba"])
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

    dias = carregar_dias(canal_nome, novos, ab_forcado)
    sx2, info = inserir(sx, dias, cfg, sufixo_esperado=cfg["sufixo"])
    conferir_xml(sx2, info, cfg)
    valor_col = cfg["valor_col"]
    for d, df, w in zip(novos, dias, info["somas"]):
        aviso = "" if abs(w - df[valor_col].sum()) < 0.01 else "  <-- ATENCAO: %s" % cfg["aviso_formula"]
        print("  %s: soma da coluna %s %s%s" % (d, cfg["letra_valor"], format(w, ",.2f"), aviso))
    print("linhas novas: %d (%d a %d)" % (info["linhas"], info["primeira_nova"], info["nova_ult"]))
    if not publicar:
        print("OK. Nada foi gravado no X:. Para gravar: --publicar")
        return

    tmp = arquivo + ".novo"
    try:
        print("gravando...", flush=True)
        escrever_zip(z, tmp, base_p, sx2, info["nova_ult"], cfg["ultima_letra"])
        if os.path.exists(lock) or _assinatura(arquivo) != assin:
            sys.exit("o arquivo do X: mudou/abriu durante a montagem -- nao publiquei")
        z.close()                                                # no Windows nao troca com o original aberto
        os.replace(tmp, arquivo)
        print("PUBLICADO:", arquivo)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


# ======================================================================================
# testes offline (um por canal, --teste roda os 5)
# ======================================================================================
def _teste_center():
    cfg = CANAIS["CENTER"]
    linha2 = ('<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="G2" s="2"><v>10</v></c>'
              '<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c><c r="X2" s="3"><v>1</v></c><c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_CENTER_MERCADO")</f><v>1_CENTER_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:Y2"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5, "PACK_ID": None, "VALOR": 84.5,
              "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]
    sx2, info = inserir(sx, [df], cfg, sufixo_esperado="CENTER_MERCADO")
    conferir_xml(sx2, info, cfg)
    assert info["sufixo"] == "CENTER_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Y4"/>' in sx2 and '<c r="X3" s="3"><v>2</v></c>' in sx2 and "<v>3_CENTER_MERCADO</v>" in sx2
    assert '<c r="C3" s="1"' not in sx2 and 't="inlineStr"><is><t xml:space="preserve">123</t>' in sx2 and 'r="T3"' not in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_canais.py [CENTER]: ok")


def _teste_coletas05():
    cfg = CANAIS["COLETAS_CO05"]

    def base(w2, x2, y2, extra=""):
        return ('<worksheet><dimension ref="A1:Y%d"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>' % (3 if extra else 2)
                + '<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="I2" s="2"><v>10</v></c>'
                + w2 + x2 + y2 + '</row>' + extra + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "INSTALLMENTS": "12", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5,
              "PACK_ID": None, "VALOR": 84.5, "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]

    # caso 1: SEQUENCIA numero fixo (X2 = 2), formula com virgulas
    sx = base('<c r="W2" s="2"><f>SUM(I2,J2,L2,K2)</f><v>10</v></c>', '<c r="X2" s="3"><v>2</v></c>',
              '<c r="Y2" s="3" t="str"><f t="shared" ref="Y2:Y2" si="0">CONCATENATE(X2,"_COLETAS_CO05_MERCADO")</f><v>2_COLETAS_CO05_MERCADO</v></c>')
    sx2, info = inserir(sx, [df], cfg, sufixo_esperado="COLETAS_CO05_MERCADO")
    conferir_xml(sx2, info, cfg)
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
    sx2, info = inserir(sx, [df], cfg)
    conferir_xml(sx2, info, cfg)
    assert info["primeira_nova"] == 4 and info["nova_ult"] == 5 and abs(info["somas"][0] - 169.0) < 1e-9
    assert "<f>X3+1</f><v>3</v>" in sx2 and "<f>SUM(I4:K4)</f>" in sx2 and "<v>4_WEB_MERCADO</v>" in sx2
    print("empilhar_canais.py [COLETAS_CO05]: ok")


def _teste_moveis():
    cfg = CANAIS["MOVEIS"]

    def base(w2, x2, y2, extra=""):
        return ('<worksheet><dimension ref="A1:Y%d"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>' % (3 if extra else 2)
                + '<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="I2" s="2"><v>10</v></c>'
                + w2 + x2 + y2 + '</row>' + extra + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 4), "DATE": "2026-09-03T10:00:00,000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "INSTALLMENTS": "12", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5,
              "PACK_ID": None, "VALOR": 84.5, "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]

    # caso 1: como o Moveis (SEQUENCIA fixa = linha - 1, I+J++K, sufixo com acento)
    sx = base('<c r="W2" s="2"><f>I2+J2++K2</f><v>10</v></c>', '<c r="X2" s="3"><v>1</v></c>',
              '<c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_MÓVEIS_MERCADO")</f><v>1_MÓVEIS_MERCADO</v></c>')
    sx2, info = inserir(sx, [df], cfg, sufixo_esperado="MÓVEIS_MERCADO")
    conferir_xml(sx2, info, cfg)
    assert info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Y4"/>' in sx2 and "<f>I3+J3++K3</f>" in sx2 and "<f>I4+J4++K4</f>" in sx2
    assert '<c r="X3" s="3"><v>2</v></c>' in sx2 and "<f>X" not in sx2 and "<v>3_MÓVEIS_MERCADO</v>" in sx2
    assert 't="inlineStr"><is><t xml:space="preserve">12</t>' in sx2 and 'r="T3"' not in sx2 and 'r="C3" s="1"' not in sx2
    assert _dias_gravados(sx2) == {46275, _serial(datetime.date(2026, 9, 4))}

    # caso 2: como o WEB (X+1, SUM(I:K))
    extra = ('<row r="3" spans="1:25"><c r="A3" s="1"><v>46275</v></c><c r="W3" s="2"><f>SUM(I3:K3)</f><v>10</v></c>'
             '<c r="X3" s="3"><f>X2+1</f><v>2</v></c><c r="Y3" s="3" t="str"><f>CONCATENATE(X3,"_WEB_MERCADO")</f><v>2_WEB_MERCADO</v></c></row>')
    sx = base('<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c>', '<c r="X2" s="3"><v>1</v></c>',
              '<c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_WEB_MERCADO")</f><v>1_WEB_MERCADO</v></c>', extra)
    sx2, info = inserir(sx, [df], cfg)
    conferir_xml(sx2, info, cfg)
    assert info["primeira_nova"] == 4 and info["nova_ult"] == 5 and abs(info["somas"][0] - 169.0) < 1e-9
    assert "<f>X3+1</f><v>3</v>" in sx2 and "<f>SUM(I4:K4)</f>" in sx2 and "<v>4_WEB_MERCADO</v>" in sx2

    # caso 3: como o Coletas 05 (fixa = linha, SUM(I,J,L,K))
    sx = base('<c r="W2" s="2"><f>SUM(I2,J2,L2,K2)</f><v>10</v></c>', '<c r="X2" s="3"><v>2</v></c>',
              '<c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_COLETAS_CO05_MERCADO")</f><v>2_COLETAS_CO05_MERCADO</v></c>')
    sx2, info = inserir(sx, [df], cfg)
    conferir_xml(sx2, info, cfg)
    assert "<f>SUM(I3,J3,L3,K3)</f>" in sx2 and '<c r="X3" s="3"><v>3</v></c>' in sx2

    # data e nomes do AB
    assert _data_texto("2026-08-28 11:13:17", ",000-03:00") == "2026-08-28T11:13:17,000-03:00"
    assert _data_texto("2023-06-21T17:06:15,000-03:00", ",000-03:00") == "2023-06-21T17:06:15,000-03:00" and _data_texto(None, ",000-03:00") is None
    assert _data_texto("46262.510092592594", ",000-03:00") == "2026-08-28T12:14:32,000-03:00"
    assert _norm("MÓVEIS mercado") == "MOVEIS MERCADO"
    assert set(_PT_EN_MOVEIS.values()) >= set(_OBRIGATORIAS_SIMPLES)
    print("empilhar_canais.py [MOVEIS]: ok")


def _teste_prime():
    cfg = CANAIS["PRIME"]
    linha2 = ('<row r="2" spans="1:40"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="I2" s="2"><v>10</v></c>'
              '<c r="AL2" s="2"><f>SUM(I2+J2+K2+L2)</f><v>10</v></c><c r="AM2" s="3"><v>1</v></c>'
              '<c r="AN2" s="3" t="str"><f>CONCATENATE(AM2,"_PRIME_MERCADO")</f><v>1_PRIME_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:AN2"/><sheetData><row r="1" spans="1:40"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "RECORD_TYPE": "release",
              "DESCRIPTION": "payment", "GROSS_AMOUNT": 100.0, "SELLER_AMOUNT": 0.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5,
              "PACK_ID": None, "REFUND_ID": None, "VALOR LÍQUIDO": 84.5, "PEDIDO TOTVS": "47", "OPERACAO": "release_payment",
              "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]
    sx2, info = inserir(sx, [df], cfg)
    conferir_xml(sx2, info, cfg)
    assert info["sufixo"] == "PRIME_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:AN4"/>' in sx2 and "<f>AM3+1</f><v>3</v>" in sx2 and "<v>3_PRIME_MERCADO</v>" in sx2
    assert "<f>SUM(I3+J3+K3+L3)</f>" in sx2 and '<c r="C3" s="1"' not in sx2 and 'r="AH3"' not in sx2
    assert 't="inlineStr"><is><t xml:space="preserve">123</t>' in sx2 and 'r="AF3"' not in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_canais.py [PRIME]: ok")


def _teste_web():
    cfg = CANAIS["WEB"]
    linha2 = ('<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="G2" s="2"><v>10</v></c>'
              '<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c><c r="X2" s="3"><v>1</v></c><c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_WEB_MERCADO")</f><v>1_WEB_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:Y2"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5, "PACK_ID": None, "VALOR": 84.5,
              "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]
    sx2, info = inserir(sx, [df], cfg)
    conferir_xml(sx2, info, cfg)
    assert info["sufixo"] == "WEB_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Y4"/>' in sx2 and "<f>X3+1</f><v>3</v>" in sx2 and "<v>3_WEB_MERCADO</v>" in sx2
    assert '<c r="C3" s="1"' not in sx2 and 't="inlineStr"><is><t xml:space="preserve">123</t>' in sx2 and 'r="T3"' not in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_canais.py [WEB]: ok")


def _teste_full():
    cfg = CANAIS["FULL"]
    linha2 = ('<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="G2" s="2"><v>10</v></c>'
              '<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c><c r="X2" s="3"><v>1</v></c><c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_FULL_MERCADO")</f><v>1_FULL_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:Y2"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5, "PACK_ID": None, "VALOR": 84.5,
              "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]
    sx2, info = inserir(sx, [df], cfg, sufixo_esperado="FULL_MERCADO")
    conferir_xml(sx2, info, cfg)
    assert info["sufixo"] == "FULL_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Y4"/>' in sx2 and "<v>3_FULL_MERCADO</v>" in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_canais.py [FULL]: ok")


def _teste_es():
    cfg = CANAIS["ES"]
    linha2 = ('<row r="2" spans="1:26"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="G2" s="2"><v>10</v></c>'
              '<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c><c r="X2" s="3"><v>1</v></c><c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_COLETAS_MERCADO")</f><v>1_COLETAS_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:Z2"/><sheetData><row r="1" spans="1:26"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5, "PACK_ID": None, "VALOR": 84.5,
              "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]
    sx2, info = inserir(sx, [df], cfg, sufixo_esperado="COLETAS_MERCADO")
    conferir_xml(sx2, info, cfg)
    assert info["sufixo"] == "COLETAS_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Z4"/>' in sx2 and "<v>3_COLETAS_MERCADO</v>" in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_canais.py [ES]: ok")


def _teste_eletro():
    cfg = CANAIS["ELETRO"]
    linha2 = ('<row r="2" spans="1:25"><c r="A2" s="1"><v>46275</v></c><c r="C2" t="inlineStr"><is><t>1</t></is></c><c r="G2" s="2"><v>10</v></c>'
              '<c r="W2" s="2"><f>SUM(I2:K2)</f><v>10</v></c><c r="X2" s="3"><v>1</v></c><c r="Y2" s="3" t="str"><f>CONCATENATE(X2,"_ELETRODOMÉSTICOS_MERCADO")</f><v>1_ELETRODOMÉSTICOS_MERCADO</v></c></row>')
    sx = ('<worksheet><dimension ref="A1:Y2"/><sheetData><row r="1" spans="1:25"><c r="A1" t="s"><v>0</v></c></row>'
          + linha2 + '</sheetData></worksheet>')
    v = dict.fromkeys(cfg["colunas"], 0.0)
    v.update({"DATA_REPASSE": pd.Timestamp(2026, 9, 11), "DATE": "2026-09-10T10:00:00.000-03:00", "SOURCE_ID": 123.0, "DESCRIÇÃO 1": "release",
              "DESCRIÇÃO 2": "payment", "GROSS_AMOUNT": 100.0, "MP_FEE_AMOUNT": -10.0, "FINANCING_FEE_AMOUNT": -5.5, "PACK_ID": None, "VALOR": 84.5,
              "Nº QUE ESTÁ NO TOTVS": "47", "OPERAÇÃO": "release_payment", "SEQUENCIA": None, "CHAVE_PRIMÁRIA": None})
    df = pd.DataFrame([v, v])[cfg["colunas"]]
    sx2, info = inserir(sx, [df], cfg, sufixo_esperado="ELETRODOMÉSTICOS_MERCADO")
    conferir_xml(sx2, info, cfg)
    assert info["sufixo"] == "ELETRODOMÉSTICOS_MERCADO" and info["primeira_nova"] == 3 and info["nova_ult"] == 4 and abs(info["somas"][0] - 169.0) < 1e-9
    assert '<dimension ref="A1:Y4"/>' in sx2 and "<v>3_ELETRODOMÉSTICOS_MERCADO</v>" in sx2
    assert _dias_gravados(sx2) == {46275, 46276}
    print("empilhar_canais.py [ELETRO]: ok")


TESTES = {
    "CENTER": _teste_center, "COLETAS_CO05": _teste_coletas05, "MOVEIS": _teste_moveis,
    "PRIME": _teste_prime, "WEB": _teste_web, "FULL": _teste_full, "ES": _teste_es, "ELETRO": _teste_eletro,
}


def teste(canal_nome):
    TESTES[canal_nome]()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in CANAIS:
        _canal_nome, _resto = sys.argv[1], sys.argv[2:]
        if "--teste" in _resto:
            teste(_canal_nome)
        else:
            main(_canal_nome, _resto)
    elif "--teste" in sys.argv:
        for _canal in TESTES:
            teste(_canal)
    else:
        sys.exit(
            "uso: python empilhar_canais.py <CANAL> [dias...] [--publicar]\n"
            "     python empilhar_canais.py <CANAL> --teste\n"
            "     python empilhar_canais.py --teste  (roda os 8 canais)\n"
            f"canais: {', '.join(CANAIS)}"
        )
