"""Grava no empilhado do Mercado (FULL MERCADO_2024.xlsx / COLETAS_ES_MERCADO_2024.xlsx)
as linhas que o ensaio (empilhar.py) montou e deixou em empilhar_cache/.

Acrescenta direto no XML do .xlsx: o arquivo tem 400+ MB de XML, uma tabela do
Excel e fórmulas; carregar tudo no openpyxl e regravar perderia os valores em
cache das fórmulas e recursos do Excel. Textos novos entram como inlineStr (o
Excel converte quando salvar). O que faz, em ordem:
  1. SEQUENCIA errada (cauda toda com o mesmo valor) -> volta a ser linha-1
  2. linhas novas na Base (SEQUENCIA = linha-1, CHAVE = <seq>_<canal>, VALOR = formula da coluna)
  3. uma linha por dia na aba CALENDARIO (dias que tem VALOR_EB no calendario principal)
  4. filtros/tabela ajustados; calcChain removido (o Excel refaz)

Uso (o arquivo do X: so e trocado com --publicar; sem ele so monta e confere local):
    python empilhar_gravar.py FULL              # monta em C:\\BACKUP EMPILHADO\\_novo_FULL.xlsx e confere
    python empilhar_gravar.py FULL --publicar   # confere e troca o arquivo do X: (com backup)
"""

import datetime
import glob
import os
import pickle
import re
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape

_AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _AQUI)

import pandas as pd

from empilhar import CACHE, CHAVE_SUFIXO
from calendario_alvo import buscar_alvo_loja
from koncili import CANAIS_KONCILI

PASTA_EMPILHADO = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\ARQUIVOS EMPILHADOS DE BAIXAS\ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS"
PASTA_BACKUP = r"C:\BACKUP EMPILHADO"
ARQUIVO = {"FULL": "FULL MERCADO_2024.xlsx", "ES": "COLETAS_ES_MERCADO_2024.xlsx"}
LETRAS = "ABCDEFGHIJKLMNOPQRSTUVWXY"                       # as 25 colunas de empilhar.COLUNAS, na ordem
COLUNAS_TEXTO = set("BCDEFPQRSTUV")                        # o resto (A, G-O, X) e numero; W e Y sao formula


def _serial(d):
    return (d - datetime.date(1899, 12, 30)).days


def _texto(v):
    """Texto de uma celula (IDs sem '.0'); vazio = None."""
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def _numero(x):
    return repr(float(x)) if not float(x).is_integer() else str(int(x))


def _formula(sx, ref):
    """Texto da fórmula da célula `ref` com a linha trocada por {r}. Se for fórmula
    compartilhada, procura a mestra (que carrega o texto)."""
    f = re.search(r'<c r="%s"[^>]*>(<f[^>]*/>|<f[^>]*>[^<]*</f>)' % ref, sx).group(1)
    if f.endswith("/>"):
        si = re.search(r'si="(\d+)"', f).group(1)
        _, lin, txt = re.search(r'<c r="([A-Z]+)(\d+)"[^>]*><f t="shared" ref="[^"]+" si="%s">([^<]*)</f>' % si, sx).groups()
    else:
        txt, lin = re.search(r">([^<]*)</f>", f).group(1), re.search(r"(\d+)", ref).group(1)
    return re.sub(r"([A-Z]+)%s\b" % lin, r"\g<1>{r}", txt)


def _atributos(el):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', el))


def _partes(z):
    """Caminhos internos da aba Base, da aba de calendário e da tabela do calendário."""
    rels = {a["Id"]: a["Target"] for a in map(_atributos, re.findall(r"<Relationship [^>]*>", z.read("xl/_rels/workbook.xml.rels").decode("utf-8")))}
    abas = {a["name"]: "xl/" + rels[a["r:id"]] for a in map(_atributos, re.findall(r"<sheet [^>]*>", z.read("xl/workbook.xml").decode("utf-8")))}
    return abas["Base"], next(p for n, p in abas.items() if n.upper().startswith("CALEND")), next(n for n in z.namelist() if n.startswith("xl/tables/"))


def _ultima_linha_dados(sx, fim):
    for m in reversed(list(re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sx[fim - 100000:fim], flags=re.S))):
        if re.search(r'<c r="A%s"[^>]*><v>' % m.group(1), m.group(2)):
            return int(m.group(1))
    raise ValueError("nao achei a ultima linha de dados")


def _corrigir_linha(m, sufixo):
    n = int(m.group(1))
    t = re.sub(r'(<c r="X%d" s="\d+"><v>)[^<]*(</v></c>)' % n, r"\g<1>%d\g<2>" % (n - 1), m.group(0))
    return re.sub(r'(<c r="Y%d"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>))<v>[^<]*</v>' % n, r"\g<1><v>%d_%s</v>" % (n - 1, sufixo), t)


def montar(canal, atual, saida, datas=None):
    """Le `atual` (copia do empilhado) e grava `saida` com as linhas novas. Devolve um resumo.
    `datas` (DD.MM.AAAA) limita os dias do cache; dias que ja estao no empilhado sao pulados
    (rodar de novo nao duplica nada)."""
    z = zipfile.ZipFile(atual)
    base_p, cal_p, tab_p = _partes(z)
    sx = z.read(base_p).decode("utf-8")
    fim = sx.rindex("</sheetData>")
    ult = _ultima_linha_dados(sx, fim)
    sufixo = CHAVE_SUFIXO[canal]

    # 1) SEQUENCIA errada: deve ser linha-1; as erradas formam uma cauda continua
    ruins = [int(a) for a, v in re.findall(r'<c r="X(\d+)" s="\d+"><v>([^<]*)</v></c>', sx) if a != "1" and v != str(int(a) - 1)]
    if ruins:
        assert ruins == list(range(ruins[0], ult + 1)), "SEQUENCIA errada nao e uma cauda continua -- conferir a mao"
        ini = sx.index('<row r="%d"' % ruins[0])
        cauda = re.sub(r'<row r="(\d+)".*?</row>', lambda m: _corrigir_linha(m, sufixo) if int(m.group(1)) >= ruins[0] else m.group(0),
                       sx[ini:fim], flags=re.S)
        sx = sx[:ini] + cauda + sx[fim:]
        fim = sx.rindex("</sheetData>")

    # 2) linhas novas da Base (estilos e formulas copiados da ultima linha de dados)
    fim_dados = sx[fim - 200000:fim]      # celulas vazias nao existem em toda linha: pega o estilo mais recente de cada coluna
    estilo = {c: re.findall(r'<c r="%s\d+" s="(\d+)"' % c, fim_dados)[-1] for c in LETRAS}
    f_w, f_y = _formula(sx, "W%d" % ult), _formula(sx, "Y%d" % ult)
    attrs = re.search(r'<row r="%d"([^>]*)>' % ult, sx).group(1)
    arqs = sorted(glob.glob(os.path.join(CACHE, "%s_*.pkl" % canal)), key=lambda p: datetime.datetime.strptime(p[-14:-4], "%d.%m.%Y"))
    if datas:
        arqs = [p for p in arqs if p[-14:-4] in datas]
    ja_gravado = {int(v) for v in re.findall(r'<c r="A\d+" s="\d+"><v>(\d+)</v>', sx)}       # DATA_REPASSE que ja esta na Base
    dias, pulados = [], []
    for p in arqs:
        df = pickle.load(open(p, "rb"))
        (pulados if _serial(df["DATA_REPASSE"].iloc[0].date()) in ja_gravado else dias).append(df)
    pulados = [d["DATA_REPASSE"].iloc[0].strftime("%d/%m") for d in pulados]
    if not dias:
        return dict(linhas_novas=0, ja_gravados=pulados, base=base_p)
    novas, r, soma_w = [], ult, []
    for df in dias:
        soma_dia = 0.0
        for v in df.itertuples(index=False):
            r += 1
            w = float(v[8]) + float(v[9]) + float(v[11]) + float(v[10])       # I+J+L+K
            soma_dia += w
            cel = []
            for k, c in enumerate(LETRAS):
                ref, s = "%s%d" % (c, r), estilo[c]
                if c == "A":
                    cel.append('<c r="%s" s="%s"><v>%d</v></c>' % (ref, s, _serial(v[0].date())))
                elif c == "W":
                    cel.append('<c r="%s" s="%s"><f>%s</f><v>%r</v></c>' % (ref, s, f_w.format(r=r), w))
                elif c == "X":
                    cel.append('<c r="%s" s="%s"><v>%d</v></c>' % (ref, s, r - 1))
                elif c == "Y":
                    cel.append('<c r="%s" s="%s" t="str"><f>%s</f><v>%d_%s</v></c>' % (ref, s, f_y.format(r=r), r - 1, sufixo))
                elif c in COLUNAS_TEXTO:
                    t = _texto(v[k])
                    cel.append('<c r="%s" s="%s"/>' % (ref, s) if t is None else
                               '<c r="%s" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, escape(t)))
                else:
                    x = v[k]
                    cel.append('<c r="%s" s="%s"/>' % (ref, s) if x is None or pd.isna(x) else '<c r="%s" s="%s"><v>%s</v></c>' % (ref, s, _numero(x)))
            novas.append('<row r="%d"%s>%s</row>' % (r, attrs, "".join(cel)))
        soma_w.append(soma_dia)
    nova_ult = r
    sx = sx[:fim] + "".join(novas) + sx[fim:]
    sx = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: '<dimension ref="A1:%s%d"/>' % (m.group(1), nova_ult), sx, count=1)
    sx = re.sub(r'<autoFilter ref="A1:([A-Z]+)\d+"', lambda m: '<autoFilter ref="A1:%s%d"' % (m.group(1), nova_ult), sx, count=1)

    # 3) CALENDARIO: uma linha por dia com VALOR_EB no calendario principal
    cx = z.read(cal_p).decode("utf-8")
    cfim = cx.rindex("</sheetData>")
    ucal = int(re.findall(r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*?<f>SUMIFS', cx, flags=re.S)[-1])
    lcal = re.search(r'<row r="%d".*?</row>' % ucal, cx, flags=re.S).group(0)
    ec = {c: re.search(r'<c r="%s%d" s="(\d+)"' % (c, ucal), lcal).group(1) for c in "ABCDEFGHIJKL"}
    ss = {c: re.search(r'<c r="%s%d" s="\d+" t="s"><v>(\d+)</v>' % (c, ucal), lcal).group(1) for c in "DE"}
    abc = {c: re.search(r'<c r="%s%d" s="\d+"><v>([^<]*)</v>' % (c, ucal), lcal).group(1) for c in "ABC"}   # banco/conta/agencia
    attrs_cal = re.search(r'<row r="%d"([^>]*)>' % ucal, lcal).group(1)
    brancas = {int(m.group(1)): m.group(0) for m in re.finditer(r'<row r="(\d+)".*?</row>', cx[cx.index(lcal) + len(lcal):cfim], flags=re.S)}
    cx = cx[:cx.index(lcal) + len(lcal)] + cx[cfim:]      # tira as linhas em branco depois da ultima de dados (o ES tem)
    cfim = cx.rindex("</sheetData>")
    fh, fi, fj, fk = (_formula(cx, "%s%d" % (c, ucal)) for c in "HIJK")
    loja = CANAIS_KONCILI[canal]["loja"]
    cal_novas, rc, sem_cal = [], ucal, []
    for df, w_dia in zip(dias, soma_w):
        d = df["DATA_REPASSE"].iloc[0].date()
        eb = buscar_alvo_loja(loja, d.strftime("%d/%m/%Y"))
        if eb is None:
            sem_cal.append(d.strftime("%d/%m"))
            continue
        rc += 1
        j = -float(df["MP_FEE_AMOUNT"].sum())
        cal_novas.append(
            '<row r="%d"%s>' % (rc, attrs_cal)
            + '<c r="A%d" s="%s"><v>%s</v></c><c r="B%d" s="%s"><v>%s</v></c><c r="C%d" s="%s"><v>%s</v></c>' % (rc, ec["A"], abc["A"], rc, ec["B"], abc["B"], rc, ec["C"], abc["C"])
            + '<c r="D%d" s="%s" t="s"><v>%s</v></c><c r="E%d" s="%s" t="s"><v>%s</v></c>' % (rc, ec["D"], ss["D"], rc, ec["E"], ss["E"])
            + '<c r="F%d" s="%s"><v>%d</v></c><c r="G%d" s="%s"><v>%r</v></c>' % (rc, ec["F"], _serial(d), rc, ec["G"], float(eb))
            + '<c r="H%d" s="%s"><f>%s</f><v>%r</v></c>' % (rc, ec["H"], fh.format(r=rc), w_dia)
            + '<c r="I%d" s="%s"><f>%s</f><v>%r</v></c>' % (rc, ec["I"], fi.format(r=rc), float(eb) - w_dia)
            + '<c r="J%d" s="%s"><f>%s</f><v>%r</v></c>' % (rc, ec["J"], fj, j)
            + '<c r="K%d" s="%s"><f>%s</f><v>%r</v></c>' % (rc, ec["K"], fk, j / (j + w_dia))
            + '<c r="L%d" s="%s"/></row>' % (rc, ec["L"]))
    resto = "".join(x for n, x in sorted(brancas.items()) if n > rc)      # linhas em branco que ainda sobram depois das novas
    fim_cal = max([rc] + list(brancas))
    cx = cx[:cfim] + "".join(cal_novas) + resto + cx[cfim:]
    cx = re.sub(r'<dimension ref="A1:L\d+"/>', '<dimension ref="A1:L%d"/>' % fim_cal, cx, count=1)
    tx = re.sub(r"A1:L\d+", "A1:L%d" % fim_cal, z.read(tab_p).decode("utf-8"))

    # 4) filtros do workbook, sem calcChain (o Excel refaz), zip novo
    wbx = re.sub(r"(Base!\$A\$1:\$[A-Z]+\$)\d+", r"\g<1>%d" % nova_ult, z.read("xl/workbook.xml").decode("utf-8"))
    wbx = re.sub(r"(!\$A\$1:\$L\$)\d+", r"\g<1>%d" % fim_cal, wbx)
    ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", z.read("[Content_Types].xml").decode("utf-8"))
    rl = re.sub(r"<Relationship [^>]*calcChain[^>]*/>", "", z.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
    troca = {base_p: sx, cal_p: cx, tab_p: tx, "xl/workbook.xml": wbx, "[Content_Types].xml": ct, "xl/_rels/workbook.xml.rels": rl}
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as out:
        for item in z.infolist():
            if item.filename != "xl/calcChain.xml":
                out.writestr(item.filename, troca[item.filename].encode("utf-8") if item.filename in troca else z.read(item.filename))
    return dict(linhas_novas=len(novas), primeira_nova=ult + 1, ultima_linha=nova_ult, seq_corrigidas=len(ruins),
                cal_novas=len(cal_novas), ultima_cal=fim_cal, sem_calendario=sem_cal, dias=len(dias), ja_gravados=pulados, base=base_p)


def conferir(canal, novo, resumo):
    """Confere o arquivo novo LENDO de volta (openpyxl parseia o XML inteiro): tamanhos,
    SEQUENCIA/CHAVE continuas, soma de VALOR por dia == calendario."""
    import openpyxl

    z = zipfile.ZipFile(novo)
    assert "xl/calcChain.xml" not in z.namelist()
    sx = z.read(resumo["base"]).decode("utf-8")
    erradas = [a for a, v in re.findall(r'<c r="X(\d+)" s="\d+"><v>([^<]*)</v></c>', sx) if a != "1" and v != str(int(a) - 1)]
    assert not erradas, "SEQUENCIA ainda errada em %d linhas" % len(erradas)
    assert sx.count("<row ") == resumo["ultima_linha"], "quantidade de linhas diferente do esperado"
    for parte in (resumo["base"],):
        nums = [int(n) for n in re.findall(r'<row r="(\d+)"', sx)]
        assert nums == sorted(set(nums)), "linhas duplicadas/fora de ordem na Base"
    cal_p = next(n for n in z.namelist() if n.startswith("xl/worksheets/") and n != resumo["base"] and n.endswith(".xml"))
    nums = [int(n) for n in re.findall(r'<row r="(\d+)"', z.read(cal_p).decode("utf-8"))]
    assert nums == sorted(set(nums)) and nums[-1] == resumo["ultima_cal"], "linhas duplicadas/fora de ordem no calendario"
    wb = openpyxl.load_workbook(novo, read_only=True, data_only=True)
    ws = wb["Base"]
    por_dia, n, seq_ok, chave_ok = {}, 0, True, True
    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        n = i
        if i >= resumo["primeira_nova"]:
            d = row[0].date()
            por_dia[d] = por_dia.get(d, 0.0) + float(row[22])
            seq_ok &= (row[23] == i - 1)
            chave_ok &= (row[24] == "%d_%s" % (i - 1, CHAVE_SUFIXO[canal]))
    assert n == resumo["ultima_linha"] and seq_ok and chave_ok, (n, seq_ok, chave_ok)
    loja = CANAIS_KONCILI[canal]["loja"]
    out = []
    for d, v in sorted(por_dia.items()):
        eb = buscar_alvo_loja(loja, d.strftime("%d/%m/%Y"))
        out.append((d, round(v, 2), eb))
        assert eb is None or abs(v - eb) < 0.01, "soma do dia %s difere do calendario" % d
    return out


def main(argv, datas=None):
    canal = argv[0].upper()
    publicar = "--publicar" in argv
    origem = os.path.join(PASTA_EMPILHADO, ARQUIVO[canal])
    if os.path.exists(os.path.join(PASTA_EMPILHADO, "~$" + ARQUIVO[canal])):
        sys.exit("ARQUIVO ABERTO por alguem (~$ na pasta) -- feche e rode de novo")
    os.makedirs(PASTA_BACKUP, exist_ok=True)
    atual = os.path.join(PASTA_BACKUP, "_atual_%s.xlsx" % canal)
    shutil.copy2(origem, atual)                     # copia de trabalho = o que esta no X: agora
    datado = os.path.join(PASTA_BACKUP, "%s - backup %s.xlsx" % (ARQUIVO[canal][:-5], datetime.datetime.now().strftime("%d.%m.%Y %Hh%M")))
    if not any(os.path.getsize(p) == os.path.getsize(origem) for p in glob.glob(os.path.join(PASTA_BACKUP, ARQUIVO[canal][:-5] + " - backup*.xlsx"))):
        shutil.copy2(origem, datado)                # o X: mudou desde o ultimo backup -> guarda esse tambem
        print("backup novo:", datado)
    novo = os.path.join(PASTA_BACKUP, "_novo_%s.xlsx" % canal)
    print("montando arquivo novo...", flush=True)
    resumo = montar(canal, atual, novo, datas)
    print("resumo:", {k: v for k, v in resumo.items() if k != "base"}, flush=True)
    if resumo["linhas_novas"] == 0:
        print("Nada novo pra gravar (os dias ja estao no empilhado ou nao ha dias no cache). Nada foi trocado.")
        return
    print("conferindo (le o arquivo novo de volta)...", flush=True)
    for d, v, eb in conferir(canal, novo, resumo):
        print("   %s  soma VALOR %14s  calendario %s" % (d.strftime("%d/%m"), format(v, ",.2f"), "sem linha" if eb is None else "OK"))
    if not publicar:
        print("OK local. Nada foi trocado no X:. Para publicar: --publicar")
        return
    if os.path.exists(os.path.join(PASTA_EMPILHADO, "~$" + ARQUIVO[canal])) or os.path.getsize(origem) != os.path.getsize(atual):
        sys.exit("o arquivo do X: mudou/abriu durante a montagem -- nao publiquei")
    tmp = origem + ".novo"
    shutil.copy2(novo, tmp)
    os.replace(tmp, origem)
    print("PUBLICADO:", origem)
    feito = os.path.join(CACHE, "publicado")                 # o cache do que ja foi gravado sai do caminho
    os.makedirs(feito, exist_ok=True)
    for p in glob.glob(os.path.join(CACHE, "%s_*.pkl" % canal)):
        if not datas or p[-14:-4] in datas:
            shutil.move(p, os.path.join(feito, os.path.basename(p)))


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1].upper() not in ARQUIVO:
        sys.exit("uso: python empilhar_gravar.py <FULL|ES> [--publicar]")
    main(sys.argv[1:])
