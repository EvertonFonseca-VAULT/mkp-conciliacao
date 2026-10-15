"""Empilhado dos marketplaces de repasse fechado (LIVELO, COLOMBO e os que vierem depois) --
nenhum deles tem cruzamento diario com TOTVS/AB pra resolver numero TOTVS (por isso NAO e o
modulo do empilhar_canais.py -- familia Mercado, ver ADR-0001). O que varia de verdade entre eles
e so isto: de onde vem a data de cada periodo (de dentro do arquivo -- LIVELO -- ou do nome do
arquivo -- COLOMBO) e se tem ou nao um resumo (IR/Comissao/Repasse) embutido. Cadencia (mensal ou
diaria) e so consequencia disso, nao e um eixo separado.

Modulo profundo: varrer pasta, resolver a data real de cada arquivo candidato, decidir o que
falta (dados e/ou calendario) comparando por data (nao por mes -- cadencia diaria tem mais de um
periodo pendente no mesmo mes), escrever, gravar backup. Cada chave (marketplace+loja, ex.
LIVELO_CENTER, COLOMBO_CENTER) declara: como resolver a data de um arquivo a partir do nome,
como extrair as linhas (e o resumo, se tiver), as colunas do empilhado (tabela declarativa) e a
funcao pura que monta a linha do calendario.

Uso:
    python empilhar_repasses.py CHAVE           # empilha o que faltar (dados e/ou calendario)
    python empilhar_repasses.py CHAVE --teste    # roda o teste offline de 1 chave
    python empilhar_repasses.py --teste          # roda os testes offline de cada chave
"""
import datetime
import pathlib
import re
import shutil
import sys

import openpyxl
from openpyxl.utils import get_column_letter

BASE_TARGET = pathlib.Path(
    "X:/109_GC_CTR_FINANCEIRO CONTAS A RECEBER/MKP/ARQUIVOS EMPILHADOS DE BAIXAS/"
    "ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS"
)

_EXCLUIR_NOME = re.compile(r"c[óo]pia|fechamento", re.I)
_EXCLUIR_ARQUIVO = re.compile(r"^~\$")  # lock temporario do Excel (arquivo aberto por alguem)


def _parse_data(s):
    if s is None or isinstance(s, datetime.datetime):
        return s
    return datetime.datetime.strptime(s, "%Y-%m-%d")


# ======================================================================================
# descoberta de periodo pendente -- generico: cada chave so declara como resolver a data
# de um arquivo a partir do nome (sem abrir o arquivo -- so olhando pasta/nome)
# ======================================================================================
def listar_periodos(pasta_fonte, padrao_arquivo, resolver_data, janela_a_partir_de=None, meses_de_tolerancia=2):
    """{data: (caminho_ou_None, candidatos_ambiguos_ou_None)} -- um item por data resolvida.
    None no caminho quer dizer que 2+ arquivos resolveram pra mesma data (so importa se a data
    for mesmo pendente -- quem decide isso e o main()).

    So entra pasta de mes dentro da janela (mes de `janela_a_partir_de` menos `meses_de_tolerancia`
    em diante) -- sem isso, um buraco de calendario de anos atras (nunca preenchido de proposito)
    vira "pendente" so porque a data nao esta lá, e um dia esquecido recente (ex. dia 21 chegando
    depois do dia 30 do mesmo mes) precisa continuar alcancavel. `janela_a_partir_de=None` nao
    filtra nada (primeira carga, sem nada empilhado ainda)."""
    limite_mes = None
    if janela_a_partir_de is not None:
        limite_mes = janela_a_partir_de.year * 12 + janela_a_partir_de.month - meses_de_tolerancia

    candidatos = {}
    for pasta_ano in sorted(pasta_fonte.glob("[0-9][0-9][0-9][0-9]")):
        for pasta_mes in sorted(pasta_ano.glob("*.????")):
            m = re.match(r"(\d\d)\.(\d\d\d\d)$", pasta_mes.name)
            if not m:
                continue
            mes, ano = int(m.group(1)), int(m.group(2))
            if limite_mes is not None and (ano * 12 + mes) < limite_mes:
                continue
            for f in pasta_mes.glob("*.xlsx"):
                if _EXCLUIR_ARQUIVO.match(f.name) or _EXCLUIR_NOME.search(f.stem) or not padrao_arquivo.search(f.stem):
                    continue
                data = resolver_data(f, mes, ano)
                if data is not None:
                    candidatos.setdefault(data, []).append(f)
    return {
        data: (arqs[0] if len(arqs) == 1 else None, arqs if len(arqs) > 1 else None)
        for data, arqs in sorted(candidatos.items())
    }


# ======================================================================================
# linha do empilhado -- tabela declarativa de colunas, identico em mecanica pra qualquer chave
# campo None = constante; campo inteiro = indice posicional em campos["_linha_origem"];
# campo string = chave do dict de campos (sempre tem "_data_periodo" disponivel)
# ======================================================================================
def montar_valores_linha(colunas, campos):
    valores = []
    for col, campo, fmt, eh_data in colunas:
        if campo is None:
            v = None
        elif isinstance(campo, int):
            v = campos["_linha_origem"][campo]
        else:
            v = campos[campo]
        if eh_data:
            v = _parse_data(v)
        valores.append((col, v, fmt))
    return valores


def _escrever(ws, linha, valores):
    for col, v, fmt in valores:
        c = ws.cell(linha, col, v)
        if fmt:
            c.number_format = fmt


def _ultima_linha_com_dados(ws, col):
    ultima = 1
    idx = col - 1
    for row in ws.iter_rows(min_row=2):
        if row[idx].value is not None:
            ultima = row[idx].row
    return ultima


def _ultima_data(ws, col):
    ultima = None
    idx = col - 1
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = row[idx]
        if isinstance(d, datetime.datetime) and (ultima is None or d > ultima):
            ultima = d
    return ultima


def _datas_existentes(ws, col):
    """Datas ja gravadas (nao so a ultima) -- um dia preenchido depois de um buraco (ex. dia 21
    chegando depois do dia 30 ja empilhado) precisa continuar pendente, nao so o que e > ultima."""
    idx = col - 1
    datas = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = row[idx]
        if isinstance(d, datetime.datetime):
            datas.add(d)
    return datas


# ======================================================================================
# principal -- generico, despacha pro cadastro da chave
# ======================================================================================
def main(chave, argv):
    cfg = REPASSES[chave]
    wb = openpyxl.load_workbook(cfg["arquivo"])
    ws = wb[cfg["aba_dados"]]
    ws_cal = wb[cfg["aba_calendario"]]

    datas_dados = _datas_existentes(ws, 1)
    datas_cal = _datas_existentes(ws_cal, cfg["col_data_cal"])
    print("%s -- ultima data empilhada: %s | ultima no calendario: %s" % (
        chave, max(datas_dados).date() if datas_dados else "nenhuma", max(datas_cal).date() if datas_cal else "nenhuma"))

    datas_conhecidas = datas_dados | datas_cal
    janela = max(datas_conhecidas) if datas_conhecidas else None
    periodos = listar_periodos(cfg["pasta_fonte"], cfg["padrao_arquivo"], cfg["resolver_periodo"], janela)
    pendentes = {
        data: (arq, cand)
        for data, (arq, cand) in periodos.items()
        if data not in datas_dados or data not in datas_cal
    }
    if not pendentes:
        print("nada novo pra empilhar (dados e calendario em dia).")
        wb.close()
        return

    ambiguos = {data: cand for data, (arq, cand) in pendentes.items() if arq is None}
    if ambiguos:
        wb.close()
        sys.exit("\n".join(
            "%s: %d arquivos candidatos, nao consegui decidir -- %s" % (data.date(), len(cand), cand)
            for data, cand in sorted(ambiguos.items())
        ))

    backup = cfg["arquivo"].with_name(
        cfg["arquivo"].stem + "_BACKUP_" + datetime.date.today().strftime("%Y%m%d") + cfg["arquivo"].suffix
    )
    shutil.copy2(cfg["arquivo"], backup)
    print("backup:", backup)

    n_colunas = len(cfg["colunas"])
    col_seq, col_chave = n_colunas + 1, n_colunas + 2
    letra_seq = get_column_letter(col_seq)
    linha_dados = _ultima_linha_com_dados(ws, 1)
    linha_cal = _ultima_linha_com_dados(ws_cal, cfg["col_data_cal"])

    for data in sorted(pendentes):
        arq, _cand = pendentes[data]
        precisa_dados = data not in datas_dados
        precisa_cal = data not in datas_cal

        linhas, dados_periodo = cfg["extrair"](arq, data)
        if not linhas:
            print("  %s: 0 linhas de detalhe, pulando" % arq.name)
            continue

        if precisa_dados:
            for campos in linhas:
                nova = linha_dados + 1
                _escrever(ws, nova, montar_valores_linha(cfg["colunas"], campos))
                ws.cell(nova, col_seq, cfg["formula_sequencia"].format(ant=linha_dados, letra_seq=letra_seq))
                ws.cell(nova, col_chave, cfg["formula_chave"].format(nova=nova, letra_seq=letra_seq))
                linha_dados = nova
            print("  %s: +%d linhas em %s" % (arq.name, len(linhas), cfg["aba_dados"]))

        if precisa_cal:
            nova_cal = linha_cal + 1
            anterior = [c.value for c in ws_cal[linha_cal]] if linha_cal > 1 else []
            _escrever(ws_cal, nova_cal, cfg["calendario"](dados_periodo, nova_cal, anterior))
            linha_cal = nova_cal
            print("  %s: calendario linha %d (conferir manual o valor do extrato bancario)" % (arq.name, nova_cal))

    wb.save(cfg["arquivo"])
    print("%s: ok, ultima linha dados=%d, ultima linha calendario=%d" % (chave, linha_dados, linha_cal))


# ======================================================================================
# LIVELO -- resumo embutido no arquivo mensal, data do periodo vem do proprio resumo
# ======================================================================================
_CAMPOS_DETALHE_LIVELO = (
    "data_pedido", "pedido", "entrega", "ordem", "qtdd",
    "vtotal", "vprod", "frete", "comissao", "vliq", "data_repasse",
)
_CAMPOS_RESUMO_LIVELO = (
    "nome_seller", "_", "cnpj", "_2", "total_vendas", "_3",
    "comissao_tot", "ir", "campanhas", "valor_repasse", "data_repasse",
)


def _resolver_periodo_livelo(caminho, mes, ano):
    """So aceita o arquivo cujo mes (dd.mm) no proprio nome bate com o mes da pasta -- a pasta do
    mes corrente as vezes guarda sobra do arquivo do mes anterior."""
    m = re.search(r"\d{1,2}\.0*(\d{1,2})\b", caminho.stem)
    if not m or int(m.group(1)) != mes:
        return None
    import calendar
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return datetime.datetime(ano, mes, ultimo_dia)


def _extrair_livelo(caminho, data_periodo, header_row=4):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    linhas = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if row[0] is None and row[1] is None:
            break
        campos = dict(zip(_CAMPOS_DETALHE_LIVELO, row[:11]))
        campos["_data_periodo"] = data_periodo
        linhas.append(campos)
    resumo = {"_data_periodo": data_periodo}
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        resumo.update(dict(zip(_CAMPOS_RESUMO_LIVELO, row)))
    wb.close()
    return linhas, resumo


def _calendario_livelo_center(dados_periodo, nova, anterior):
    banco, conta, agencia = (anterior[0], anterior[1], anterior[2]) if anterior else (None, None, None)
    return [
        (1, banco, None), (2, conta, None), (3, agencia, None),
        (4, "LIVELO", None), (5, "CENTER", None),
        (6, dados_periodo["_data_periodo"], "mm-dd-yy"),
        # VALOR_EB (col 7) fica em branco: valor do extrato bancario, preencher manual
        (8, "=SUMIFS(EMPILHADO!L:L,EMPILHADO!A:A,CALENDARIO!F%d)" % nova, "#,##0.00"),
        (9, "=G%d-H%d" % (nova, nova), "#,##0.00"),
        (10, dados_periodo["ir"], "#,##0.00"),
        (11, dados_periodo["campanhas"], "#,##0.00"),
        (12, "Faturamento Campanhas+ IR" if (dados_periodo["ir"] or dados_periodo["campanhas"]) else None, None),
    ]


def _calendario_livelo_lg(dados_periodo, nova, anterior):
    return [
        (2, "LIVELO", None), (3, "LG", None),
        (4, dados_periodo["_data_periodo"], "mm-dd-yy"),
        # Repasse (col 5 / E) fica em branco: valor do extrato bancario, preencher manual
        (6, dados_periodo["valor_repasse"], "#,##0.00"),
        (7, "=E%d-SUM(F%d,H%d,I%d)" % (nova, nova, nova, nova), "#,##0.00"),
        (8, dados_periodo["ir"], "#,##0.00"),
        (9, dados_periodo["campanhas"], "#,##0.00"),
    ]


# (col, campo, numero_formato, eh_data) -- campo None = constante (CNPJ/Nome Seller nao vem no arquivo mensal)
_COLUNAS_LIVELO_CENTER = [
    (1, "_data_periodo", "mm-dd-yy", True),
    (2, "data_pedido", None, False),
    (3, "pedido", None, False),
    (4, "entrega", None, False),
    (5, "ordem", None, False),
    (6, None, None, False),
    (7, None, None, False),
    (8, "vtotal", None, False),
    (9, "vprod", None, False),
    (10, "frete", None, False),
    (11, "comissao", None, False),
    (12, "vliq", None, False),
]
_COLUNAS_LIVELO_LG = [
    (1, "_data_periodo", "mm-dd-yy", True),
    (2, "data_pedido", "mm-dd-yy", False),
    (3, "pedido", None, False),
    (4, "entrega", None, False),
    (5, "ordem", "mm-dd-yy", False),
    (6, "qtdd", "mm-dd-yy", False),
    (7, "vtotal", None, False),
    (8, "vprod", None, False),
    (9, "frete", None, False),
    (10, "comissao", None, False),
    (11, "vliq", None, False),
    (12, "data_repasse", None, False),  # mesma data, como texto (igual ja esta gravado)
]


# ======================================================================================
# COLOMBO -- sem resumo, arquivo e so a exportacao diaria "AB"; data do periodo vem do nome
# do arquivo (nao tem coluna de data de repasse dentro do arquivo -- a data E a captura do dia)
# ======================================================================================
def _resolver_periodo_colombo(caminho, mes, ano):
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", caminho.stem)
    if not m:
        return None
    dia, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if mo != mes or y != ano:
        return None
    return datetime.datetime(y, mo, dia)


def _extrair_colombo(caminho, data_periodo):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    linhas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        linhas.append({"_linha_origem": row[:33], "_data_periodo": data_periodo})
    wb.close()
    return linhas, {"_data_periodo": data_periodo}


def _calendario_colombo_center(dados_periodo, nova, anterior):
    banco, conta, agencia, mkt, loja = (anterior[1], anterior[2], anterior[3], anterior[4], anterior[5]) if anterior else (None,) * 5
    return [
        (2, banco, None), (3, conta, None), (4, agencia, None),
        (5, mkt, None), (6, loja, None),
        (7, dados_periodo["_data_periodo"], "mm-dd-yy"),
        # VALOR_EB (col 8 / H) fica em branco: valor do extrato bancario, preencher manual
        (9, "=SUMIFS(CENTER!AF:AF,CENTER!A:A,Calend\u00e1rio!G%d)" % nova, "#,##0.00"),
        (11, "=H%d-I%d" % (nova, nova), "0.00"),
    ]


# 33 colunas do arquivo diario "AB" do COLOMBO, copiadas por posicao (0-based) -- so a coluna 1
# (Data Repasse) nao vem do arquivo, vem do nome dele (ver _resolver_periodo_colombo)
_COLS_DATA_COLOMBO = {5, 6, 7, 8, 26, 27, 31}  # posicoes (0-based) que sao data no arquivo de origem
_COLUNAS_COLOMBO_CENTER = [(1, "_data_periodo", "mm-dd-yy", True)] + [
    (i + 2, i, ("mm-dd-yy" if i in _COLS_DATA_COLOMBO else None), False) for i in range(33)
]


# ======================================================================================
# cadastro das chaves -- o que varia entre LIVELO_CENTER, LIVELO_LG, COLOMBO_CENTER e as
# que vierem depois. BASE_TARGET e sempre a mesma pasta; pasta_fonte e o caminho completo
# de cada marketplace (bases diferentes).
# ======================================================================================
_BASE_LIVELO = pathlib.Path("X:/109_GC_CTR_FINANCEIRO CONTAS A RECEBER/MKP/MKP - LIVELO")
_BASE_COLOMBO = pathlib.Path("X:/109_GC_CTR_FINANCEIRO CONTAS A RECEBER/MKP/MKP - COLOMBO")

REPASSES = {
    "LIVELO_CENTER": dict(
        pasta_fonte=_BASE_LIVELO / "CENTER", padrao_arquivo=re.compile(r"livelo.*center", re.I),
        resolver_periodo=_resolver_periodo_livelo, extrair=_extrair_livelo,
        arquivo=BASE_TARGET / "CENTER_LIVELO.xlsx", aba_dados="EMPILHADO", aba_calendario="CALENDARIO",
        col_data_cal=6, colunas=_COLUNAS_LIVELO_CENTER,
        formula_sequencia="={letra_seq}{ant}+1", formula_chave='={letra_seq}{nova}&"_LIVELO_CENTER"',
        calendario=_calendario_livelo_center,
    ),
    "LIVELO_LG": dict(
        pasta_fonte=_BASE_LIVELO / "LG", padrao_arquivo=re.compile(r"livelo.*lg", re.I),
        resolver_periodo=_resolver_periodo_livelo, extrair=_extrair_livelo,
        arquivo=BASE_TARGET / "LG_LIVELO.xlsx", aba_dados="Base", aba_calendario="Calendario",
        col_data_cal=4, colunas=_COLUNAS_LIVELO_LG,
        formula_sequencia="={letra_seq}{ant}+1", formula_chave='=CONCATENATE({letra_seq}{nova},"_","LG_LIVELO")',
        calendario=_calendario_livelo_lg,
    ),
    "COLOMBO_CENTER": dict(
        pasta_fonte=_BASE_COLOMBO / "CENTER" / "AB", padrao_arquivo=re.compile(r"colombo.*center", re.I),
        resolver_periodo=_resolver_periodo_colombo, extrair=_extrair_colombo,
        arquivo=BASE_TARGET / "CENTER_COLOMBO.xlsx", aba_dados="CENTER", aba_calendario="Calend\u00e1rio",
        col_data_cal=7, colunas=_COLUNAS_COLOMBO_CENTER,
        formula_sequencia="={letra_seq}{ant}+1", formula_chave='={letra_seq}{nova}&"_CENTER_COLOMBO"',
        calendario=_calendario_colombo_center,
    ),
}


# ======================================================================================
# testes offline (um por chave, --teste roda todas) -- so a mecanica pura: montar_valores_linha
# e a funcao de calendario. Sem abrir xlsx nenhum.
# ======================================================================================
def _teste_livelo_center():
    campos = dict(data_pedido="2026-08-01", pedido="L01", entrega="E01", ordem="O01", qtdd="1",
                  vtotal=100.0, vprod=90.0, frete=10.0, comissao=-15.0, vliq=75.0,
                  _data_periodo=datetime.datetime(2026, 9, 30))
    valores = montar_valores_linha(_COLUNAS_LIVELO_CENTER, campos)
    assert valores[0] == (1, datetime.datetime(2026, 9, 30), "mm-dd-yy")
    assert valores[1] == (2, "2026-08-01", None)
    assert valores[5] == (6, None, None) and valores[6] == (7, None, None)
    assert valores[-1] == (12, 75.0, None)

    dados_periodo = dict(ir=10.5, campanhas=-3.2, _data_periodo=datetime.datetime(2026, 9, 30))
    linhas = _calendario_livelo_center(dados_periodo, 38, [1, "210-1", 3416])
    d = {c: (v, f) for c, v, f in linhas}
    assert d[1] == (1, None) and d[2] == ("210-1", None) and d[3] == (3416, None)
    assert d[8] == ("=SUMIFS(EMPILHADO!L:L,EMPILHADO!A:A,CALENDARIO!F38)", "#,##0.00")
    assert d[10] == (10.5, "#,##0.00") and d[12][0] == "Faturamento Campanhas+ IR"

    assert _resolver_periodo_livelo(pathlib.Path("LIVELO CENTER 30.09.xlsx"), 9, 2026) == datetime.datetime(2026, 9, 30)
    assert _resolver_periodo_livelo(pathlib.Path("LIVELO CENTER 31.08.xlsx"), 9, 2026) is None  # sobra do mes anterior
    print("empilhar_repasses.py [LIVELO_CENTER]: ok")


def _teste_livelo_lg():
    dados_periodo = dict(ir=5.0, campanhas=0.0, valor_repasse=1234.56, _data_periodo=datetime.datetime(2026, 9, 30))
    linhas = _calendario_livelo_lg(dados_periodo, 26, [])
    d = {c: (v, f) for c, v, f in linhas}
    assert d[2] == ("LIVELO", None) and d[3] == ("LG", None)
    assert d[6] == (1234.56, "#,##0.00")
    assert d[7] == ("=E26-SUM(F26,H26,I26)", "#,##0.00")

    campos = dict(data_pedido="2026-08-01", pedido="L01", entrega="E01", ordem="O01", qtdd="1",
                  vtotal=100.0, vprod=90.0, frete=10.0, comissao=-15.0, vliq=75.0,
                  data_repasse="2026-09-30", _data_periodo=datetime.datetime(2026, 9, 30))
    valores = montar_valores_linha(_COLUNAS_LIVELO_LG, campos)
    assert valores[5] == (6, "1", "mm-dd-yy")       # qtdd preservado (so no LG)
    assert valores[-1] == (12, "2026-09-30", None)   # segunda data, como texto
    print("empilhar_repasses.py [LIVELO_LG]: ok")


def _teste_colombo_center():
    linha = tuple(range(33))  # 33 valores de origem, so pra conferir posicao
    campos = {"_linha_origem": linha, "_data_periodo": datetime.datetime(2026, 9, 22)}
    valores = montar_valores_linha(_COLUNAS_COLOMBO_CENTER, campos)
    assert valores[0] == (1, datetime.datetime(2026, 9, 22), "mm-dd-yy")
    assert valores[1] == (2, 0, None)          # posicao 0 da origem -> coluna 2
    assert valores[6] == (7, 5, "mm-dd-yy")    # posicao 5 (Data Pedido) -> coluna 7, com formato de data
    assert valores[-1] == (34, 32, None)       # ultima coluna de origem -> coluna 34

    linhas = _calendario_colombo_center({"_data_periodo": datetime.datetime(2026, 9, 22)}, 1428,
                                         [None, 237, 103925, 3708, "COLOMBO", "CENTER"])
    d = {c: (v, f) for c, v, f in linhas}
    assert d[2] == (237, None) and d[5] == ("COLOMBO", None) and d[6] == ("CENTER", None)
    assert d[9] == ("=SUMIFS(CENTER!AF:AF,CENTER!A:A,Calend\u00e1rio!G1428)", "#,##0.00")
    assert d[11] == ("=H1428-I1428", "0.00")

    assert _resolver_periodo_colombo(pathlib.Path("COLOMBO CENTER 22.09.2026.xlsx"), 9, 2026) == datetime.datetime(2026, 9, 22)
    assert _resolver_periodo_colombo(pathlib.Path("COLOMBO CENTER 22.09.2026.xlsx"), 8, 2026) is None
    print("empilhar_repasses.py [COLOMBO_CENTER]: ok")


TESTES = {"LIVELO_CENTER": _teste_livelo_center, "LIVELO_LG": _teste_livelo_lg, "COLOMBO_CENTER": _teste_colombo_center}


def teste(chave):
    TESTES[chave]()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in REPASSES:
        _chave, _resto = sys.argv[1], sys.argv[2:]
        if "--teste" in _resto:
            teste(_chave)
        else:
            main(_chave, _resto)
    elif "--teste" in sys.argv:
        for _chave in TESTES:
            teste(_chave)
    else:
        sys.exit(
            "uso: python empilhar_repasses.py <CHAVE>  (empilha o que faltar, grava direto)\n"
            "     python empilhar_repasses.py <CHAVE> --teste\n"
            "     python empilhar_repasses.py --teste  (roda todas as chaves)\n"
            f"chaves: {', '.join(REPASSES)}"
        )
