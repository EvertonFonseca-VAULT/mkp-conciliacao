"""Empilhado dos marketplaces de repasse fechado (LIVELO e os que vierem depois) -- cada um manda
um arquivo mensal ja fechado (com o proprio resumo de IR/Comissao/Repasse), sem cruzamento diario
com TOTVS/AB. Por isso NAO e o mesmo modulo do empilhar_canais.py (familia Mercado): aqui nao ha
numero TOTVS pra resolver, a granularidade e mensal, e o "calendario" de cada chave carrega os
dados do proprio resumo do arquivo em vez de vir do extrato bancario dia a dia.

Modulo profundo: a mecanica comum (varrer pasta, achar arquivo do mes, extrair linhas de
detalhe, escrever no empilhado, decidir se falta so o calendario) fica aqui. Cada chave
(marketplace+loja, ex. LIVELO_CENTER) so declara o que varia: colunas do empilhado, formula de
Sequencia/Chave (a convencao ja gravada no arquivo, nao inventamos uma nova) e a funcao que monta
a linha do calendario a partir do resumo.

Uso (via wrapper fino, um por chave -- ver empilhar_livelo_center.py etc.):
    python empilhar_livelo_center.py           # empilha o que faltar (dados e/ou calendario)
    python empilhar_repasses.py --teste         # roda os testes offline de cada chave
"""
import datetime
import pathlib
import re
import shutil
import sys

import openpyxl
from openpyxl.utils import get_column_letter

BASE_SRC = pathlib.Path("X:/109_GC_CTR_FINANCEIRO CONTAS A RECEBER/MKP/MKP - LIVELO")
BASE_TARGET = pathlib.Path(
    "X:/109_GC_CTR_FINANCEIRO CONTAS A RECEBER/MKP/ARQUIVOS EMPILHADOS DE BAIXAS/"
    "ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS"
)

_EXCLUIR_NOME = re.compile(r"c[óo]pia|fechamento", re.I)


# ======================================================================================
# leitura do arquivo mensal de origem -- identico pra qualquer chave da familia
# ======================================================================================
_CAMPOS_DETALHE = (
    "data_pedido", "pedido", "entrega", "ordem", "qtdd",
    "vtotal", "vprod", "frete", "comissao", "vliq", "data_repasse",
)
_CAMPOS_RESUMO = (
    "nome_seller", "_", "cnpj", "_2", "total_vendas", "_3",
    "comissao_tot", "ir", "campanhas", "valor_repasse", "data_repasse",
)


def _parse_data(s):
    if isinstance(s, datetime.datetime):
        return s
    return datetime.datetime.strptime(s, "%Y-%m-%d")


def extrair_mes(caminho, header_row=4):
    """(linhas_de_detalhe, resumo) do arquivo mensal. linhas = lista de dict por _CAMPOS_DETALHE."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    linhas = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if row[0] is None and row[1] is None:
            break
        linhas.append(dict(zip(_CAMPOS_DETALHE, row[:11])))
    resumo = None
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        resumo = dict(zip(_CAMPOS_RESUMO, row))
    wb.close()
    return linhas, resumo


# ======================================================================================
# descoberta de periodo pendente -- generico: cada chave so declara o padrao do nome do arquivo
# ======================================================================================
def listar_periodos(pasta_fonte, padrao_arquivo):
    """[(ano, mes, caminho_ou_None, candidatos_ambiguos)] em <pasta_fonte>/<ano>/<mes.ano>/.
    caminho vem None quando a pasta tem mais de 1 candidato mesmo apos o desempate por mes --
    quem decide se isso e erro e o main() (so importa se o periodo for mesmo pendente)."""
    periodos = []
    for pasta_ano in sorted(pasta_fonte.glob("[0-9][0-9][0-9][0-9]")):
        for pasta_mes in sorted(pasta_ano.glob("*.????")):
            m = re.match(r"(\d\d)\.(\d\d\d\d)$", pasta_mes.name)
            if not m:
                continue
            mes, ano = int(m.group(1)), int(m.group(2))
            candidatos = [
                f for f in pasta_mes.glob("*.xlsx")
                if not _EXCLUIR_NOME.search(f.stem) and padrao_arquivo.search(f.stem)
            ]
            if len(candidatos) > 1:
                # a pasta do mes corrente as vezes guarda sobra do arquivo do mes anterior;
                # desempata pelo mes (dd.mm) que aparece no proprio nome
                bate = [f for f in candidatos if re.search(rf"\d{{1,2}}\.0*{mes}\b", f.stem)]
                if len(bate) == 1:
                    candidatos = bate
            if len(candidatos) == 1:
                periodos.append((ano, mes, candidatos[0], None))
            elif len(candidatos) > 1:
                periodos.append((ano, mes, None, candidatos))
    return sorted(periodos, key=lambda p: (p[0], p[1]))


# ======================================================================================
# linha do empilhado -- tabela declarativa de colunas, identico em mecanica pra qualquer chave
# ======================================================================================
def montar_valores_linha(colunas, campos):
    """[(col, valor, numero_formato)] a partir da tabela declarativa e do dict de campos do detalhe."""
    valores = []
    for col, campo, fmt, eh_data in colunas:
        if campo is None:
            v = None
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


# ======================================================================================
# principal -- generico, despacha pro cadastro da chave
# ======================================================================================
def main(chave, argv):
    cfg = REPASSES[chave]
    wb = openpyxl.load_workbook(cfg["arquivo"])
    ws = wb[cfg["aba_dados"]]
    ws_cal = wb[cfg["aba_calendario"]]

    ultima_dados = _ultima_data(ws, 1)
    ultima_cal = _ultima_data(ws_cal, cfg["col_data_cal"])
    print("%s -- ultima data empilhada: %s | ultima no calendario: %s" % (
        chave, ultima_dados.date() if ultima_dados else "nenhuma", ultima_cal.date() if ultima_cal else "nenhuma"))

    periodos = listar_periodos(cfg["pasta_fonte"], cfg["padrao_arquivo"])
    pendentes = [
        (ano, mes, arq, candidatos,
         ultima_dados is None or (ano, mes) > (ultima_dados.year, ultima_dados.month),
         ultima_cal is None or (ano, mes) > (ultima_cal.year, ultima_cal.month))
        for ano, mes, arq, candidatos in periodos
    ]
    pendentes = [p for p in pendentes if p[4] or p[5]]
    if not pendentes:
        print("nada novo pra empilhar (dados e calendario em dia).")
        wb.close()
        return

    ambiguos = [(ano, mes, cand) for ano, mes, arq, cand, _, _ in pendentes if arq is None]
    if ambiguos:
        wb.close()
        sys.exit("\n".join(
            "%02d.%d: %d arquivos candidatos, nao consegui decidir -- %s" % (mes, ano, len(cand), cand)
            for ano, mes, cand in ambiguos
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

    for ano, mes, arq, _cand, precisa_dados, precisa_cal in pendentes:
        linhas, resumo = extrair_mes(arq)
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
            _escrever(ws_cal, nova_cal, cfg["calendario"](resumo, nova_cal, anterior))
            linha_cal = nova_cal
            print("  %s: calendario linha %d (conferir manual o valor do extrato bancario)" % (arq.name, nova_cal))

    wb.save(cfg["arquivo"])
    print("%s: ok, ultima linha dados=%d, ultima linha calendario=%d" % (chave, linha_dados, linha_cal))


# ======================================================================================
# calendario por chave -- funcao pura: recebe dados, devolve [(col, valor, fmt)], nao toca planilha
# ======================================================================================
def _calendario_livelo_center(resumo, nova, anterior):
    banco, conta, agencia = (anterior[0], anterior[1], anterior[2]) if anterior else (None, None, None)
    return [
        (1, banco, None), (2, conta, None), (3, agencia, None),
        (4, "LIVELO", None), (5, "CENTER", None),
        (6, _parse_data(resumo["data_repasse"]), "mm-dd-yy"),
        # VALOR_EB (col 7) fica em branco: valor do extrato bancario, preencher manual
        (8, "=SUMIFS(EMPILHADO!L:L,EMPILHADO!A:A,CALENDARIO!F%d)" % nova, "#,##0.00"),
        (9, "=G%d-H%d" % (nova, nova), "#,##0.00"),
        (10, resumo["ir"], "#,##0.00"),
        (11, resumo["campanhas"], "#,##0.00"),
        (12, "Faturamento Campanhas+ IR" if (resumo["ir"] or resumo["campanhas"]) else None, None),
    ]


def _calendario_livelo_lg(resumo, nova, anterior):
    return [
        (2, "LIVELO", None), (3, "LG", None),
        (4, _parse_data(resumo["data_repasse"]), "mm-dd-yy"),
        # Repasse (col 5 / E) fica em branco: valor do extrato bancario, preencher manual
        (6, resumo["valor_repasse"], "#,##0.00"),
        (7, "=E%d-SUM(F%d,H%d,I%d)" % (nova, nova, nova, nova), "#,##0.00"),
        (8, resumo["ir"], "#,##0.00"),
        (9, resumo["campanhas"], "#,##0.00"),
    ]


# ======================================================================================
# cadastro das chaves -- o que varia entre LIVELO_CENTER, LIVELO_LG e as que vierem depois
# ======================================================================================
# (col, campo, numero_formato, eh_data) -- campo None = constante (CNPJ/Nome Seller nao vem no arquivo mensal)
_COLUNAS_CENTER = [
    (1, "data_repasse", "mm-dd-yy", True),
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
_COLUNAS_LG = [
    (1, "data_repasse", "mm-dd-yy", True),
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

REPASSES = {
    "LIVELO_CENTER": dict(
        pasta_fonte=BASE_SRC / "CENTER", padrao_arquivo=re.compile(r"livelo.*center", re.I),
        arquivo=BASE_TARGET / "CENTER_LIVELO.xlsx", aba_dados="EMPILHADO", aba_calendario="CALENDARIO",
        col_data_cal=6, colunas=_COLUNAS_CENTER,
        formula_sequencia="={letra_seq}{ant}+1", formula_chave='={letra_seq}{nova}&"_LIVELO_CENTER"',
        calendario=_calendario_livelo_center,
    ),
    "LIVELO_LG": dict(
        pasta_fonte=BASE_SRC / "LG", padrao_arquivo=re.compile(r"livelo.*lg", re.I),
        arquivo=BASE_TARGET / "LG_LIVELO.xlsx", aba_dados="Base", aba_calendario="Calendario",
        col_data_cal=4, colunas=_COLUNAS_LG,
        formula_sequencia="={letra_seq}{ant}+1", formula_chave='=CONCATENATE({letra_seq}{nova},"_","LG_LIVELO")',
        calendario=_calendario_livelo_lg,
    ),
}


# ======================================================================================
# testes offline (um por chave, --teste roda as 2) -- so a mecanica pura: montar_valores_linha
# e a funcao de calendario. Sem abrir xlsx nenhum.
# ======================================================================================
def _teste_livelo_center():
    campos = dict(data_pedido="2026-08-01", pedido="L01", entrega="E01", ordem="O01", qtdd="1",
                  vtotal=100.0, vprod=90.0, frete=10.0, comissao=-15.0, vliq=75.0, data_repasse="2026-09-30")
    valores = montar_valores_linha(_COLUNAS_CENTER, campos)
    assert valores[0] == (1, _parse_data("2026-09-30"), "mm-dd-yy")
    assert valores[1] == (2, "2026-08-01", None)
    assert valores[5] == (6, None, None) and valores[6] == (7, None, None)
    assert valores[-1] == (12, 75.0, None)

    resumo = dict(ir=10.5, campanhas=-3.2, data_repasse="2026-09-30")
    linhas = _calendario_livelo_center(resumo, 38, [1, "210-1", 3416])
    d = {c: (v, f) for c, v, f in linhas}
    assert d[1] == (1, None) and d[2] == ("210-1", None) and d[3] == (3416, None)
    assert d[8] == ("=SUMIFS(EMPILHADO!L:L,EMPILHADO!A:A,CALENDARIO!F38)", "#,##0.00")
    assert d[10] == (10.5, "#,##0.00") and d[12][0] == "Faturamento Campanhas+ IR"
    print("empilhar_repasses.py [LIVELO_CENTER]: ok")


def _teste_livelo_lg():
    resumo = dict(ir=5.0, campanhas=0.0, valor_repasse=1234.56, data_repasse="2026-09-30")
    linhas = _calendario_livelo_lg(resumo, 26, [])
    d = {c: (v, f) for c, v, f in linhas}
    assert d[2] == ("LIVELO", None) and d[3] == ("LG", None)
    assert d[6] == (1234.56, "#,##0.00")
    assert d[7] == ("=E26-SUM(F26,H26,I26)", "#,##0.00")

    campos = dict(data_pedido="2026-08-01", pedido="L01", entrega="E01", ordem="O01", qtdd="1",
                  vtotal=100.0, vprod=90.0, frete=10.0, comissao=-15.0, vliq=75.0, data_repasse="2026-09-30")
    valores = montar_valores_linha(_COLUNAS_LG, campos)
    assert valores[5] == (6, "1", "mm-dd-yy")       # qtdd preservado (so no LG)
    assert valores[-1] == (12, "2026-09-30", None)   # segunda data, como texto
    print("empilhar_repasses.py [LIVELO_LG]: ok")


TESTES = {"LIVELO_CENTER": _teste_livelo_center, "LIVELO_LG": _teste_livelo_lg}


def teste(chave):
    TESTES[chave]()


if __name__ == "__main__":
    if "--teste" in sys.argv:
        for _chave in TESTES:
            teste(_chave)
    else:
        sys.exit("uso: python empilhar_livelo_center.py | empilhar_livelo_lg.py  (ou --teste)")
