"""Empilhado KABUM (gohorse -- regra provisoria, estruturar depois, ver conversa 2026-10-13).

Fonte: X:\\...\\MKP - KABUM\\AB\\<CANAL>\\<ano>\\<mes>.<ano>\\DD.MM.AAAA KABUM[ CE].xlsx
  aba "Sheet1". So pega arquivos "DD.MM.AAAA KABUM*.xlsx" -- ignora a familia
  "KABUM_REPASSE_*" (csv/xlsx), que e outro layout (schema antigo, outro destino).
  Cada arquivo = um dia so (confirmado: "Data do repasse" nao varia dentro do arquivo).
  Colunas por NOME (nao posicao) -- a contagem de coluna varia (19 ou 20, uma coluna em
  branco extra em alguns arquivos).

Destino: KABUM_NOVO_CENTER_LAYOUT_NOVO.xlsx, aba KABUM_CENTER (confirmado com o usuario --
  NAO e o KABUM_NOVO_CENTER.xlsx antigo, schema diferente, parado em maio/2026).
  Layout = Data_Repasse (= Data do repasse, convertida) + colunas da fonte (ate Numero do
  pedido) + Pedido TOTVS (= Numero do pedido + "-A", validado contra as linhas existentes,
  Transacao e Estorno) + resto das colunas da fonte (sem "Observacoes", que o destino nao tem).

Dedupe: por chave natural (TID, Parcela atual, Valor do repasse financeiro da parcela) --
  testada contra as linhas ja empilhadas, 0 colisao (TID+Parcela sozinho colide 9x).

Formatacao: "Data_Repasse", "Data do repasse" e "Data do pedido" sao datetime na planilha
  original (nao texto -- conferido direto na ultima linha autentica), formato "mm-dd-yy".
  NAO usa ws.append() -- a planilha tem linhas em branco ja formatadas logo apos a ultima
  linha real (sobra de uma tentativa anterior), e append() ia escrever em cima delas sem
  corrigir o formato. Acha a primeira linha realmente vazia (coluna B = ID do repasse) e
  escreve celula a celula, copiando number_format da ultima linha autentica.

Calendario: aba EMPILHADO, uma linha por dia novo (BANCO/CONTA/AGENCIA/MARKETPLACE/LOJA
  copiados da ultima linha existente, DATA = o dia). VALOR_EB/VALOR_AB/DIF sao FORMULA
  (=SUMIFS(KABUM_CENTER!T:T,KABUM_CENTER!A:A,EMPILHADO!F<linha>) nos dois, =G-H no DIF) --
  conferido direto na ultima linha autentica, NAO sao valor estatico (achado so depois de
  ler com data_only=False; com data_only=True parecia numero fixo e nao era).

Uso: python empilhar_kabum.py CENTER
"""
import datetime
import os
import re
import shutil
import sys
from copy import copy

import openpyxl

PASTA_FONTE_BASE = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - KABUM\AB"
PASTA_DESTINO = (r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\ARQUIVOS EMPILHADOS DE BAIXAS"
                 r"\ARQUIVOS REPASSES COM CHAVES PRIMÁRIAS")
PASTA_BACKUP = r"C:\BACKUP EMPILHADO"

_PADRAO_ARQUIVO = re.compile(r"^\d{2}\.\d{2}\.\d{4} KABUM", re.I)

CANAIS = {
    "CENTER": dict(arquivo="KABUM_NOVO_CENTER_LAYOUT_NOVO.xlsx", aba_destino="KABUM_CENTER",
                   aba_calendario="EMPILHADO"),
}

# ordem das colunas do destino, e de onde cada uma vem ("src:<nome na fonte>" ou "calc")
CAMPOS_DESTINO = [
    ("Data_Repasse", "calc_data_repasse"),
    ("ID do repasse", "src:ID do repasse"),
    ("Origem", "src:Origem"),
    ("ID da transação", "src:ID da transação"),
    ("TID", "src:TID"),
    ("Autorização", "src:Autorização"),
    ("NSU", "src:NSU"),
    ("Data do repasse", "calc_data_repasse"),
    ("Número do pedido", "src:Número do pedido"),
    ("Pedido TOTVS", "calc_pedido_totvs"),
    ("Data do pedido", "calc_data_pedido"),
    ("Forma de pagamento", "src:Forma de pagamento"),
    ("Tipo de operação", "src:Tipo de operação"),
    ("Total de parcelas", "src:Total de parcelas"),
    ("Parcela atual", "src:Parcela atual"),
    ("Modelo de repasse da parcela", "src:Modelo de repasse da parcela"),
    ("% de taxa de Antecipação por parcela", "src:% de taxa de Antecipação por parcela"),
    ("Custo da Antecipação por parcela (R$)", "src:Custo da Antecipação por parcela (R$)"),
    ("Taxa de transferência (R$)", "src:Taxa de transferência (R$)"),
    ("Valor do repasse financeiro da parcela (R$)", "src:Valor do repasse financeiro da parcela (R$)"),
]


def _mapa_colunas(header):
    return {nome: i for i, nome in enumerate(header) if nome}


def _linhas_fonte(canal):
    pasta = os.path.join(PASTA_FONTE_BASE, canal)
    achados = []
    for raiz, _dirs, arquivos in os.walk(pasta):
        for nome in arquivos:
            if _PADRAO_ARQUIVO.match(nome):
                achados.append(os.path.join(raiz, nome))

    candidatas = []
    for caminho in achados:
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        if "Sheet1" not in wb.sheetnames:
            print(f"  ignorado (sem aba Sheet1, schema diferente): {os.path.basename(caminho)} {wb.sheetnames}")
            wb.close()
            continue
        ws = wb["Sheet1"]
        linhas = list(ws.iter_rows(values_only=True))
        wb.close()
        header, *dados = linhas
        col = _mapa_colunas(header)
        for row in dados:
            if row[col["ID do repasse"]] is None:
                continue
            candidatas.append({nome: row[i] for nome, i in col.items()})
    return candidatas, achados


def _data_br(txt):
    return datetime.datetime.strptime(txt, "%d/%m/%Y") if isinstance(txt, str) else txt


def _valor_linha(nome_dest, origem, linha):
    if origem == "calc_data_repasse":
        return _data_br(linha["Data do repasse"])
    if origem == "calc_data_pedido":
        return _data_br(linha["Data do pedido"])
    if origem == "calc_pedido_totvs":
        return f"{linha['Número do pedido']}-A"
    return linha[origem.split(":", 1)[1]]


def _chave_natural(linha):
    return (linha["TID"], linha["Parcela atual"], linha["Valor do repasse financeiro da parcela (R$)"])


def _primeira_linha_vazia(ws, col_chave):
    """Ultima linha com dado em `col_chave` (1-based) + 1 -- NAO usa ws.max_row (que conta
    linhas em branco ja formatadas que sobraram de uma gravacao anterior)."""
    ultima = 1  # header
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=col_chave).value is not None:
            ultima = r
    return ultima + 1


def _copia_estilo(origem_cell, destino_cell):
    destino_cell.number_format = origem_cell.number_format
    destino_cell.font = copy(origem_cell.font)
    destino_cell.alignment = copy(origem_cell.alignment)
    destino_cell.fill = copy(origem_cell.fill)
    destino_cell.border = copy(origem_cell.border)


def _grava_dados(ws, novas, col_dest):
    linha_modelo = _primeira_linha_vazia(ws, col_dest["ID do repasse"]) - 1
    proxima = linha_modelo + 1
    for i, linha in enumerate(novas):
        r = proxima + i
        for c, (nome, origem) in enumerate(CAMPOS_DESTINO, start=1):
            cel = ws.cell(row=r, column=c, value=_valor_linha(nome, origem, linha))
            _copia_estilo(ws.cell(row=linha_modelo, column=c), cel)


def _grava_calendario(wb, cfg, novas):
    ws = wb[cfg["aba_calendario"]]
    header = [c.value for c in ws[1]]
    col = _mapa_colunas(header)
    linha_modelo = _primeira_linha_vazia(ws, col["BANCO"] + 1) - 1
    modelo = [c.value for c in ws[linha_modelo]]

    por_dia = {}
    for linha in novas:
        dia = _data_br(linha["Data do repasse"])
        por_dia[dia] = por_dia.get(dia, 0) + (linha["Valor do repasse financeiro da parcela (R$)"] or 0)

    proxima = linha_modelo + 1
    for i, dia in enumerate(sorted(por_dia)):
        r = proxima + i
        formula = f"=SUMIFS(KABUM_CENTER!T:T,KABUM_CENTER!A:A,EMPILHADO!F{r})"
        valores = [modelo[col["BANCO"]], modelo[col["CONTA"]], modelo[col["AGÊNCIA"]],
                   modelo[col["MARKETPLACE"]], modelo[col["LOJA"]], dia, formula, formula, f"=G{r}-H{r}"]
        for c, v in enumerate(valores, start=1):
            cel = ws.cell(row=r, column=c, value=v)
            _copia_estilo(ws.cell(row=linha_modelo, column=c), cel)
    print(f"  calendario: {len(por_dia)} dia(s) novo(s) ({', '.join(d.strftime('%d/%m') for d in sorted(por_dia))})")


def empilhar(canal):
    cfg = CANAIS[canal]
    destino = os.path.join(PASTA_DESTINO, cfg["arquivo"])

    candidatas, arquivos_fonte = _linhas_fonte(canal)

    wb = openpyxl.load_workbook(destino)
    ws = wb[cfg["aba_destino"]]
    header_dest = [c.value for c in ws[1]]
    col_dest = _mapa_colunas(header_dest)

    existentes = set()
    for r in range(2, _primeira_linha_vazia(ws, col_dest["TID"] + 1)):
        existentes.add((ws.cell(row=r, column=col_dest["TID"] + 1).value,
                         ws.cell(row=r, column=col_dest["Parcela atual"] + 1).value,
                         ws.cell(row=r, column=col_dest["Valor do repasse financeiro da parcela (R$)"] + 1).value))

    novas = [l for l in candidatas if _chave_natural(l) not in existentes]

    print(f"{canal}: {len(arquivos_fonte)} arquivo(s) fonte | candidatas={len(candidatas)} | novas={len(novas)}")
    if not novas:
        return

    os.makedirs(PASTA_BACKUP, exist_ok=True)
    backup = os.path.join(PASTA_BACKUP, f"{cfg['arquivo']}.bak_{datetime.datetime.now():%Y%m%d_%H%M%S}")
    shutil.copy2(destino, backup)
    print(f"  backup: {backup}")

    _grava_dados(ws, novas, col_dest)
    _grava_calendario(wb, cfg, novas)

    wb.save(destino)
    print(f"  gravado: {len(novas)} linhas novas")


def _teste():
    """So a mecanica pura (mapa de colunas, calculo de campo, chave de
    dedupe) -- sem abrir xlsx nenhum."""
    header = ["ID do repasse", None, "TID", "Parcela atual"]
    assert _mapa_colunas(header) == {"ID do repasse": 0, "TID": 2, "Parcela atual": 3}

    assert _data_br("10/10/2026") == datetime.datetime(2026, 10, 10)
    assert _data_br(datetime.datetime(2026, 10, 10)) == datetime.datetime(2026, 10, 10)

    linha = {
        "Data do repasse": "10/10/2026", "Data do pedido": "09/10/2026",
        "Número do pedido": "12345", "TID": "T1", "Parcela atual": 2,
        "Valor do repasse financeiro da parcela (R$)": 50.0,
    }
    assert _valor_linha("Data_Repasse", "calc_data_repasse", linha) == datetime.datetime(2026, 10, 10)
    assert _valor_linha("Data do pedido", "calc_data_pedido", linha) == datetime.datetime(2026, 10, 9)
    assert _valor_linha("Pedido TOTVS", "calc_pedido_totvs", linha) == "12345-A"
    assert _valor_linha("TID", "src:TID", linha) == "T1"

    assert _chave_natural(linha) == ("T1", 2, 50.0)
    print("empilhar_kabum.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    elif len(sys.argv) != 2 or sys.argv[1] not in CANAIS:
        sys.exit(f"uso: python empilhar_kabum.py <{'|'.join(CANAIS)}>  |  --teste")
    else:
        empilhar(sys.argv[1])
