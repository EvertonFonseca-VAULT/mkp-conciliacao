"""Para Processar dos canais FULL, ES, MOVEIS, ELETRO, CENTER, CO05 e PRIME, num comando so.
Mesmo processo do rodar_es.py -- muda so o que cada canal usa:

  FULL    TOTVS/SRI + PTI (750+751)    -> PROCESSAR\\FULL\\PROCESSAR_MERCADO FULL dd.mm.aaaa.xlsx (igual rodar_full.py)
  ES      TOTVS/SRI + PTI + frete      -> PROCESSAR\\COLETAS ES\\PROCESSAR_MERCADO COLETAS ES dd.mm.xlsx
  MOVEIS  TOTVS/SRI + PTI, SEM frete   -> PROCESSAR\\MOVEIS\\PROCESSAR_MERCADO MOVEIS dd.mm.aaaa.xlsx
  ELETRO  so TOTVS/SRI (sem PTI/frete) -> ELETRODOMESTICO\\MERCADO ELETRODOMESTICO dd.mm.xlsx
  CENTER  TOTVS/SRI + PTI, SEM frete   -> PROCESSAR\\CENTER\\PROCESSAR_MERCADO CENTER dd.mm.aaaa.xlsx (vendedor 944)
  CO05    TOTVS/SRI + PTI + frete      -> PROCESSAR\\COLETAS CO05\\PROCESSAR_MERCADO COLETAS CO05 dd.mm.xlsx (vendedor 847,
          mesmo layout de 60 colunas do ES, so em ordem diferente -- processar_es.carregar_ab ja reordena sozinho)
  PRIME   TOTVS/SRI + PTI, SEM frete   -> PROCESSAR\\PRIME\\PROCESSAR_MERCADO PRIME dd.mm.aaaa.xlsx (vendedor 657,
          adicionado 23/09/2026 -- frete=False confirmado contra o calendario, ver validacao abaixo)

Validado em 19-20/09/2026: MOVEIS 04.09 e ELETRO 04.09 contra os arquivos manuais,
ES 01-10.09 (ver rodar_es.py) e H:M de todos os dias contra o calendario/Koncili.
PRIME validado em 23/09/2026, 5 dias (11, 14, 16, 17 e 18.09.2026): H:M do
Para Processar = RESULTADO do Koncili = VALOR_EB do calendario, os 3 batendo
exato nos 5 dias (confirma de quebra que frete=False estava certo -- se
estivesse errado o H:M teria vindo diferente dos outros dois).

Uso:
    python rodar_canal.py FULL 22.09.2026                    # 1 canal, 1 dia unico
    python rodar_canal.py FULL 11.09.2026 18.09.2026         # 1 canal, varios dias -- lista cada data
                                                               # (NAO e intervalo: dias que faltam no meio
                                                               # ficam de fora, precisa listar um a um)
    python rodar_canal.py TODOS 22.09.2026                   # FULL, ES, MOVEIS, ELETRO e CENTER do dia
    python rodar_canal.py ES 11.09.2026 --saida C:\\temp      # grava em outra pasta
    python rodar_canal.py ELETRO 11.09.2026 --forcar          # sobrescreve
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "ES"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "MOVEIS"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "FULL"))

import pandas as pd

from rodar_es import rodar
from rodar_full import rodar as rodar_full

PASTA_AB = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB"
RODAR_TOTVS = r"C:\RODAR TOTVS"

# ab_sub: subpasta do AB (ES: "antigo" = layout de 60 colunas); ano_no_nome:
# a saida de MOVEIS leva o ano no nome (dd.mm.aaaa), as outras so dd.mm.
CANAIS_PP = {
    "FULL": dict(ab_pasta="FULL", ab_prefixo="MERCADO FULL", ab_sub="", pti=r"C:\TEMP\PTI FULL*.csv", frete=True,
                 saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "FULL"), saida_prefixo="PROCESSAR_MERCADO FULL", ano_no_nome=True),
    "ES": dict(ab_pasta="COLETAS ES", ab_prefixo="MERCADO COLETAS ES", ab_sub="antigo", pti=r"C:\TEMP\PTI ES.csv", frete=True,
               saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "COLETAS ES"), saida_prefixo="PROCESSAR_MERCADO COLETAS ES", ano_no_nome=False),
    "MOVEIS": dict(ab_pasta="MOVEIS", ab_prefixo="MERCADO MOVEIS", ab_sub="", pti=r"C:\TEMP\PTI MOVEIS.csv", frete=False,
                   saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "MOVEIS"), saida_prefixo="PROCESSAR_MERCADO MOVEIS", ano_no_nome=True),
    "ELETRO": dict(ab_pasta="ELETRODOMESTICO", ab_prefixo="MERCADO ELETRODOMESTICO", ab_sub="", pti=None, frete=False,
                   saida=os.path.join(RODAR_TOTVS, "ELETRODOMESTICO"), saida_prefixo="MERCADO ELETRODOMESTICO", ano_no_nome=False),
    "CENTER": dict(ab_pasta="CENTER", ab_prefixo="MERCADO CENTER", ab_sub="", pti=r"C:\TEMP\PTI CENTER.csv", frete=False,
                   saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "CENTER"), saida_prefixo="PROCESSAR_MERCADO CENTER", ano_no_nome=True),
    "CO05": dict(ab_pasta="COLETAS CO05", ab_prefixo="MERCADO COLETAS CO05", ab_sub="", pti=r"C:\TEMP\PTI COLETAS 05.csv", frete=True,
                 saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "COLETAS CO05"), saida_prefixo="PROCESSAR_MERCADO COLETAS CO05", ano_no_nome=False),
    # PRIME: PTI chegou 23/09/2026 (vendedor 657) -- frete=False (loja propria,
    # nao e canal de coleta), confirmado com dado real no mesmo dia: H:M = Koncili
    # = calendario nos 5 dias testados (11, 14, 16, 17 e 18.09.2026).
    "PRIME": dict(ab_pasta="PRIME", ab_prefixo="MERCADO PRIME", ab_sub="", pti=r"C:\TEMP\PTI_PRIME.csv", frete=False,
                  saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "PRIME"), saida_prefixo="PROCESSAR_MERCADO PRIME", ano_no_nome=True),
}


def caminhos(canal, data):
    """`data` = DD.MM.AAAA -> (AB de entrada, nome do arquivo de saida)."""
    c = CANAIS_PP[canal]
    dia, mes, ano = data.split(".")
    partes = [PASTA_AB, c["ab_pasta"], ano, f"{mes}.{ano}"] + ([c["ab_sub"]] if c["ab_sub"] else [])
    ab = os.path.join(*partes, f"{c['ab_prefixo']} {data}.xlsx")
    nome = f"{c['saida_prefixo']} {data if c['ano_no_nome'] else f'{dia}.{mes}'}.xlsx"
    return ab, nome


def processar(canal, data, pasta_saida=None, forcar=False):
    c = CANAIS_PP[canal]
    ab, nome = caminhos(canal, data)
    destino = os.path.join(pasta_saida or c["saida"], nome)
    if os.path.exists(destino) and not forcar:
        print(f"  {canal} {data}: ja existe, nao sobrescrevi ({destino}) -- use --forcar")
        return None
    if canal == "FULL":  # pipeline proprio (2 vendedores no PTI, modelo_ab_full)
        validacao, _pag, _fr, pp = rodar_full(ab, c["pti"])
    else:
        validacao, _pag, _fr, pp = rodar(ab, c["pti"], frete=c["frete"])
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    with pd.ExcelWriter(destino) as w:
        pp.to_excel(w, sheet_name="Para Processar", index=False)
    hm = pp.iloc[:, 7:13].apply(pd.to_numeric, errors="coerce").sum().sum()
    extra = "" if validacao is None else f" | Verificar {int((validacao.categoria == 'Verificar').sum())} de {len(validacao)} pedidos"
    print(f"  {canal} {data}: {len(pp)} linhas | H:M {hm:,.2f}{extra} -> {destino}")
    return pp


def main(argv):
    forcar = "--forcar" in argv
    pasta = argv[argv.index("--saida") + 1] if "--saida" in argv else None
    args = [a for i, a in enumerate(argv) if not a.startswith("--") and not (i and argv[i - 1] == "--saida")]
    if len(args) < 2 or args[0].upper() not in [*CANAIS_PP, "TODOS"]:
        sys.exit(f"uso: python rodar_canal.py <{'|'.join(CANAIS_PP)}|TODOS> DD.MM.AAAA [DD.MM.AAAA ...] [--saida PASTA] [--forcar]")
    canais = list(CANAIS_PP) if args[0].upper() == "TODOS" else [args[0].upper()]
    for data in args[1:]:
        for canal in canais:
            try:
                processar(canal, data, pasta, forcar)
            except FileNotFoundError as e:  # AB do canal ainda nao esta no X: -- segue com os outros
                print(f"  {canal} {data}: sem AB ({e.filename or e})")


def _teste():
    """Offline: caminhos e nomes de saida no padrao dos arquivos existentes."""
    ab, nome = caminhos("MOVEIS", "11.09.2026")
    assert ab.endswith(r"MOVEIS\2026\09.2026\MERCADO MOVEIS 11.09.2026.xlsx") and nome == "PROCESSAR_MERCADO MOVEIS 11.09.2026.xlsx"
    ab, nome = caminhos("ES", "18.09.2026")
    assert ab.endswith(r"COLETAS ES\2026\09.2026\antigo\MERCADO COLETAS ES 18.09.2026.xlsx") and nome == "PROCESSAR_MERCADO COLETAS ES 18.09.xlsx"
    ab, nome = caminhos("FULL", "11.09.2026")
    assert ab.endswith(r"FULL\2026\09.2026\MERCADO FULL 11.09.2026.xlsx") and nome == "PROCESSAR_MERCADO FULL 11.09.2026.xlsx"
    assert caminhos("ELETRO", "04.09.2026")[1] == "MERCADO ELETRODOMESTICO 04.09.xlsx"
    ab, nome = caminhos("CENTER", "22.09.2026")
    assert ab.endswith(r"CENTER\2026\09.2026\MERCADO CENTER 22.09.2026.xlsx") and nome == "PROCESSAR_MERCADO CENTER 22.09.2026.xlsx"
    ab, nome = caminhos("CO05", "18.09.2026")
    assert ab.endswith(r"COLETAS CO05\2026\09.2026\MERCADO COLETAS CO05 18.09.2026.xlsx") and nome == "PROCESSAR_MERCADO COLETAS CO05 18.09.xlsx"
    print("rodar_canal.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
