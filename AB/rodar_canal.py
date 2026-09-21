"""Para Processar dos canais ES, MOVEIS e ELETRO (o FULL tem o rodar_full.py).
Mesmo processo do rodar_es.py -- muda so o que cada canal usa:

  ES      TOTVS/SRI + PTI + frete      -> PROCESSAR\\COLETAS ES\\PROCESSAR_MERCADO COLETAS ES dd.mm.xlsx
  MOVEIS  TOTVS/SRI + PTI, SEM frete   -> PROCESSAR\\MOVEIS\\PROCESSAR_MERCADO MOVEIS dd.mm.aaaa.xlsx
  ELETRO  so TOTVS/SRI (sem PTI/frete) -> ELETRODOMESTICO\\MERCADO ELETRODOMESTICO dd.mm.xlsx

Validado em 19-20/09/2026: MOVEIS 04.09 e ELETRO 04.09 contra os arquivos manuais,
ES 01-10.09 (ver rodar_es.py) e H:M de todos os dias contra o calendario/Koncili.

Uso:
    python rodar_canal.py MOVEIS 11.09.2026 18.09.2026     # grava (nao sobrescreve)
    python rodar_canal.py ES 11.09.2026 --saida C:\\temp    # grava em outra pasta
    python rodar_canal.py ELETRO 11.09.2026 --forcar         # sobrescreve
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "ES"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "MOVEIS"))

import pandas as pd

from rodar_es import rodar

PASTA_AB = r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\MKP\MKP - MERCADO\AB"
RODAR_TOTVS = r"C:\RODAR TOTVS"

# ab_sub: subpasta do AB (ES: "antigo" = layout de 60 colunas); ano_no_nome:
# a saida de MOVEIS leva o ano no nome (dd.mm.aaaa), as outras so dd.mm.
CANAIS_PP = {
    "ES": dict(ab_pasta="COLETAS ES", ab_prefixo="MERCADO COLETAS ES", ab_sub="antigo", pti=r"C:\TEMP\PTI ES.csv", frete=True,
               saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "COLETAS ES"), saida_prefixo="PROCESSAR_MERCADO COLETAS ES", ano_no_nome=False),
    "MOVEIS": dict(ab_pasta="MOVEIS", ab_prefixo="MERCADO MOVEIS", ab_sub="", pti=r"C:\TEMP\PTI MOVEIS.csv", frete=False,
                   saida=os.path.join(RODAR_TOTVS, "PROCESSAR", "MOVEIS"), saida_prefixo="PROCESSAR_MERCADO MOVEIS", ano_no_nome=True),
    "ELETRO": dict(ab_pasta="ELETRODOMESTICO", ab_prefixo="MERCADO ELETRODOMESTICO", ab_sub="", pti=None, frete=False,
                   saida=os.path.join(RODAR_TOTVS, "ELETRODOMESTICO"), saida_prefixo="MERCADO ELETRODOMESTICO", ano_no_nome=False),
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
    if len(args) < 2 or args[0].upper() not in CANAIS_PP:
        sys.exit(f"uso: python rodar_canal.py <{'|'.join(CANAIS_PP)}> DD.MM.AAAA [DD.MM.AAAA ...] [--saida PASTA] [--forcar]")
    for data in args[1:]:
        processar(args[0].upper(), data, pasta, forcar)


def _teste():
    """Offline: caminhos e nomes de saida no padrao dos arquivos existentes."""
    ab, nome = caminhos("MOVEIS", "11.09.2026")
    assert ab.endswith(r"MOVEIS\2026\09.2026\MERCADO MOVEIS 11.09.2026.xlsx") and nome == "PROCESSAR_MERCADO MOVEIS 11.09.2026.xlsx"
    ab, nome = caminhos("ES", "18.09.2026")
    assert ab.endswith(r"COLETAS ES\2026\09.2026\antigo\MERCADO COLETAS ES 18.09.2026.xlsx") and nome == "PROCESSAR_MERCADO COLETAS ES 18.09.xlsx"
    assert caminhos("ELETRO", "04.09.2026")[1] == "MERCADO ELETRODOMESTICO 04.09.xlsx"
    print("rodar_canal.py: ok")


if __name__ == "__main__":
    if "--teste" in sys.argv:
        _teste()
    else:
        main(sys.argv[1:])
