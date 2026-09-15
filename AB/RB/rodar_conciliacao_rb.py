"""Roda a conciliacao com o RB: agrega o "Para Processar" de um canal/dia
por pedido, cruza com o extrato bancario, classifica o STATUS de cada
pedido (SEM DIFF / EST COM / REBATE / MEDIACAO / VERIFICAR) e -- pros
VERIFICAR -- busca no ML/MP os valores reais e sugere a acao + historico
de baixa. Substitui o ciclo manual AB_MERCADO_TOTVS.xlsm + CONTROLE_V2.xlsx
+ abrir ML pedido a pedido.

Uso (motor de classificacao, sem scraper -- roda com qualquer Python 3
com pandas/openpyxl):
    python rodar_conciliacao_rb.py <CANAL> <DATA>
    CANAL: MOVEIS | FULL | ES | ELETRO
    DATA: dd.mm ou dd.mm.aaaa, igual ao nome do arquivo "Para Processar" do dia

Uso com scraper do ML/MP (precisa do Chrome aberto com debug remoto, ja
logado -- ver Manuais/Documentação Screaper MP e ML.md -- e rodar com o
Python do venv do Screaper Mercado, que tem playwright):
    "...\\Screaper Mercado\\.venv\\Scripts\\python.exe" rodar_conciliacao_rb.py <CANAL> <DATA> --scraper
"""

import os
import sys

import pandas as pd

from agregador import agregar_por_pedido
from canais import CAMINHO_RB_PADRAO, CANAIS, resolver_para_processar, validar_data_rb
from cruzamento import cruzar_com_rb
from enriquecer import enriquecer
from parser_rb import carregar_rb, metadata_rb
from rebate_cmm import gerar_rebate_cmm


def rodar(caminho_para_processar, caminho_rb):
    para_processar = pd.read_excel(caminho_para_processar, sheet_name="Para Processar")
    rb = carregar_rb(caminho_rb)
    agregado = agregar_por_pedido(para_processar)
    return cruzar_com_rb(agregado, rb, para_processar), rb, para_processar


def main():
    usar_scraper = "--scraper" in sys.argv
    args = [a for a in sys.argv[1:] if a != "--scraper"]

    canal = args[0].upper()
    if canal not in CANAIS:
        sys.exit(f"Canal '{canal}' desconhecido. Opcoes: {', '.join(CANAIS)}")
    data = args[1]

    caminho_pp = resolver_para_processar(canal, data)
    caminho_rb = CAMINHO_RB_PADRAO
    raiz = CANAIS[canal]["raiz"]
    saida = f"conciliacao_{canal.lower()}_{data.replace('.', '')}.xlsx"

    data_ref, responsavel = metadata_rb(caminho_rb)
    try:
        validar_data_rb(data, data_ref)
    except ValueError as exc:
        sys.exit(str(exc))

    resultado, rb, para_processar = rodar(caminho_pp, caminho_rb)
    verificar = resultado[resultado["STATUS"] == "VERIFICAR"]

    ml_dados = None
    if usar_scraper and len(verificar):
        from scraper_ml import resumir_pedidos  # so importa playwright se pedido

        print(f"Buscando {len(verificar)} pedido(s) VERIFICAR no ML/MP...")
        ml_dados = resumir_pedidos(list(verificar["PEDIDO_MKP"]))

    resultado = enriquecer(resultado, raiz, ml_dados)
    fila_verificar = resultado[resultado["STATUS"] == "VERIFICAR"].sort_values(
        "DIF", key=abs, ascending=False
    )
    rebate_cmm = gerar_rebate_cmm(resultado, raiz, data_ref, responsavel, rb)

    # AB (Para Processar) importado junto -- so pra nao precisar abrir o
    # arquivo original a parte enquanto o time se acostuma com essa planilha.
    # Remover quando nao for mais necessario.
    with pd.ExcelWriter(saida) as writer:
        resultado.to_excel(writer, sheet_name="Conciliacao", index=False)
        fila_verificar.to_excel(writer, sheet_name="Fila VERIFICAR", index=False)
        rebate_cmm.to_excel(writer, sheet_name="Rebate CMM", index=False)
        para_processar.to_excel(writer, sheet_name="AB", index=False)

    print(f"Para Processar: {caminho_pp}")
    print(f"RB: {caminho_rb}")
    print(f"Total de pedidos: {len(resultado)}")
    for status, qtd in resultado["STATUS"].value_counts().items():
        print(f"  {status}: {qtd}")
    print(f"Fila VERIFICAR: {len(fila_verificar)}")
    print(f"Rebate CMM (linhas pra colar no Movimento_Bancario): {len(rebate_cmm)}")
    print(f"Salvo em: {os.path.abspath(saida)}")


if __name__ == "__main__":
    main()
