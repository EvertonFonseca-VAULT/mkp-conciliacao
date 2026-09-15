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
from calendario_alvo import buscar_alvo
from canais import CAMINHO_RB_PADRAO, CANAIS, resolver_para_processar, validar_data_rb
from cruzamento import cruzar_com_rb
from descricoes_ml import gerar_descricoes_ml
from enriquecer import enriquecer
from parser_rb import carregar_rb, metadata_rb
from rebate_cmm import gerar_rebate_cmm


def rodar(caminho_para_processar, caminho_rb):
    para_processar = pd.read_excel(caminho_para_processar, sheet_name="Para Processar")
    rb = carregar_rb(caminho_rb)
    agregado = agregar_por_pedido(para_processar)
    return cruzar_com_rb(agregado, rb, para_processar), rb, para_processar


def processar(canal, data, pasta_saida=".", usar_scraper=False):
    """Roda o ciclo inteiro (agregar/cruzar/enriquecer/rebate) e salva os
    xlsx. Usado tanto pelo `main()` (CLI) quanto pelo interativo.py
    (dashboard) -- unica funcao com a logica de ponta a ponta."""
    caminho_pp = resolver_para_processar(canal, data)
    caminho_rb = CAMINHO_RB_PADRAO
    raiz = CANAIS[canal]["raiz"]

    data_ref, responsavel = metadata_rb(caminho_rb)
    validar_data_rb(data, data_ref)

    resultado, rb, para_processar = rodar(caminho_pp, caminho_rb)
    verificar = resultado[resultado["STATUS"] == "VERIFICAR"]

    ml_dados = None
    if usar_scraper and len(verificar):
        from scraper_ml import resumir_pedidos  # so importa playwright se pedido

        print(f"Buscando {len(verificar)} pedido(s) VERIFICAR no ML/MP...")
        ml_dados = resumir_pedidos(list(verificar["PEDIDO_MKP"]))

    resultado = enriquecer(resultado, raiz, ml_dados)
    fila_verificar = resultado[resultado["STATUS"] == "VERIFICAR"].sort_values(
        "DIF", ascending=False
    )
    rebate_cmm = gerar_rebate_cmm(resultado, raiz, data_ref, responsavel, rb)
    alvo = buscar_alvo(canal, data_ref)

    pasta_conciliacoes = os.path.join(pasta_saida, "Conciliacoes")
    pasta_cmm = os.path.join(pasta_saida, "CMM")
    os.makedirs(pasta_conciliacoes, exist_ok=True)
    os.makedirs(pasta_cmm, exist_ok=True)
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    caminho_xlsx = os.path.join(pasta_conciliacoes, f"conciliacao_{sufixo}.xlsx")
    caminho_rebate_xlsx = os.path.join(pasta_cmm, f"rebate_cmm_{sufixo}.xlsx")

    descricoes_ml = gerar_descricoes_ml(ml_dados)
    caminho_descricoes_xlsx = (
        os.path.join(pasta_conciliacoes, f"descricoes_ml_{sufixo}.xlsx")
        if len(descricoes_ml) else None
    )

    # AB (Para Processar) importado junto -- so pra nao precisar abrir o
    # arquivo original a parte enquanto o time se acostuma com essa planilha.
    # Remover quando nao for mais necessario.
    try:
        with pd.ExcelWriter(caminho_xlsx) as writer:
            resultado.to_excel(writer, sheet_name="Conciliacao", index=False)
            fila_verificar.to_excel(writer, sheet_name="Fila VERIFICAR", index=False)
            rebate_cmm.to_excel(writer, sheet_name="Rebate CMM", index=False)
            para_processar.to_excel(writer, sheet_name="AB", index=False)

        with pd.ExcelWriter(caminho_rebate_xlsx) as writer:
            rebate_cmm.to_excel(writer, sheet_name="Rebate CMM", index=False)

        if caminho_descricoes_xlsx:
            descricoes_ml.to_excel(caminho_descricoes_xlsx, index=False)
    except PermissionError as exc:
        raise PermissionError(
            f"Feche o arquivo no Excel antes de rodar de novo: {exc.filename}"
        ) from exc

    return {
        "resultado": resultado,
        "fila_verificar": fila_verificar,
        "rebate_cmm": rebate_cmm,
        "alvo": alvo,
        "ml_atualizado": ml_dados is not None,
        "rb": rb,
        "para_processar": para_processar,
        "data_ref": data_ref,
        "responsavel": responsavel,
        "raiz": raiz,
        "caminho_pp": caminho_pp,
        "caminho_rb": caminho_rb,
        "caminho_xlsx": caminho_xlsx,
        "caminho_rebate_xlsx": caminho_rebate_xlsx,
        "caminho_descricoes_xlsx": caminho_descricoes_xlsx,
    }


def main():
    usar_scraper = "--scraper" in sys.argv
    args = [a for a in sys.argv[1:] if a != "--scraper"]

    canal = args[0].upper()
    if canal not in CANAIS:
        sys.exit(f"Canal '{canal}' desconhecido. Opcoes: {', '.join(CANAIS)}")
    data = args[1]

    try:
        r = processar(canal, data, usar_scraper=usar_scraper)
    except ValueError as exc:
        sys.exit(str(exc))

    resultado, fila_verificar, rebate_cmm = r["resultado"], r["fila_verificar"], r["rebate_cmm"]

    print(f"Para Processar: {r['caminho_pp']}")
    print(f"RB: {r['caminho_rb']}")
    print(f"Total de pedidos: {len(resultado)}")
    for status, qtd in resultado["STATUS"].value_counts().items():
        print(f"  {status}: {qtd}")
    print(f"Fila VERIFICAR: {len(fila_verificar)}")
    print(f"Rebate CMM (linhas pra colar no Movimento_Bancario): {len(rebate_cmm)}")
    print(f"Salvo em: {os.path.abspath(r['caminho_xlsx'])}")


if __name__ == "__main__":
    main()
