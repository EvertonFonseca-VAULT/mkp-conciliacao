"""Assistente interativo pra rodar a conciliacao sem digitar comando no
terminal: pergunta canal e data, roda a mesma logica de sempre
(rodar_conciliacao_rb.processar) e abre um dashboard HTML com o resultado.

Uso do dia a dia (time): duplo-clique em iniciar_conciliacao.bat.

Uso interno (nao pergunta nada -- o proprio programa se re-invoca assim
quando precisa do scraper, trocando pro Python do venv do Screaper):
    python interativo.py --canal FULL --data 08.09 --saida-pasta X --scraper
"""

import argparse
import glob
import importlib.util
import os
import subprocess
import sys
import webbrowser

from canais import CANAIS, CANAIS_SEM_RAIZ_CONFIRMADA
from dashboard import gerar_dashboard
from rodar_conciliacao_rb import processar

# NAO importar scraper_ml aqui em cima -- ele importa playwright no topo do
# arquivo, e o Python normal (sem scraper) nao tem playwright instalado.
# CAMINHO_SCRAPER e calculado (nao mais hardcoded) -- tem que bater com o
# mesmo calculo em CAMINHO_SCRAPER de scraper_ml.py.
_HOME = os.path.expanduser("~")
CAMINHO_SCRAPER = os.path.join(_HOME, "Desktop", "Projetos", "Screaper Mercado")
PASTA_REDE_PADRAO = os.path.join(
    r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\TEMP\CONCILIACAO",
    os.path.basename(_HOME),
)
SCRAPER_PYTHON = os.path.join(CAMINHO_SCRAPER, ".venv", "Scripts", "python.exe")


def _listar_datas_disponiveis(canal):
    cfg = CANAIS[canal]
    padrao = f"{cfg['pasta']}\\{cfg['prefixo']} *.xlsx"
    prefixo_len = len(cfg["prefixo"]) + 1
    datas = {os.path.basename(f)[prefixo_len:-5] for f in glob.glob(padrao)}
    return sorted(datas, reverse=True)


def perguntar_canal():
    opcoes = list(CANAIS)
    print("\nCanais disponiveis:")
    for i, canal in enumerate(opcoes, 1):
        print(f"  {i}) {canal}")
    while True:
        escolha = input("Escolha o canal: ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(opcoes):
            return opcoes[int(escolha) - 1]
        print("Opcao invalida.")


def perguntar_data(canal):
    pasta = CANAIS[canal]["pasta"]
    print(f"\nBuscando 'Para Processar' em: {pasta}")
    datas = _listar_datas_disponiveis(canal)
    if not datas:
        sys.exit(f"Nenhum arquivo 'Para Processar' encontrado em {pasta}.")
    print("Arquivos 'Para Processar' encontrados:")
    for i, data in enumerate(datas, 1):
        marca = "  <- mais recente" if i == 1 else ""
        print(f"  {i}) {data}{marca}")
    escolha = input("Escolha a data (Enter = mais recente): ").strip()
    if not escolha:
        return datas[0]
    if escolha.isdigit() and 1 <= int(escolha) <= len(datas):
        return datas[int(escolha) - 1]
    return escolha  # deixa digitar a data direto tambem


def perguntar_sim_nao(pergunta, padrao_sim):
    sufixo = "[S/n]" if padrao_sim else "[s/N]"
    resposta = input(f"{pergunta} {sufixo}: ").strip().lower()
    return padrao_sim if not resposta else resposta == "s"


def scraper_disponivel():
    return importlib.util.find_spec("playwright") is not None


def relancar_com_scraper(canal, data, pasta_saida):
    if not os.path.exists(SCRAPER_PYTHON):
        print(f"Nao encontrei o Python do venv do Screaper em {SCRAPER_PYTHON}.")
        return 1
    cmd = [SCRAPER_PYTHON, __file__, "--canal", canal, "--data", data,
           "--saida-pasta", pasta_saida, "--scraper"]
    return subprocess.run(cmd).returncode


def gerar_bat_scraper(canal, data, pasta_saida, pasta_destino):
    """.bat de 2 cliques que reroda essa mesma conciliacao com --scraper --
    fica salvo em pasta_destino (Conciliacoes/, do lado do dashboard) mas
    reinvoca com --saida-pasta=pasta_saida (a RAIZ, nao Conciliacoes/) --
    executar()/processar() recalculam as subpastas Conciliacoes/CMM a
    partir dela; passar a subpasta aqui geraria Conciliacoes/Conciliacoes
    aninhado, escrevendo num lugar diferente do dashboard que o usuario
    esta olhando."""
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    nome_bat = f"rodar_scraper_{sufixo}.bat"
    caminho_bat = os.path.join(pasta_destino, nome_bat)
    conteudo = (
        "@echo off\n"
        f'"{SCRAPER_PYTHON}" "{os.path.abspath(__file__)}" '
        f'--canal {canal} --data {data} --saida-pasta "{pasta_saida}" --scraper\n'
        "pause\n"
    )
    with open(caminho_bat, "w", encoding="cp1252") as f:
        f.write(conteudo)
    return nome_bat


def gerar_bat_atualizar(canal, data, pasta_saida, pasta_destino):
    """.bat de 2 cliques que reroda essa conciliacao SEM scraper -- so pra
    reler rb.xlsx/rbx.xlsx (AUXILIAR/) e o RB/AB mais recentes e regerar o
    dashboard, sem precisar abrir o Chrome/scraper (rapido, so pandas).
    Pedido do Everton, 16/09/2026: botao pra atualizar a conferencia
    Original x Transitoria depois de lancar mais movimentos."""
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    nome_bat = f"atualizar_{sufixo}.bat"
    caminho_bat = os.path.join(pasta_destino, nome_bat)
    conteudo = (
        "@echo off\n"
        f'"{sys.executable}" "{os.path.abspath(__file__)}" '
        f'--canal {canal} --data {data} --saida-pasta "{pasta_saida}"\n'
        "pause\n"
    )
    with open(caminho_bat, "w", encoding="cp1252") as f:
        f.write(conteudo)
    return nome_bat


def executar(canal, data, pasta_saida, usar_scraper):
    print(f"\nProcessando {canal} {data}...")
    r = processar(canal, data, pasta_saida=pasta_saida, usar_scraper=usar_scraper)
    pasta_conciliacoes = os.path.join(pasta_saida, "Conciliacoes")
    raiz_confirmada = r["raiz"] if canal not in CANAIS_SEM_RAIZ_CONFIRMADA else None

    caminho_bat_scraper = None
    if not usar_scraper and len(r["fila_verificar"]):
        caminho_bat_scraper = gerar_bat_scraper(canal, data, pasta_saida, pasta_conciliacoes)
    caminho_bat_atualizar = gerar_bat_atualizar(canal, data, pasta_saida, pasta_conciliacoes)

    caminho_html_esperado = os.path.join(pasta_conciliacoes, f"dashboard_{data.replace('.', '')}.html")
    dashboard_ja_existia = os.path.exists(caminho_html_esperado)

    caminho_html = gerar_dashboard(
        r["resultado"], r["fila_verificar"], r["rebate_cmm"],
        canal=canal, data=data, data_ref_rb=r["data_ref"],
        caminho_xlsx_completo=os.path.basename(r["caminho_xlsx"]),
        caminho_xlsx_rebate=os.path.basename(r["caminho_rebate_xlsx"]),
        pasta_saida=pasta_conciliacoes, alvo=r["alvo"], ml_atualizado=r["ml_atualizado"],
        caminho_bat_scraper=caminho_bat_scraper, raiz=raiz_confirmada,
        caminho_xlsx_descricoes=(
            os.path.basename(r["caminho_descricoes_xlsx"]) if r["caminho_descricoes_xlsx"] else None
        ),
        conferencia_rbx=r["conferencia_rbx"], caminho_bat_atualizar=caminho_bat_atualizar,
        ml_de_cache=r["ml_de_cache"],
    )

    print(f"Total de pedidos: {len(r['resultado'])}")
    for status, qtd in r["resultado"]["STATUS"].value_counts().items():
        print(f"  {status}: {qtd}")
    print(f"Dashboard do dia ({canal}): {caminho_html}")

    # So abre aba nova na 1a vez do dia -- da 2a em diante a aba que ja
    # esta aberta se recarrega sozinha (visibilitychange, ver dashboard.py)
    # quando o Everton volta pra ela, sem precisar de outro link.
    # ponytail: deteccao e por existencia do arquivo, nao da aba em si -- se
    # o Everton fechar a aba, nao abre outra sozinho ate o dia seguinte
    # (teria que apagar o dashboard_<data>.html pra forcar reabrir).
    if dashboard_ja_existia:
        print("Ja tem uma aba aberta com esse dashboard -- volta pra ela, atualiza sozinha.")
    else:
        webbrowser.open(caminho_html)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--canal")
    parser.add_argument("--data")
    parser.add_argument("--saida-pasta", default=PASTA_REDE_PADRAO)
    parser.add_argument("--scraper", action="store_true")
    args = parser.parse_args()

    if args.canal and args.data:
        # re-execucao interna (com o Python certo pro scraper) -- sem perguntas
        try:
            executar(args.canal.upper(), args.data, args.saida_pasta, args.scraper)
        except (ValueError, FileNotFoundError, PermissionError) as exc:
            sys.exit(str(exc))
        return

    print("=== Conciliacao Mercado Livre / TOTVS ===")
    while True:
        canal = perguntar_canal()
        data = perguntar_data(canal)
        usar_scraper = perguntar_sim_nao(
            "\nRodar busca no ML/MP pros pedidos VERIFICAR? (precisa do Chrome aberto e logado)",
            padrao_sim=False,
        )

        try:
            if usar_scraper and not scraper_disponivel():
                if relancar_com_scraper(canal, data, args.saida_pasta) != 0:
                    print("Falhou ao rodar com o scraper -- veja o erro acima.")
            else:
                executar(canal, data, args.saida_pasta, usar_scraper)
        except (ValueError, FileNotFoundError, PermissionError) as exc:
            print(f"\nErro: {exc}")

        if not perguntar_sim_nao("\nRodar outra conciliacao?", padrao_sim=False):
            break


if __name__ == "__main__":
    main()
