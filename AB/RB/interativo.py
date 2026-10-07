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

from canais import CAMINHO_RB_PADRAO, CANAIS, CANAIS_SEM_RAIZ_CONFIRMADA
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


def gerar_bat_scraper(canal, data, pasta_saida):
    """.bat de 2 cliques que reroda essa mesma conciliacao com --scraper --
    fica salvo na RAIZ (pasta_saida, do lado do iniciar_conciliacao.bat), nao
    dentro de Conciliacoes/: so reinvoca com --saida-pasta=pasta_saida, nao
    precisa morar junto do dashboard/xlsx. Pedido do Everton, 15/09/2026 --
    antes ia pra pasta_destino (Conciliacoes/) so por causa do calculo de
    --saida-pasta, e misturava .bat com arquivo de dado."""
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    nome_bat = f"rodar_scraper_{sufixo}.bat"
    caminho_bat = os.path.join(pasta_saida, nome_bat)
    conteudo = (
        "@echo off\n"
        f'"{SCRAPER_PYTHON}" "{os.path.abspath(__file__)}" '
        f'--canal {canal} --data {data} --saida-pasta "{pasta_saida}" --scraper\n'
        "pause\n"
    )
    with open(caminho_bat, "w", encoding="cp1252") as f:
        f.write(conteudo)
    return nome_bat


def gerar_bat_atualizar(canal, data, pasta_saida):
    """.bat de 2 cliques que reroda essa conciliacao SEM scraper -- so pra
    reler rb.xlsx/rbx.xlsx (AUXILIAR/) e o RB/AB mais recentes e regerar o
    dashboard, sem precisar abrir o Chrome/scraper (rapido, so pandas).
    Pedido do Everton, 16/09/2026: botao pra atualizar a conferencia
    Original x Transitoria depois de lancar mais movimentos. Fica salvo na
    RAIZ (pasta_saida), ver gerar_bat_scraper."""
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    nome_bat = f"atualizar_{sufixo}.bat"
    caminho_bat = os.path.join(pasta_saida, nome_bat)
    conteudo = (
        "@echo off\n"
        f'"{sys.executable}" "{os.path.abspath(__file__)}" '
        f'--canal {canal} --data {data} --saida-pasta "{pasta_saida}"\n'
        "pause\n"
    )
    with open(caminho_bat, "w", encoding="cp1252") as f:
        f.write(conteudo)
    return nome_bat


def gerar_bat_arquivar(canal, data, pasta_saida):
    """.bat de 2 cliques que arquiva uma copia do rbx.xlsx (AUXILIAR/) em
    CAMINHO_ARQUIVO_MARKETPLACE, com as formulas E3/F3/G3 -- ver
    rodar_conciliacao_rb._arquivar_rbx. So gerado quando a Transitoria fica
    pronta pra transferencia (ver executar()) -- NAO roda sozinho, precisa
    do clique do Everton, pra nao arquivar antes da transferencia manual de
    verdade ter acontecido. Fica salvo na RAIZ (pasta_saida), mesmo padrao
    de gerar_bat_scraper/gerar_bat_atualizar. Sem parametro raiz: o proprio
    main() (--arquivar-rbx) reconsulta CANAIS[canal]["raiz"] na hora, o .bat
    so precisa levar canal/data. Pedido do Everton, 17/09/2026."""
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    nome_bat = f"arquivar_rbx_{sufixo}.bat"
    caminho_bat = os.path.join(pasta_saida, nome_bat)
    conteudo = (
        "@echo off\n"
        f'"{sys.executable}" "{os.path.abspath(__file__)}" '
        f'--canal {canal} --data {data} --saida-pasta "{pasta_saida}" --arquivar-rbx\n'
        "pause\n"
    )
    with open(caminho_bat, "w", encoding="cp1252") as f:
        f.write(conteudo)
    return nome_bat


def gerar_bat_rebate_fora_da_regra(canal, data, pasta_saida, pasta_destino):
    """.bat de arrastar-e-soltar: o Everton arrasta o arquivo de resultado
    do TOTVS (marketplace...xls, nome com timestamp imprevisivel -- nao da
    pra saber o caminho na hora de gerar o .bat) em cima do .bat e solta;
    %~1 recebe o caminho largado. Gerado todo dia (nao condicional, ao
    contrario do gerar_bat_arquivar) -- ver rebate_fora_da_regra.py.

    Fica em pasta_destino (Conciliacoes/REBATE/, a mesma pasta onde
    PROCESSAR_REBATE_<data>.xlsx ja cai e onde o Everton salva o arquivo
    de recusa do TOTVS) -- excecao explicita a regra geral de "bat na
    RAIZ" dos outros (pedido do proprio Everton, 18/09/2026, pra arrastar
    sem trocar de pasta). O conteudo continua reinvocando com
    --saida-pasta=pasta_saida (a RAIZ de verdade, nao pasta_destino) --
    isso nao muda, so ONDE o arquivo .bat fica salvo em disco. REBATE so e
    criada quando tem rebate de verdade no dia (ver
    rodar_conciliacao_rb._caminho_rebate_processar) -- pode nao existir
    ainda."""
    sufixo = f"{canal.lower()}_{data.replace('.', '')}"
    nome_bat = f"rebate_fora_da_regra_{sufixo}.bat"
    os.makedirs(pasta_destino, exist_ok=True)
    caminho_bat = os.path.join(pasta_destino, nome_bat)
    conteudo = (
        "@echo off\n"
        'if "%~1"=="" (\n'
        '  echo Arraste o arquivo de resultado do TOTVS ^(marketplace...xls^) para cima deste .bat e solte.\n'
        "  pause\n"
        "  exit /b 1\n"
        ")\n"
        f'"{sys.executable}" "{os.path.abspath(__file__)}" '
        f'--canal {canal} --data {data} --saida-pasta "{pasta_saida}" --rebate-fora-da-regra "%~1"\n'
        "pause\n"
    )
    with open(caminho_bat, "w", encoding="cp1252") as f:
        f.write(conteudo)
    return nome_bat


def _lotes_status(canal):
    """Le o status dos lotes do pipeline AB bruto -> Para Processar (ver
    Mkp/AB/pipeline.py, conciliacao.db) pra guia 2 do dashboard -- [] quando
    o banco ainda nao existe (pipeline.py nunca rodou, feature opcional,
    mesmo padrao de conferencia_rbx=None). So leitura, nao cria o banco.
    Pedido do Everton, 18/09/2026."""
    caminho_db = os.path.join(os.path.dirname(__file__), "..", "conciliacao.db")
    if not os.path.exists(caminho_db):
        return []
    import sqlite3

    con = sqlite3.connect(caminho_db)
    try:
        return con.execute(
            "SELECT lote, status, pedidos, acao_manual, rodado_em FROM lotes "
            "WHERE canal=? ORDER BY lote",
            (canal,),
        ).fetchall()
    finally:
        con.close()


def executar(canal, data, pasta_saida, usar_scraper):
    print(f"\nProcessando {canal} {data}...")
    r = processar(canal, data, pasta_saida=pasta_saida, usar_scraper=usar_scraper)
    pasta_conciliacoes = os.path.join(pasta_saida, "Conciliacoes")
    raiz_confirmada = r["raiz"] if canal not in CANAIS_SEM_RAIZ_CONFIRMADA else None

    caminho_bat_scraper = None
    if not usar_scraper and len(r["fila_verificar"]):
        caminho_bat_scraper = gerar_bat_scraper(canal, data, pasta_saida)
    caminho_bat_atualizar = gerar_bat_atualizar(canal, data, pasta_saida)
    gerar_bat_rebate_fora_da_regra(canal, data, pasta_saida, os.path.join(pasta_conciliacoes, "REBATE"))
    # So gera o bat de arquivar quando a Transitoria bate com o alvo (pronta
    # pra transferencia) -- NAO arquiva sozinho: o Everton ainda precisa
    # clicar, depois de fazer a transferencia manual de verdade. Pedido do
    # Everton, 17/09/2026.
    caminho_bat_arquivar = None
    if r["conferencia_rbx"] and r["conferencia_rbx"]["pronta_para_transferencia"]:
        caminho_bat_arquivar = gerar_bat_arquivar(canal, data, pasta_saida)

    caminho_html_esperado = os.path.join(pasta_conciliacoes, f"dashboard_{data.replace('.', '')}.html")
    dashboard_ja_existia = os.path.exists(caminho_html_esperado)

    caminho_html = gerar_dashboard(
        r["resultado"], r["fila_verificar"], r["rebate_cmm"],
        canal=canal, data=data, data_ref_rb=r["data_ref"],
        caminho_xlsx_completo=os.path.basename(r["caminho_xlsx"]),
        caminho_xlsx_rebate=os.path.basename(r["caminho_rebate_xlsx"]),
        pasta_saida=pasta_conciliacoes, alvo=r["alvo"], ml_atualizado=r["ml_atualizado"],
        # Os .bat ficam na RAIZ (pasta_saida), 1 nivel acima do dashboard.html
        # (Conciliacoes/) -- ver gerar_bat_scraper/gerar_bat_atualizar.
        caminho_bat_scraper=f"../{caminho_bat_scraper}" if caminho_bat_scraper else None,
        raiz=raiz_confirmada,
        # Descricoes agora e pasta irma de Conciliacoes (onde o dashboard.html
        # fica) -- basename sozinho quebraria o link, precisa subir 1 nivel.
        caminho_xlsx_descricoes=(
            f"../Descricoes/{os.path.basename(r['caminho_descricoes_xlsx'])}"
            if r["caminho_descricoes_xlsx"] else None
        ),
        conferencia_rbx=r["conferencia_rbx"],
        valores_pedidos=r["valores_pedidos"],
        caminho_bat_atualizar=f"../{caminho_bat_atualizar}",
        caminho_bat_arquivar=f"../{caminho_bat_arquivar}" if caminho_bat_arquivar else None,
        ml_de_cache=r["ml_de_cache"], responsavel=r["responsavel"],
        rebate_processar=r["rebate_processar"],
        # REBATE e subpasta de Conciliacoes (onde o dashboard.html ja fica),
        # nao irma dela -- sem "../" (diferente de Descricoes, que e irma).
        caminho_xlsx_rebate_processar=(
            f"REBATE/{os.path.basename(r['caminho_rebate_processar_xlsx'])}"
            if r["caminho_rebate_processar_xlsx"] else None
        ),
        para_processar=r["para_processar"],
        lotes_status=_lotes_status(canal),
    )

    print(f"Total de pedidos: {len(r['resultado'])}")
    for status, qtd in r["resultado"]["STATUS"].value_counts().items():
        print(f"  {status}: {qtd}")
    print(f"Dashboard do dia ({canal}): {caminho_html}")
    # Resumo parseavel pelo .bat (para modo data unica)
    hm = r["resultado"].loc[:, "H:M"].sum() if "H:M" in r["resultado"].columns else 0
    verificar_qtd = len(r["fila_verificar"])
    alvo = r["alvo"]
    total_rb = r["resultado"]["TOTAL_RB"].sum() if "TOTAL_RB" in r["resultado"].columns else 0
    bateu_calendario = "SIM" if alvo and abs(total_rb - alvo) < 0.01 else ("NAO" if alvo else "SEM ALVO")
    print(f"RESUMO|{canal}|{data}|H:M={hm:,.2f}|VERIFICAR={verificar_qtd}|CALENDARIO={bateu_calendario}")

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


def _datas_do_intervalo(data_inicio, data_fim):
    """Gera lista de datas dd.mm de inicio a fim (inclusive)."""
    from datetime import datetime, timedelta
    fmt = "%d.%m"
    try:
        inicio = datetime.strptime(data_inicio, fmt)
        fim = datetime.strptime(data_fim, fmt)
    except ValueError:
        fmt = "%d.%m.%Y"
        inicio = datetime.strptime(data_inicio, fmt)
        fim = datetime.strptime(data_fim, fmt)
    datas = []
    atual = inicio
    while atual <= fim:
        datas.append(atual.strftime("%d.%m"))
        atual += timedelta(days=1)
    return datas


def main():
    parser = argparse.ArgumentParser()
    # Aceita tanto --canal/--data (flag) quanto canal/data (posicional) --
    # o .bat usa posicional p/ simplicidade. Ordem: canal data
    parser.add_argument("canal_pos", nargs="?", help="Canal (posicional, opcional)")
    parser.add_argument("data_pos", nargs="?", help="Data dd.mm (posicional, opcional)")
    parser.add_argument("--canal", dest="canal_flag")
    parser.add_argument("--data", dest="data_flag")
    parser.add_argument("--data-inicio", dest="data_inicio")
    parser.add_argument("--data-fim", dest="data_fim")
    parser.add_argument("--saida-pasta", default=PASTA_REDE_PADRAO)
    parser.add_argument("--scraper", action="store_true")
    parser.add_argument("--arquivar-rbx", action="store_true")
    parser.add_argument("--rebate-fora-da-regra")  # caminho do arquivo largado no .bat, nao store_true
    args = parser.parse_args()

    # Precedencia: flag > posicional
    canal = (args.canal_flag or args.canal_pos)
    data = (args.data_flag or args.data_pos)

    if args.rebate_fora_da_regra:
        # Reinvocado pelo bat de arrastar-e-soltar (ver
        # gerar_bat_rebate_fora_da_regra) -- so gera o _cmm.xlsx e sai, nao
        # roda o resto do pipeline. Pedido do Everton, 18/09/2026.
        from parser_rb import metadata_rb
        from rebate_fora_da_regra import gerar_rebate_fora_da_regra_cmm

        canal = (args.canal_flag or args.canal_pos).upper()
        data_ref, responsavel = metadata_rb(CAMINHO_RB_PADRAO)
        raiz = CANAIS[canal]["raiz"]
        df = gerar_rebate_fora_da_regra_cmm(args.rebate_fora_da_regra, raiz, data_ref, responsavel)
        caminho_saida = args.rebate_fora_da_regra.rsplit(".", 1)[0] + "_cmm.xlsx"
        df.to_excel(caminho_saida, index=False)
        print(f"{len(df)} pedido(s) fora da regra -> {caminho_saida}")
        return

    if args.arquivar_rbx:
        # Reinvocado pelo bat de 2 cliques (ver gerar_bat_arquivar) -- so
        # arquiva e sai, nao roda o resto do pipeline (agregar/cruzar/
        # dashboard). Pedido do Everton, 17/09/2026.
        from rodar_conciliacao_rb import _arquivar_rbx

        canal = (args.canal_flag or args.canal_pos).upper()
        data = args.data_flag or args.data_pos
        caminho_rbx = os.path.join(args.saida_pasta, "AUXILIAR", "rbx.xlsx")
        caminho_destino = _arquivar_rbx(canal, data, CANAIS[canal]["raiz"], caminho_rbx)
        print(f"RBX arquivado: {caminho_destino}")
        return

    # Processamento em intervalo de datas (--data-inicio --data-fim)
    if canal and args.data_inicio and args.data_fim:
        datas = _datas_do_intervalo(args.data_inicio, args.data_fim)
        for data in datas:
            try:
                r = processar(canal.upper(), data, args.saida_pasta, args.scraper)
                # Imprime resumo parseavel pelo .bat
                hm = r["resultado"].loc[:, "H:M"].sum() if "H:M" in r["resultado"].columns else 0
                verificar_qtd = len(r["fila_verificar"])
                alvo = r["alvo"]
                total_rb = r["resultado"]["TOTAL_RB"].sum() if "TOTAL_RB" in r["resultado"].columns else 0
                bateu_calendario = "SIM" if alvo and abs(total_rb - alvo) < 0.01 else ("NAO" if alvo else "SEM ALVO")
                print(f"RESUMO|{canal}|{data}|H:M={hm:,.2f}|VERIFICAR={verificar_qtd}|CALENDARIO={bateu_calendario}")
                # Gera dashboard (mesma logica de executar mas sem abrir navegador repetidamente)
                from dashboard import gerar_dashboard
                from canais import CANAIS_SEM_RAIZ_CONFIRMADA
                pasta_conciliacoes = os.path.join(args.saida_pasta, "Conciliacoes")
                raiz_confirmada = r["raiz"] if canal not in CANAIS_SEM_RAIZ_CONFIRMADA else None
                gerar_dashboard(
                    r["resultado"], r["fila_verificar"], r["rebate_cmm"],
                    canal=canal, data=data, data_ref_rb=r["data_ref"],
                    caminho_xlsx_completo=os.path.basename(r["caminho_xlsx"]),
                    caminho_xlsx_rebate=os.path.basename(r["caminho_rebate_xlsx"]),
                    pasta_saida=pasta_conciliacoes, alvo=r["alvo"], ml_atualizado=r["ml_atualizado"],
                    caminho_bat_scraper=None,
                    raiz=raiz_confirmada,
                    caminho_xlsx_descricoes=(
                        f"../Descricoes/{os.path.basename(r['caminho_descricoes_xlsx'])}"
                        if r["caminho_descricoes_xlsx"] else None
                    ),
                    conferencia_rbx=r["conferencia_rbx"],
                    valores_pedidos=r["valores_pedidos"],
                    caminho_bat_atualizar=None,
                    caminho_bat_arquivar=None,
                    ml_de_cache=r["ml_de_cache"], responsavel=r["responsavel"],
                    rebate_processar=r["rebate_processar"],
                    caminho_xlsx_rebate_processar=(
                        f"REBATE/{os.path.basename(r['caminho_rebate_processar_xlsx'])}"
                        if r["caminho_rebate_processar_xlsx"] else None
                    ),
                    para_processar=r["para_processar"],
                    lotes_status=_lotes_status(canal),
                )
            except (ValueError, FileNotFoundError, PermissionError) as exc:
                print(f"ERRO|{canal}|{data}|{exc}")
        return

    if canal and data:
        # re-execucao interna (com o Python certo pro scraper) -- sem perguntas
        try:
            executar(canal.upper(), data, args.saida_pasta, args.scraper)
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


def _testar_bats():
    """gerar_bat_scraper/gerar_bat_atualizar tem que salvar na RAIZ
    (pasta_saida), nao em Conciliacoes/ -- pedido do Everton, 15/09/2026:
    os .bat todos juntos, so xlsx/html dentro das subpastas."""
    import tempfile

    with tempfile.TemporaryDirectory() as pasta_saida:
        nome_scraper = gerar_bat_scraper("FULL", "01.09", pasta_saida)
        nome_atualizar = gerar_bat_atualizar("FULL", "01.09", pasta_saida)
        nome_arquivar = gerar_bat_arquivar("FULL", "01.09", pasta_saida)
        # pasta_destino (Conciliacoes/REBATE/) de proposito NAO criada
        # antes -- dia sintetico sem nenhum rebate, REBATE so e criada
        # por rodar_conciliacao_rb.py quando tem rebate de verdade. O bat
        # tem que criar sozinho (os.makedirs), senao a escrita falha.
        pasta_rebate = os.path.join(pasta_saida, "Conciliacoes", "REBATE")
        assert not os.path.exists(pasta_rebate)
        nome_fora_da_regra = gerar_bat_rebate_fora_da_regra("FULL", "01.09", pasta_saida, pasta_rebate)
        assert "WEB" in CANAIS, "WEB nao cadastrado em CANAIS"
        assert nome_scraper == "rodar_scraper_full_0109.bat", nome_scraper
        assert nome_atualizar == "atualizar_full_0109.bat", nome_atualizar
        assert nome_arquivar == "arquivar_rbx_full_0109.bat", nome_arquivar
        assert nome_fora_da_regra == "rebate_fora_da_regra_full_0109.bat", nome_fora_da_regra
        assert os.path.exists(os.path.join(pasta_saida, nome_scraper))
        assert os.path.exists(os.path.join(pasta_saida, nome_atualizar))
        assert os.path.exists(os.path.join(pasta_saida, nome_arquivar))
        # excecao explicita: esse bat fica em Conciliacoes/REBATE/, nao na
        # RAIZ como os outros 3 -- pedido do Everton (arrastar sem trocar
        # de pasta, mesmo lugar do PROCESSAR_REBATE_<data>.xlsx).
        assert not os.path.exists(os.path.join(pasta_saida, nome_fora_da_regra))
        assert os.path.exists(os.path.join(pasta_rebate, nome_fora_da_regra))
        # nao pode ter ido pra dentro de nenhuma subpasta
        assert not os.path.exists(os.path.join(pasta_saida, "Conciliacoes", nome_scraper))
        # o bat reinvoca com --arquivar-rbx, sem rodar o resto do pipeline
        conteudo_arquivar = open(os.path.join(pasta_saida, nome_arquivar), encoding="cp1252").read()
        assert "--arquivar-rbx" in conteudo_arquivar, conteudo_arquivar
        # arrastar-e-soltar: sem arquivo, avisa e sai (exit /b 1) sem travar
        # no resto do .bat; com arquivo, %~1 vira o caminho pro flag; e o
        # --saida-pasta embutido continua sendo a RAIZ de verdade (pasta_saida),
        # nao a pasta_destino onde o .bat foi parar.
        conteudo_fora_da_regra = open(os.path.join(pasta_rebate, nome_fora_da_regra), encoding="cp1252").read()
        assert 'if "%~1"=="" (' in conteudo_fora_da_regra
        assert "exit /b 1" in conteudo_fora_da_regra
        assert '--rebate-fora-da-regra "%~1"' in conteudo_fora_da_regra
        assert f'--saida-pasta "{pasta_saida}"' in conteudo_fora_da_regra
    print("interativo.py: ok")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _testar_bats()
    else:
        main()
