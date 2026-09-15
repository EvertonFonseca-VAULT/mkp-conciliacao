"""Self-check pro interativo.py -- garante que rodar FULL e depois ES no
mesmo dia so abre 1 aba de navegador (a 2a rodada so sobrescreve o
dashboard existente, ver comentario "ponytail" em interativo.executar)."""

import os
import webbrowser

chamadas = []
webbrowser.open = lambda caminho: chamadas.append(caminho)

import interativo  # noqa: E402 (import depois do mock de webbrowser.open)
from canais import CAMINHO_RB_PADRAO  # noqa: E402
from parser_rb import metadata_rb  # noqa: E402

# Data pega do RB real (CONCILIADOR_RB.xlsx) em vez de fixa -- esse arquivo
# e reexportado todo dia como parte do uso normal, entao fixar "01.09" aqui
# quebra sozinho quando o Everton exporta um RB novo.
_dia, _mes = metadata_rb(CAMINHO_RB_PADRAO)[0].split("/")[:2]
DATA = f"{_dia}.{_mes}"
SUFIXO_ARQUIVO = f"{_dia}{_mes}"

PASTA = os.path.join(os.path.dirname(__file__), "_preview")
CAMINHO_DASHBOARD = os.path.join(PASTA, "Conciliacoes", f"dashboard_{SUFIXO_ARQUIVO}.html")
if os.path.exists(CAMINHO_DASHBOARD):
    os.remove(CAMINHO_DASHBOARD)

interativo.executar("FULL", DATA, PASTA, False)
interativo.executar("ES", DATA, PASTA, False)

assert len(chamadas) == 1, f"deveria abrir a aba so na 1a vez do dia, abriu {len(chamadas)}x"
print("test_interativo.py: ok (abre so 1 vez, mesmo processando FULL e ES)")
