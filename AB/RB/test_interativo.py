"""Self-check pro interativo.py -- garante que rodar FULL e depois ES no
mesmo dia so abre 1 aba de navegador (a 2a rodada so sobrescreve o
dashboard existente, ver comentario "ponytail" em interativo.executar)."""

import os
import webbrowser

chamadas = []
webbrowser.open = lambda caminho: chamadas.append(caminho)

import interativo  # noqa: E402 (import depois do mock de webbrowser.open)

PASTA = os.path.join(os.path.dirname(__file__), "_preview")
CAMINHO_DASHBOARD = os.path.join(PASTA, "dashboard_0109.html")
if os.path.exists(CAMINHO_DASHBOARD):
    os.remove(CAMINHO_DASHBOARD)

interativo.executar("FULL", "01.09", PASTA, False)
interativo.executar("ES", "01.09", PASTA, False)

assert len(chamadas) == 1, f"deveria abrir a aba so na 1a vez do dia, abriu {len(chamadas)}x"
print("test_interativo.py: ok (abre so 1 vez, mesmo processando FULL e ES)")
