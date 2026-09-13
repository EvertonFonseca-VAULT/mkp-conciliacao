"""Roda o pipeline de conciliacao Mercado (FULL) por lote e registra o
progresso em SQLite, pra saber o que ja foi rodado e o que falta.

Uso:
    python pipeline.py sync [pasta_ab]          # descobre lotes novos na pasta do AB
    python pipeline.py rodar <lote> [pti] [--force]
    python pipeline.py status

`lote` e a data no nome do arquivo, ex "08.09.2026" (de "MERCADO FULL 08.09.2026.xlsx").

O banco (conciliacao.db, ao lado deste script) pode ser copiado/movido pra
uma pasta de rede compartilhada pra todo mundo ver o mesmo status.
"""
import os
import re
import sqlite3
import sys
from datetime import datetime

_AQUI = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(_AQUI, "FULL"))
sys.path.insert(0, os.path.join(_AQUI, "MOVEIS"))

import pandas as pd
from rodar_full import rodar, CAMINHO_AB_PADRAO, CAMINHO_PTI_PADRAO
from modelo_ab import resumo_acao_manual

DB = os.path.join(_AQUI, "conciliacao.db")
PASTA_AB_PADRAO = os.path.dirname(CAMINHO_AB_PADRAO)
REGEX_LOTE = re.compile(r"(\d{2}\.\d{2}\.\d{4})\.xlsx$")


def conectar():
    con = sqlite3.connect(DB)
    con.execute(
        """CREATE TABLE IF NOT EXISTS lotes (
            canal TEXT NOT NULL,
            lote TEXT NOT NULL,
            caminho_ab TEXT NOT NULL,
            caminho_pti TEXT,
            caminho_saida TEXT,
            status TEXT NOT NULL DEFAULT 'pendente',
            pedidos INTEGER,
            acao_manual INTEGER,
            rodado_em TEXT,
            erro TEXT,
            PRIMARY KEY (canal, lote)
        )"""
    )
    return con


def sync(pasta=PASTA_AB_PADRAO, canal="FULL"):
    """Varre a pasta do AB e cadastra como 'pendente' todo lote ainda nao
    registrado no banco (nao mexe nos que ja existem)."""
    con = conectar()
    novos = 0
    for nome in os.listdir(pasta):
        m = REGEX_LOTE.search(nome)
        if not m:
            continue
        lote = m.group(1)
        caminho = os.path.join(pasta, nome)
        cur = con.execute(
            "INSERT OR IGNORE INTO lotes (canal, lote, caminho_ab) VALUES (?, ?, ?)",
            (canal, lote, caminho),
        )
        novos += cur.rowcount
    con.commit()
    print(f"{novos} lote(s) novo(s) adicionado(s).")


def rodar_lote(lote, caminho_pti=None, canal="FULL", force=False):
    con = conectar()
    row = con.execute(
        "SELECT status, caminho_ab FROM lotes WHERE canal=? AND lote=?", (canal, lote)
    ).fetchone()
    if row is None:
        print(f"Lote {lote} nao encontrado no banco. Rode 'sync' primeiro.")
        return
    status_atual, caminho_ab = row
    if status_atual == "rodado" and not force:
        print(f"Lote {lote} ja rodado. Use --force pra repetir.")
        return

    caminho_pti = caminho_pti or CAMINHO_PTI_PADRAO
    saida = os.path.join(_AQUI, "FULL", f"resultado_full_{lote}.xlsx")
    try:
        validacao, _pagamento, _frete, processar = rodar(caminho_ab, caminho_pti)
        manual = resumo_acao_manual(validacao)
        with pd.ExcelWriter(saida) as writer:
            processar.to_excel(writer, sheet_name="Para Processar", index=False)
    except Exception as e:
        con.execute(
            "UPDATE lotes SET status='erro', erro=? WHERE canal=? AND lote=?",
            (str(e), canal, lote),
        )
        con.commit()
        raise

    con.execute(
        """UPDATE lotes SET status='rodado', caminho_pti=?, caminho_saida=?,
           pedidos=?, acao_manual=?, rodado_em=?, erro=NULL
           WHERE canal=? AND lote=?""",
        (caminho_pti, saida, len(validacao), len(manual),
         datetime.now().isoformat(timespec="seconds"), canal, lote),
    )
    con.commit()
    print(f"Lote {lote}: {len(validacao)} pedidos, {len(manual)} em acao manual -> {saida}")


def status():
    con = conectar()
    linhas = con.execute(
        "SELECT canal, lote, status, pedidos, acao_manual, rodado_em FROM lotes ORDER BY canal, lote"
    ).fetchall()
    if not linhas:
        print("Nenhum lote registrado. Rode 'sync' primeiro.")
        return
    print(f"{'CANAL':6} {'LOTE':12} {'STATUS':10} {'PEDIDOS':8} {'ACAO MANUAL':12} RODADO EM")
    for canal, lote, st, pedidos, acao, rodado_em in linhas:
        print(f"{canal:6} {lote:12} {st:10} {str(pedidos or '-'):8} {str(acao or '-'):12} {rodado_em or '-'}")


def _teste():
    """Checa a camada de banco (sync idempotente, transicao de status) sem
    depender dos arquivos de rede do AB/TOTVS/PTI."""
    import tempfile
    global DB
    original_db = DB
    with tempfile.TemporaryDirectory() as tmp:
        DB = os.path.join(tmp, "teste.db")
        con = conectar()
        con.execute("INSERT OR IGNORE INTO lotes (canal, lote, caminho_ab) VALUES ('FULL', '01.01.2026', 'x.xlsx')")
        con.execute("INSERT OR IGNORE INTO lotes (canal, lote, caminho_ab) VALUES ('FULL', '01.01.2026', 'x.xlsx')")
        con.commit()
        assert con.execute("SELECT COUNT(*) FROM lotes").fetchone()[0] == 1, "sync deveria ser idempotente"

        con.execute(
            "UPDATE lotes SET status='rodado', pedidos=10, acao_manual=2 WHERE canal='FULL' AND lote='01.01.2026'"
        )
        con.commit()
        linha = con.execute("SELECT status, pedidos, acao_manual FROM lotes").fetchone()
        assert linha == ("rodado", 10, 2), linha
        con.close()
    DB = original_db
    print("ok")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    cmd = sys.argv[1]
    if cmd == "sync":
        sync(sys.argv[2] if len(sys.argv) > 2 else PASTA_AB_PADRAO)
    elif cmd == "rodar":
        if len(sys.argv) < 3:
            print("Uso: python pipeline.py rodar <lote> [pti] [--force]")
            return
        lote = sys.argv[2]
        force = "--force" in sys.argv
        resto = [a for a in sys.argv[3:] if a != "--force"]
        rodar_lote(lote, resto[0] if resto else None, force=force)
    elif cmd == "status":
        status()
    elif cmd == "teste":
        _teste()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
