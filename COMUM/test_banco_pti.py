"""Self-check do banco do PTI em SQLite (a conexao Fabric nao da pra testar
fora da maquina com login Microsoft; o SQL e o mesmo)."""

import os
import sqlite3
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(__file__))
import banco_pti as b  # noqa: E402


def _pti(vendedor="000578", pedidos=("A1", "A2"), valor=10.5):
    linhas = []
    for p in pedidos:
        reg = {c: "x" for c in b.COLUNAS}
        reg.update({"VENDEDOR": vendedor, "PED. MKP": p, "VALOR REAIS": valor, "VALOR": valor, "FRETE": 0.0})
        linhas.append(reg)
    return pd.DataFrame(linhas, columns=list(b.COLUNAS))


@pytest.fixture
def conn():
    c = sqlite3.connect(":memory:")
    b.criar_tabelas(c)
    b.criar_tabelas(c)  # idempotente
    return c


def test_ida_e_volta_mantem_formato(conn):
    b.importar_pti(conn, "WEB", _pti())
    r = b.buscar_pti(conn, "WEB")
    assert list(r.columns) == list(b.COLUNAS)
    assert len(r) == 2 and r["VALOR REAIS"].dtype.kind == "f" and r["VALOR REAIS"].iloc[0] == 10.5


def test_reimportar_substitui_so_o_mesmo_vendedor(conn):
    b.importar_pti(conn, "FULL", _pti("000750", ("A",)))
    b.importar_pti(conn, "FULL", _pti("000751", ("B",)))
    res = b.importar_pti(conn, "FULL", _pti("000750", ("A", "C")))
    assert res["substituidas"] == 1
    assert sorted(b.buscar_pti(conn, "FULL")["PED. MKP"]) == ["A", "B", "C"]


def test_canais_isolados_e_filtro_por_pedido(conn):
    b.importar_pti(conn, "WEB", _pti(pedidos=("A1", "A2")))
    b.importar_pti(conn, "ES", _pti("000763", ("A1",)))
    assert len(b.buscar_pti(conn, "WEB", ["A2", "ZZ"])) == 1
    assert len(b.buscar_pti(conn, "ES")) == 1
    assert b.buscar_pti(conn, "WEB", []).empty


def test_busca_com_mais_pedidos_que_um_lote(conn):
    ped = [f"P{i}" for i in range(b._LOTE_IN * 2 + 5)]
    b.importar_pti(conn, "WEB", _pti(pedidos=ped))
    assert len(b.buscar_pti(conn, "WEB", ped)) == len(ped)


def test_falha_no_meio_preserva_o_dado_antigo(conn):
    b.importar_pti(conn, "WEB", _pti(pedidos=("A1",)))
    ruim = _pti(pedidos=("N1",)).astype(object)
    ruim.at[0, "LOJA"] = object()  # estoura no INSERT, depois do DELETE
    with pytest.raises(Exception):
        b.importar_pti(conn, "WEB", ruim)
    assert list(b.buscar_pti(conn, "WEB")["PED. MKP"]) == ["A1"]


def test_cabecalho_diferente_e_recusado(conn):
    with pytest.raises(ValueError, match="Cabecalho"):
        b.importar_pti(conn, "WEB", _pti().drop(columns=["FRETE"]))


def test_vazio_nao_apaga(conn):
    b.importar_pti(conn, "WEB", _pti())
    with pytest.raises(ValueError):
        b.importar_pti(conn, "WEB", _pti(pedidos=()))
    assert len(b.buscar_pti(conn, "WEB")) == 2


def test_conectar_sem_config_da_erro_claro(monkeypatch):
    monkeypatch.delenv(b.ENV_CONN, raising=False)
    monkeypatch.delenv(b.ENV_SQLITE, raising=False)
    with pytest.raises(RuntimeError, match=b.ENV_CONN):
        b.conectar()
