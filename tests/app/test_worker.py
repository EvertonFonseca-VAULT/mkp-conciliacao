import sys
from pathlib import Path
from unittest.mock import patch

_PATHS = {
    "projeto_root": str(Path(__file__).parent.parent.parent),
    "rodar_totvs":  r"C:\RODAR TOTVS",
    "pasta_saida":  r"C:\Conciliacoes",
    "x_drive":      "X:\\",
}
_CREDS = {"usuario": "u@test.com", "senha": "pw123"}


def _cmd(step, canal="FULL", data="22.09.2026", usar_scraper=False):
    with patch("app.worker.get_paths",       return_value=_PATHS), \
         patch("app.worker.get_credentials", return_value=_CREDS):
        from app.worker import build_command
        return build_command(step, canal, data, usar_scraper)


def test_extrair_ab_full():
    r = _cmd("extrair_ab", canal="FULL", data="22.09.2026")
    assert "koncili.py" in " ".join(r["cmd"])
    assert "FULL" in r["cmd"]
    assert "22.09.2026" in r["cmd"]
    assert "--salvar" in r["cmd"]
    assert r["env"]["KONCILI_USUARIO"] == "u@test.com"


def test_extrair_ab_co05_mapeia_para_coletas_co05():
    r = _cmd("extrair_ab", canal="CO05")
    assert "COLETAS_CO05" in r["cmd"]


def test_para_processar():
    r = _cmd("para_processar", canal="FULL", data="22.09.2026")
    assert "rodar_canal.py" in " ".join(r["cmd"])
    assert "FULL" in r["cmd"]
    assert "22.09.2026" in r["cmd"]


def test_empilhar_full_usa_empilhar_py():
    r = _cmd("empilhar", canal="FULL", data="22.09.2026")
    assert "empilhar.py" in " ".join(r["cmd"])
    assert "FULL" in r["cmd"]
    assert "--publicar" in r["cmd"]


def test_empilhar_center_usa_empilhar_center_py():
    r = _cmd("empilhar", canal="CENTER", data="22.09.2026")
    assert "empilhar_center.py" in " ".join(r["cmd"])
    assert "--publicar" in r["cmd"]


def test_empilhar_web():
    r = _cmd("empilhar", canal="WEB", data="22.09.2026")
    assert "empilhar_web.py" in " ".join(r["cmd"])


def test_empilhar_prime():
    r = _cmd("empilhar", canal="PRIME", data="22.09.2026")
    assert "empilhar_prime.py" in " ".join(r["cmd"])


def test_empilhar_moveis():
    r = _cmd("empilhar", canal="MOVEIS", data="22.09.2026")
    assert "empilhar_moveis.py" in " ".join(r["cmd"])


def test_empilhar_co05():
    r = _cmd("empilhar", canal="CO05", data="22.09.2026")
    assert "empilhar_coletas05.py" in " ".join(r["cmd"])


def test_empilhar_eletro_retorna_none():
    r = _cmd("empilhar", canal="ELETRO")
    assert r is None


def test_conciliar_sem_scraper():
    r = _cmd("conciliar", canal="FULL", data="22.09.2026")
    assert "interativo.py" in " ".join(r["cmd"])
    assert "FULL" in r["cmd"]
    assert "--scraper" not in r["cmd"]


def test_conciliar_com_scraper():
    r = _cmd("conciliar", canal="FULL", data="22.09.2026", usar_scraper=True)
    assert "--scraper" in r["cmd"]


def test_data_invalida_levanta_valueerror():
    import pytest
    with pytest.raises(ValueError, match="dd.mm.aaaa"):
        _cmd("extrair_ab", data="abc")
