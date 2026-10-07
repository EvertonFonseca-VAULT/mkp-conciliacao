import pytest
from unittest.mock import patch, MagicMock
from PyQt6.QtCore import Qt


@pytest.fixture
def mw():
    m = MagicMock()
    m.canal = "FULL"
    m.data  = "22.09.2026"
    return m


@pytest.fixture
def tela(qtbot, mw):
    from PyQt6.QtWidgets import QWidget
    with patch("app.tela_conciliacao.QWebEngineView", QWidget):
        from app.tela_conciliacao import TelaConciliacao
        t = TelaConciliacao(main_window=mw)
        qtbot.addWidget(t)
        return t


def test_botao_desabilitado_sem_arquivos(tela):
    assert not tela._btn_conciliar.isEnabled()


def test_botao_habilitado_com_dois_arquivos(tela):
    tela._rb_path  = r"C:\rb.xlsx"
    tela._rbx_path = r"C:\rbx.xlsx"
    tela._atualizar_btn()
    assert tela._btn_conciliar.isEnabled()


def test_botao_desabilitado_com_apenas_rb(tela):
    tela._rb_path  = r"C:\rb.xlsx"
    tela._rbx_path = ""
    tela._atualizar_btn()
    assert not tela._btn_conciliar.isEnabled()


def test_clicar_conciliar_inicia_worker(tela, qtbot, mw):
    tela._rb_path  = r"C:\rb.xlsx"
    tela._rbx_path = r"C:\rbx.xlsx"
    tela._atualizar_btn()
    with patch("app.tela_conciliacao.build_command",
               return_value={"cmd": ["py"], "cwd": ".", "env": {}}) as mock_build, \
         patch("app.tela_conciliacao.WorkerThread") as MockWT, \
         patch("shutil.copy2"):
        inst = MagicMock()
        MockWT.return_value = inst
        qtbot.mouseClick(tela._btn_conciliar, Qt.MouseButton.LeftButton)
        mock_build.assert_called_once_with("conciliar", "FULL", "22.09.2026", usar_scraper=False)
        inst.start.assert_called_once()


def test_linha_dashboard_detecta_html(tela):
    caminho = r"C:\Conciliacoes\Conciliacao_FULL.html"
    tela._processar_linha(f"Dashboard do dia (FULL): {caminho}")
    assert tela._html_path == caminho


def test_linha_sem_dashboard_nao_altera_html_path(tela):
    tela._html_path = ""
    tela._processar_linha("Processando pedido 123")
    assert tela._html_path == ""
