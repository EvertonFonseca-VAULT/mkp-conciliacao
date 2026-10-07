import pytest
from unittest.mock import MagicMock, patch


@pytest.fixture
def window(qtbot):
    from PyQt6.QtWidgets import QWidget
    with patch("app.main_window.TelaOperacao",   return_value=QWidget()), \
         patch("app.main_window.TelaConciliacao", return_value=QWidget()), \
         patch("app.main_window.TelaConfig",      return_value=QWidget()):
        from app.main_window import MainWindow
        w = MainWindow()
        qtbot.addWidget(w)
        return w


def test_show_tela_muda_indice(window):
    window.show_tela(1)
    assert window._stack.currentIndex() == 1
    window.show_tela(2)
    assert window._stack.currentIndex() == 2
    window.show_tela(0)
    assert window._stack.currentIndex() == 0


def test_canal_e_data_iniciam_vazios(window):
    assert window.canal == ""
    assert window.data  == ""
