import pytest
from unittest.mock import patch, MagicMock
from PyQt6.QtCore import Qt


@pytest.fixture
def tela(qtbot):
    with patch("app.tela_operacao.WorkerThread"), \
         patch("app.tela_operacao.build_command",
               return_value={"cmd": ["py"], "cwd": ".", "env": {}}), \
         patch("app.tela_operacao.get_credentials",
               return_value={"usuario": "u", "senha": "p"}):
        from app.tela_operacao import TelaOperacao
        t = TelaOperacao(main_window=MagicMock())
        qtbot.addWidget(t)
        return t


def test_data_invalida_desabilita_botoes(tela):
    tela._campo_data.setText("abc")
    tela._validar_inputs()
    assert not tela._btn_extrair.isEnabled()
    assert not tela._btn_pp.isEnabled()


def test_data_valida_habilita_botoes(tela):
    tela._campo_data.setText("22.09.2026")
    tela._validar_inputs()
    assert tela._btn_extrair.isEnabled()
    assert tela._btn_conciliar.isEnabled()


def test_clicar_extrair_inicia_worker(tela, qtbot):
    tela._campo_data.setText("22.09.2026")
    tela._validar_inputs()
    with patch("app.tela_operacao.build_command",
               return_value={"cmd": ["py"], "cwd": ".", "env": {}}) as mock_build, \
         patch("app.tela_operacao.WorkerThread") as MockWT, \
         patch("app.tela_operacao.get_credentials",
               return_value={"usuario": "u", "senha": "p"}):
        inst = MagicMock()
        MockWT.return_value = inst
        qtbot.mouseClick(tela._btn_extrair, Qt.MouseButton.LeftButton)
        mock_build.assert_called_once_with("extrair_ab", "FULL", "22.09.2026")
        inst.start.assert_called_once()


def test_empilhar_eletro_mostra_aviso_no_log(tela, qtbot):
    tela._campo_data.setText("22.09.2026")
    tela._validar_inputs()
    tela._combo_canal.setCurrentText("ELETRO")
    with patch("app.tela_operacao.build_command", return_value=None):
        qtbot.mouseClick(tela._btn_empilhar, Qt.MouseButton.LeftButton)
    assert "não implementado" in tela._log.toPlainText().lower()


def test_sem_credenciais_abre_config(tela, qtbot):
    tela._campo_data.setText("22.09.2026")
    tela._validar_inputs()
    with patch("app.tela_operacao.get_credentials",
               return_value={"usuario": "", "senha": ""}):
        qtbot.mouseClick(tela._btn_extrair, Qt.MouseButton.LeftButton)
    tela._mw.show_tela.assert_called_with(2)


def test_clicar_conciliar_passa_canal_data_para_mainwindow(tela, qtbot):
    tela._combo_canal.setCurrentText("ES")
    tela._campo_data.setText("15.09.2026")
    tela._validar_inputs()
    qtbot.mouseClick(tela._btn_conciliar, Qt.MouseButton.LeftButton)
    assert tela._mw.canal == "ES"
    assert tela._mw.data  == "15.09.2026"
    tela._mw.show_tela.assert_called_with(1)
