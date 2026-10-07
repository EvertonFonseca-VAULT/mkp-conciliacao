import pytest
from unittest.mock import patch, MagicMock

_CREDS = {"usuario": "u@test.com", "senha": "pw"}
_PATHS = {"projeto_root": "C:/Mkp", "rodar_totvs": "C:/RODAR TOTVS",
          "pasta_saida": "C:/Conc", "x_drive": "X:/"}


@pytest.fixture
def tela(qtbot):
    with patch("app.tela_config.get_credentials", return_value=_CREDS), \
         patch("app.tela_config.get_paths",       return_value=_PATHS), \
         patch("app.tela_config.save_credentials"), \
         patch("app.tela_config.save_paths"):
        from app.tela_config import TelaConfig
        t = TelaConfig(main_window=MagicMock())
        qtbot.addWidget(t)
        return t


def test_campos_preenchidos_ao_abrir(tela):
    assert tela._campo_usuario.text() == "u@test.com"
    assert tela._campo_senha.text()   == "pw"


def test_salvar_chama_save_credentials(tela, qtbot):
    from PyQt6.QtCore import Qt
    with patch("app.tela_config.save_credentials") as mock_save, \
         patch("app.tela_config.save_paths"):
        tela._campo_usuario.setText("novo@test.com")
        tela._campo_senha.setText("novaSenha")
        qtbot.mouseClick(tela._btn_salvar, Qt.MouseButton.LeftButton)
        mock_save.assert_called_once_with("novo@test.com", "novaSenha")
