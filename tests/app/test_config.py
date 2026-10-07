from unittest.mock import patch, MagicMock
import app.config as config


def test_save_and_get_credentials():
    with patch("keyring.set_password") as mock_set, \
         patch("keyring.get_password", side_effect=["u@test.com", "s3cr3t"]):
        config.save_credentials("u@test.com", "s3cr3t")
        result = config.get_credentials()
    assert mock_set.call_count == 2
    assert result == {"usuario": "u@test.com", "senha": "s3cr3t"}


def test_get_credentials_ausentes_retorna_vazios():
    with patch("keyring.get_password", return_value=None):
        result = config.get_credentials()
    assert result == {"usuario": "", "senha": ""}


def test_save_and_get_paths():
    stored = {}
    mock_s = MagicMock()
    mock_s.setValue.side_effect = lambda k, v: stored.update({k: v})
    mock_s.value.side_effect    = lambda k, d: stored.get(k, d)

    with patch("app.config.QSettings", return_value=mock_s):
        config.save_paths({"rodar_totvs": r"D:\OUTRO"})
        result = config.get_paths()

    assert result["rodar_totvs"] == r"D:\OUTRO"
    assert result["x_drive"] == config._DEFAULTS["x_drive"]
