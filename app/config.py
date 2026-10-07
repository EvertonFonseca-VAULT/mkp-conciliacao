import os
from pathlib import Path
import keyring
from PyQt6.QtCore import QSettings

_SERVICE = "mkp-conciliacao"
_ORG     = "WebContinental"
_APP     = "MkpConciliacao"

PROJECT_ROOT = str(Path(__file__).parent.parent)

_DEFAULTS = {
    "projeto_root": PROJECT_ROOT,
    "rodar_totvs":  r"C:\RODAR TOTVS",
    "pasta_saida":  rf"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\TEMP\CONCILIACAO\{os.environ.get('USERNAME', 'usuario')}",
    "x_drive":      "X:\\",
}


def save_credentials(usuario: str, senha: str) -> None:
    keyring.set_password(_SERVICE, "KONCILI_USUARIO", usuario)
    keyring.set_password(_SERVICE, "KONCILI_SENHA",   senha)


def get_credentials() -> dict:
    return {
        "usuario": keyring.get_password(_SERVICE, "KONCILI_USUARIO") or "",
        "senha":   keyring.get_password(_SERVICE, "KONCILI_SENHA")   or "",
    }


def save_paths(paths: dict) -> None:
    s = QSettings(_ORG, _APP)
    for k, v in paths.items():
        s.setValue(k, v)


def get_paths() -> dict:
    s = QSettings(_ORG, _APP)
    return {k: s.value(k, default) for k, default in _DEFAULTS.items()}
