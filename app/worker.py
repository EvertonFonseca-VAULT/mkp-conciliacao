import os
import re
import sys
import subprocess
from pathlib import Path
from PyQt6.QtCore import QThread, pyqtSignal
from app.config import get_credentials, get_paths

_DATA_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

# Canal UI → chave em CANAIS_KONCILI (CO05 → COLETAS_CO05, renomeado 29/09/2026)
_UI_TO_KONCILI = {
    "FULL": "FULL", "ES": "ES", "MOVEIS": "MOVEIS", "ELETRO": "ELETRO",
    "CENTER": "CENTER", "CO05": "COLETAS_CO05", "PRIME": "PRIME", "WEB": "WEB",
}

# (script relativo a AB/, argumentos com {data} substituível) ou None = não implementado
_EMPILHADORES = {
    "FULL":   ("empilhar.py",           ["FULL",  "{data}", "--publicar"]),
    "ES":     ("empilhar.py",           ["ES",    "{data}", "--publicar"]),
    "CENTER": ("empilhar_center.py",    ["{data}", "--publicar"]),
    "WEB":    ("empilhar_web.py",       ["{data}", "--publicar"]),
    "PRIME":  ("empilhar_prime.py",     ["{data}", "--publicar"]),
    "MOVEIS": ("empilhar_moveis.py",    ["{data}", "--publicar"]),
    "CO05":   ("empilhar_coletas05.py", ["{data}", "--publicar"]),
    "ELETRO": None,
}


def build_command(
    step: str,
    canal: str,
    data: str,
    usar_scraper: bool = False,
) -> dict | None:
    """Retorna {"cmd": [...], "cwd": str, "env": dict} ou None se não implementado.
    Levanta ValueError para data fora do formato dd.mm.aaaa."""
    if not _DATA_RE.match(data):
        raise ValueError(f"Data deve estar no formato dd.mm.aaaa, recebido: {data!r}")

    paths = get_paths()
    creds = get_credentials()
    ab_dir = str(Path(paths["projeto_root"]) / "AB")
    rb_dir = str(Path(paths["projeto_root"]) / "AB" / "RB")
    py     = sys.executable

    env = os.environ.copy()
    env["KONCILI_USUARIO"] = creds["usuario"]
    env["KONCILI_SENHA"]   = creds["senha"]

    if step == "extrair_ab":
        canal_k = _UI_TO_KONCILI.get(canal, canal)
        return {"cmd": [py, "koncili.py", canal_k, data, "--salvar"],
                "cwd": ab_dir, "env": env}

    if step == "para_processar":
        return {"cmd": [py, "rodar_canal.py", canal, data],
                "cwd": ab_dir, "env": env}

    if step == "empilhar":
        entry = _EMPILHADORES.get(canal)
        if entry is None:
            return None
        script, args = entry
        resolved = [a.replace("{data}", data) for a in args]
        return {"cmd": [py, script] + resolved, "cwd": ab_dir, "env": env}

    if step == "conciliar":
        cmd = [py, "interativo.py", canal, data]
        if usar_scraper:
            cmd.append("--scraper")
        return {"cmd": cmd, "cwd": rb_dir, "env": env}

    raise ValueError(f"step desconhecido: {step!r}")


class WorkerThread(QThread):
    line_emitted = pyqtSignal(str)
    finished     = pyqtSignal(bool)
    error        = pyqtSignal(str)

    def __init__(self, cmd_dict: dict, parent=None):
        super().__init__(parent)
        self._cmd_dict = cmd_dict

    def run(self):
        try:
            proc = subprocess.Popen(
                self._cmd_dict["cmd"],
                cwd=self._cmd_dict.get("cwd"),
                env=self._cmd_dict.get("env"),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            for line in proc.stdout:
                self.line_emitted.emit(line.rstrip())
            proc.wait()
            self.finished.emit(proc.returncode == 0)
        except Exception as exc:
            self.error.emit(str(exc))
            self.finished.emit(False)
