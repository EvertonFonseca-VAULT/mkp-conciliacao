# App Desktop Conciliação Mercado — Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Construir um app desktop PyQt6 com 3 telas que orquestra o fluxo diário completo (Extrair AB → Para Processar → Empilhar → Conciliar + Dashboard) sem que o operador precise abrir terminal ou navegar por pastas.

**Architecture:** `QMainWindow` com `QStackedWidget` (3 telas). Toda chamada de script roda via `subprocess.Popen` dentro de um `WorkerThread(QThread)` — UI nunca trava. Output chega ao log linha a linha via `pyqtSignal`. Credenciais Koncili em `keyring`; caminhos em `QSettings`. Tela 2 detecta o caminho do HTML gerado a partir da linha `Dashboard do dia (…): <caminho>` que `interativo.py` já imprime no stdout.

**Tech Stack:** Python 3.11+, PyQt6 ≥ 6.6, PyQt6-WebEngine ≥ 6.6, keyring ≥ 24, pytest ≥ 7, pytest-qt ≥ 4.4, pyinstaller (só build)

**Spec:** `docs/superpowers/specs/2026-10-07-app-desktop-mkp-design.md`

## Global Constraints

- Python ≥ 3.11 (venv do projeto atual)
- Nenhum arquivo em `AB/` é modificado — o app chama os scripts existentes via `subprocess`
- Subprocess usa sempre `sys.executable` para garantir o venv correto
- Credenciais nunca em arquivo de texto claro — apenas `keyring`
- Todo código novo em `app/`; testes em `tests/app/`
- `PASTA_REDE_PADRAO` padrão: `X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\TEMP\CONCILIACAO\<user>` (lido de `interativo.py` ou editável na Tela 3)
- Canal `CO05` na UI corresponde à chave `COLETAS_CO05` em `CANAIS_KONCILI` (renomeado em 29/09/2026)

## Review Focus

1. **RB com Dt.Ref errada** — `interativo.py` recusa e imprime erro; o log da Tela 1/2 deve mostrá-lo sem engolir silenciosamente. Coberto: `WorkerThread` retransmite todo stdout via `line_emitted`.
2. **Credenciais ausentes ao clicar Extrair AB** — deve abrir Tela 3 com aviso, nunca travar. Coberto: `test_sem_credenciais_abre_config` na Task 6.
3. **Canal ELETRO no empilhador** — não existe `empilhar_eletro.py`; deve exibir mensagem clara, nunca crashar. Coberto: `test_empilhar_eletro_mostra_aviso_no_log` na Task 6.
4. **Data em formato errado** — campo aceita texto livre; validar `dd.mm.aaaa` antes de disparar qualquer script. Coberto: `test_data_invalida_desabilita_botoes` na Task 6.
5. **HTML do dashboard não encontrado** — `interativo.py` pode terminar sem imprimir a linha `Dashboard do dia` (erro no meio). Coberto: `_carregar_dashboard` verifica se a linha foi capturada antes de tentar carregar.

---

## Mapa de Arquivos

| Arquivo | Responsabilidade |
|---|---|
| `app/__init__.py` | vazio |
| `app/main.py` | `QApplication`, cria `MainWindow`, `app.exec()` |
| `app/config.py` | `get/save_credentials` (keyring), `get/save_paths` (QSettings), defaults |
| `app/worker.py` | `build_command()` pura + `WorkerThread(QThread)` |
| `app/main_window.py` | `MainWindow` + `QStackedWidget` + `show_tela(n)` |
| `app/tela_operacao.py` | Tela 1 — canal/data, 4 botões, log, indicadores |
| `app/tela_conciliacao.py` | Tela 2 — seletores RB, scraper, botão rodar, `QWebEngineView` |
| `app/tela_config.py` | Tela 3 — form credenciais + caminhos editáveis |
| `app/build.bat` | `pyinstaller --onefile --windowed` |
| `tests/app/test_config.py` | round-trip credenciais e paths (mock keyring/QSettings) |
| `tests/app/test_worker.py` | `build_command` para cada canal e etapa |
| `tests/app/test_tela_operacao.py` | botões, log, abertura da Tela 2 |
| `tests/app/test_tela_conciliacao.py` | habilitação do botão Conciliar, detecção do HTML |
| `tests/app/test_tela_config.py` | save/load pelo form |

---

### Task 1: Scaffolding e dependências

**Files:**
- Modify: `requirements.txt`
- Create: `app/__init__.py`, `app/main.py`
- Create: `tests/__init__.py`, `tests/app/__init__.py`

**Interfaces:**
- Produces: `python app/main.py` abre uma janela vazia e fecha normalmente

- [ ] **Step 1: Adicionar dependências ao requirements.txt**

Acrescentar ao final do arquivo existente:

```
PyQt6>=6.6.0
PyQt6-WebEngine>=6.6.0
keyring>=24.0.0
pytest-qt>=4.4.0
pyinstaller
```

- [ ] **Step 2: Instalar**

```bash
pip install -r requirements.txt
```

- [ ] **Step 3: Criar estrutura de pastas e arquivos vazios**

```bash
mkdir app tests tests\app
type nul > app\__init__.py
type nul > tests\__init__.py
type nul > tests\app\__init__.py
```

- [ ] **Step 4: Criar `app/main.py`**

```python
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow

def main():
    app = QApplication(sys.argv)
    window = QMainWindow()
    window.setWindowTitle("Conciliação Mercado")
    window.resize(1000, 750)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
```

- [ ] **Step 5: Verificar que abre**

```bash
python app/main.py
```

Esperado: janela "Conciliação Mercado" abre sem erro.

- [ ] **Step 6: Commit**

```bash
git add requirements.txt app\ tests\
git commit -m "feat(app): scaffolding PyQt6 — janela vazia"
```

---

### Task 2: config.py — credenciais e caminhos

**Files:**
- Create: `app/config.py`
- Create: `tests/app/test_config.py`

**Interfaces:**
- Produces:
  - `get_credentials() -> dict` — `{"usuario": str, "senha": str}` (vazios se não configurado)
  - `save_credentials(usuario: str, senha: str) -> None`
  - `get_paths() -> dict` — chaves: `"projeto_root"`, `"rodar_totvs"`, `"pasta_saida"`, `"x_drive"`
  - `save_paths(paths: dict) -> None`
  - `_DEFAULTS: dict` — valores padrão dos caminhos (importado por tela_config.py)

- [ ] **Step 1: Escrever os testes**

```python
# tests/app/test_config.py
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
    # chave não salva retorna o default
    assert result["x_drive"] == config._DEFAULTS["x_drive"]
```

- [ ] **Step 2: Rodar para confirmar falha**

```bash
pytest tests/app/test_config.py -v
```

Esperado: `ModuleNotFoundError` ou `AttributeError`.

- [ ] **Step 3: Implementar `app/config.py`**

```python
import os
from pathlib import Path
import keyring
from PyQt6.QtCore import QSettings

_SERVICE = "mkp-conciliacao"
_ORG    = "WebContinental"
_APP    = "MkpConciliacao"

PROJECT_ROOT = str(Path(__file__).parent.parent)

_DEFAULTS = {
    "projeto_root": PROJECT_ROOT,
    "rodar_totvs":  r"C:\RODAR TOTVS",
    # Mesmo padrão de PASTA_REDE_PADRAO em interativo.py.
    # Editar na Tela 3 se o caminho da máquina for diferente.
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
```

- [ ] **Step 4: Rodar e confirmar que passa**

```bash
pytest tests/app/test_config.py -v
```

Esperado: 3 testes PASS.

- [ ] **Step 5: Commit**

```bash
git add app/config.py tests/app/test_config.py
git commit -m "feat(app): config.py — credenciais (keyring) e caminhos (QSettings)"
```

---

### Task 3: worker.py — build_command e WorkerThread

**Files:**
- Create: `app/worker.py`
- Create: `tests/app/test_worker.py`

**Interfaces:**
- Consumes: `get_credentials`, `get_paths` de `app.config`
- Produces:
  - `build_command(step, canal, data, usar_scraper) -> dict | None` — `{"cmd": list[str], "cwd": str, "env": dict}` ou `None` (não implementado)
  - `WorkerThread(cmd_dict: dict, parent=None)` com sinais `line_emitted(str)`, `finished(bool)`, `error(str)`

`step` aceita: `"extrair_ab"` | `"para_processar"` | `"empilhar"` | `"conciliar"`

- [ ] **Step 1: Escrever os testes de build_command**

```python
# tests/app/test_worker.py
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
    # CO05 na UI → COLETAS_CO05 em CANAIS_KONCILI
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
```

- [ ] **Step 2: Rodar para confirmar falha**

```bash
pytest tests/app/test_worker.py -v
```

Esperado: `ImportError` ou `ModuleNotFoundError`.

- [ ] **Step 3: Implementar `app/worker.py`**

```python
import os
import re
import sys
import subprocess
from pathlib import Path
from PyQt6.QtCore import QThread, pyqtSignal
from app.config import get_credentials, get_paths

_DATA_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

# Canal UI → chave em CANAIS_KONCILI (renomeia CO05 → COLETAS_CO05, 29/09/2026)
_UI_TO_KONCILI = {
    "FULL": "FULL", "ES": "ES", "MOVEIS": "MOVEIS", "ELETRO": "ELETRO",
    "CENTER": "CENTER", "CO05": "COLETAS_CO05", "PRIME": "PRIME", "WEB": "WEB",
}

# (script relativo a AB/, argumentos com {data} substituível) ou None = não implementado
_EMPILHADORES = {
    "FULL":   ("empilhar.py",          ["FULL",  "{data}", "--publicar"]),
    "ES":     ("empilhar.py",          ["ES",    "{data}", "--publicar"]),
    "CENTER": ("empilhar_center.py",   ["{data}", "--publicar"]),
    "WEB":    ("empilhar_web.py",      ["{data}", "--publicar"]),
    "PRIME":  ("empilhar_prime.py",    ["{data}", "--publicar"]),
    "MOVEIS": ("empilhar_moveis.py",   ["{data}", "--publicar"]),
    "CO05":   ("empilhar_coletas05.py",["{data}", "--publicar"]),
    "ELETRO": None,
}


def build_command(
    step: str,
    canal: str,
    data: str,
    usar_scraper: bool = False,
) -> dict | None:
    """
    Retorna {"cmd": [...], "cwd": str, "env": dict} ou None se não implementado.
    Levanta ValueError para data fora do formato dd.mm.aaaa.
    """
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
    finished     = pyqtSignal(bool)   # True = returncode 0
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
```

- [ ] **Step 4: Rodar e confirmar que passa**

```bash
pytest tests/app/test_worker.py -v
```

Esperado: todos PASS.

- [ ] **Step 5: Commit**

```bash
git add app/worker.py tests/app/test_worker.py
git commit -m "feat(app): worker.py — build_command e WorkerThread"
```

---

### Task 4: main_window.py — estrutura e navegação

**Files:**
- Modify: `app/main.py`
- Create: `app/main_window.py`
- Create: `tests/app/test_main_window.py`

**Interfaces:**
- Produces:
  - `MainWindow.show_tela(n: int)` — muda para a tela n (0=Operação, 1=Conciliação, 2=Config)
  - `MainWindow.canal: str`, `MainWindow.data: str` — escrito pela Tela 1, lido pela Tela 2

- [ ] **Step 1: Escrever o teste**

```python
# tests/app/test_main_window.py
import pytest
from unittest.mock import MagicMock, patch


@pytest.fixture
def window(qtbot):
    with patch("app.main_window.TelaOperacao",   return_value=MagicMock()), \
         patch("app.main_window.TelaConciliacao", return_value=MagicMock()), \
         patch("app.main_window.TelaConfig",      return_value=MagicMock()):
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
```

- [ ] **Step 2: Rodar para confirmar falha**

```bash
pytest tests/app/test_main_window.py -v
```

- [ ] **Step 3: Criar `app/main_window.py`**

```python
from PyQt6.QtWidgets import QMainWindow, QStackedWidget

# Imports tardios evitam circular; as telas são criadas após este módulo existir
try:
    from app.tela_operacao    import TelaOperacao
except ImportError:
    from PyQt6.QtWidgets import QWidget as TelaOperacao

try:
    from app.tela_conciliacao import TelaConciliacao
except ImportError:
    from PyQt6.QtWidgets import QWidget as TelaConciliacao

try:
    from app.tela_config      import TelaConfig
except ImportError:
    from PyQt6.QtWidgets import QWidget as TelaConfig


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conciliação Mercado")
        self.resize(1000, 750)
        self.canal: str = ""
        self.data:  str = ""

        self._stack = QStackedWidget()
        self.setCentralWidget(self._stack)

        self._tela_op   = TelaOperacao(main_window=self)
        self._tela_conc = TelaConciliacao(main_window=self)
        self._tela_cfg  = TelaConfig(main_window=self)

        self._stack.addWidget(self._tela_op)    # índice 0
        self._stack.addWidget(self._tela_conc)  # índice 1
        self._stack.addWidget(self._tela_cfg)   # índice 2

    def show_tela(self, n: int) -> None:
        self._stack.setCurrentIndex(n)
```

- [ ] **Step 4: Atualizar `app/main.py`**

```python
import sys
from PyQt6.QtWidgets import QApplication
from app.main_window import MainWindow

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
```

- [ ] **Step 5: Rodar testes**

```bash
pytest tests/app/test_main_window.py -v
```

Esperado: PASS.

- [ ] **Step 6: Commit**

```bash
git add app/main.py app/main_window.py tests/app/test_main_window.py
git commit -m "feat(app): main_window — QStackedWidget + show_tela()"
```

---

### Task 5: Tela 3 — Configurações

**Files:**
- Create: `app/tela_config.py`
- Create: `tests/app/test_tela_config.py`

**Interfaces:**
- Consumes: `save_credentials`, `get_credentials`, `save_paths`, `get_paths`, `_DEFAULTS` de `app.config`
- Produces: `TelaConfig(QWidget)` com construtor `TelaConfig(main_window)`

- [ ] **Step 1: Escrever testes**

```python
# tests/app/test_tela_config.py
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
```

- [ ] **Step 2: Rodar para confirmar falha**

```bash
pytest tests/app/test_tela_config.py -v
```

- [ ] **Step 3: Implementar `app/tela_config.py`**

```python
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                              QLineEdit, QPushButton, QGroupBox, QFormLayout)
from PyQt6.QtCore import Qt
from app.config import get_credentials, save_credentials, get_paths, save_paths, _DEFAULTS


class TelaConfig(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self._mw = main_window
        layout = QVBoxLayout(self)

        btn_voltar = QPushButton("← Voltar")
        btn_voltar.clicked.connect(lambda: self._mw.show_tela(0))
        layout.addWidget(btn_voltar, alignment=Qt.AlignmentFlag.AlignLeft)

        # Bloco Koncili
        grp = QGroupBox("Koncili")
        form = QFormLayout(grp)
        self._campo_usuario = QLineEdit()
        self._campo_senha   = QLineEdit()
        self._campo_senha.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("Usuário:", self._campo_usuario)
        form.addRow("Senha:",   self._campo_senha)
        self._btn_salvar = QPushButton("💾 Salvar credenciais")
        self._btn_salvar.clicked.connect(self._salvar)
        form.addRow(self._btn_salvar)
        layout.addWidget(grp)

        # Bloco caminhos (colapsável)
        grp_p = QGroupBox("Caminhos (avançado)")
        grp_p.setCheckable(True)
        grp_p.setChecked(False)
        form_p = QFormLayout(grp_p)
        self._campos_paths: dict[str, QLineEdit] = {}
        for k, label in [("rodar_totvs", "RODAR TOTVS:"),
                         ("pasta_saida", "Pasta saída:"),
                         ("x_drive",    "Unidade X:")]:
            c = QLineEdit()
            form_p.addRow(label, c)
            self._campos_paths[k] = c
        btn_def = QPushButton("Restaurar padrões")
        btn_def.clicked.connect(self._restaurar)
        form_p.addRow(btn_def)
        layout.addWidget(grp_p)

        layout.addStretch()
        self._carregar()

    def _carregar(self):
        creds = get_credentials()
        self._campo_usuario.setText(creds["usuario"])
        self._campo_senha.setText(creds["senha"])
        paths = get_paths()
        for k, c in self._campos_paths.items():
            c.setText(paths.get(k, ""))

    def _salvar(self):
        save_credentials(self._campo_usuario.text(), self._campo_senha.text())
        save_paths({k: c.text() for k, c in self._campos_paths.items()})

    def _restaurar(self):
        for k, c in self._campos_paths.items():
            c.setText(_DEFAULTS.get(k, ""))
```

- [ ] **Step 4: Rodar testes**

```bash
pytest tests/app/test_tela_config.py -v
```

Esperado: PASS.

- [ ] **Step 5: Commit**

```bash
git add app/tela_config.py tests/app/test_tela_config.py
git commit -m "feat(app): tela_config — credenciais Koncili e caminhos"
```

---

### Task 6: Tela 1 — Operação

**Files:**
- Create: `app/tela_operacao.py`
- Create: `tests/app/test_tela_operacao.py`

**Interfaces:**
- Consumes: `build_command`, `WorkerThread` de `app.worker`; `get_credentials` de `app.config`
- Produces: `TelaOperacao(QWidget)` com construtor `TelaOperacao(main_window)`

- [ ] **Step 1: Escrever testes**

```python
# tests/app/test_tela_operacao.py
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
         patch("app.tela_operacao.WorkerThread") as MockWT:
        inst = MagicMock()
        MockWT.return_value = inst
        qtbot.mouseClick(tela._btn_extrair, Qt.MouseButton.LeftButton)
        mock_build.assert_called_once_with("extrair_ab", "FULL", "22.09.2026", False)
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
```

- [ ] **Step 2: Rodar para confirmar falha**

```bash
pytest tests/app/test_tela_operacao.py -v
```

- [ ] **Step 3: Implementar `app/tela_operacao.py`**

```python
import re
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                              QComboBox, QLineEdit, QPushButton,
                              QTextEdit, QGridLayout)
from PyQt6.QtCore import Qt
from app.worker import build_command, WorkerThread
from app.config import get_credentials

_CANAIS  = ["FULL", "ES", "MOVEIS", "ELETRO", "CENTER", "CO05", "PRIME", "WEB"]
_DATA_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


class TelaOperacao(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self._mw     = main_window
        self._worker = None
        layout = QVBoxLayout(self)

        # Cabeçalho
        hdr = QHBoxLayout()
        hdr.addWidget(QLabel("<h2>Conciliação Mercado</h2>"))
        hdr.addStretch()
        btn_cfg = QPushButton("⚙")
        btn_cfg.setFixedWidth(32)
        btn_cfg.clicked.connect(lambda: self._mw.show_tela(2))
        hdr.addWidget(btn_cfg)
        layout.addLayout(hdr)

        # Seletores
        sel = QHBoxLayout()
        self._combo_canal = QComboBox()
        self._combo_canal.addItems(_CANAIS)
        self._campo_data  = QLineEdit()
        self._campo_data.setPlaceholderText("dd.mm.aaaa")
        self._campo_data.textChanged.connect(self._validar_inputs)
        sel.addWidget(QLabel("Canal:"))
        sel.addWidget(self._combo_canal)
        sel.addWidget(QLabel("Data:"))
        sel.addWidget(self._campo_data)
        sel.addStretch()
        layout.addLayout(sel)

        # Botões 2×2
        grid = QGridLayout()
        self._btn_extrair   = QPushButton("1. Extrair AB")
        self._btn_pp        = QPushButton("2. Para Processar")
        self._btn_empilhar  = QPushButton("3. Empilhar")
        self._btn_conciliar = QPushButton("4. Conciliar →")
        for btn in (self._btn_extrair, self._btn_pp,
                    self._btn_empilhar, self._btn_conciliar):
            btn.setEnabled(False)
            btn.setMinimumHeight(52)
        grid.addWidget(self._btn_extrair,   0, 0)
        grid.addWidget(self._btn_pp,        0, 1)
        grid.addWidget(self._btn_empilhar,  1, 0)
        grid.addWidget(self._btn_conciliar, 1, 1)
        layout.addLayout(grid)

        # Log
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        layout.addWidget(self._log)

        self._btn_extrair.clicked.connect(lambda: self._rodar("extrair_ab"))
        self._btn_pp.clicked.connect(lambda: self._rodar("para_processar"))
        self._btn_empilhar.clicked.connect(lambda: self._rodar("empilhar"))
        self._btn_conciliar.clicked.connect(self._abrir_conciliacao)

    def _validar_inputs(self):
        ok = bool(_DATA_RE.match(self._campo_data.text()))
        for btn in (self._btn_extrair, self._btn_pp,
                    self._btn_empilhar, self._btn_conciliar):
            btn.setEnabled(ok)

    def _rodar(self, step: str):
        canal = self._combo_canal.currentText()
        data  = self._campo_data.text()

        if step == "extrair_ab":
            creds = get_credentials()
            if not creds["usuario"] or not creds["senha"]:
                self._log.append("⚠ Credenciais do Koncili não configuradas — abrindo Configurações.")
                self._mw.show_tela(2)
                return

        cmd = build_command(step, canal, data)
        if cmd is None:
            self._log.append(f"⚠ Empilhador para {canal} não implementado ainda.")
            return

        self._log.append(f"\n▶ {step.replace('_', ' ').title()} — {canal} {data}")
        self._worker = WorkerThread(cmd)
        self._worker.line_emitted.connect(self._log.append)
        self._worker.finished.connect(
            lambda ok: self._log.append("✓ Concluído." if ok else "✗ Erro — ver log acima."))
        self._worker.error.connect(lambda e: self._log.append(f"✗ {e}"))
        self._worker.start()

    def _abrir_conciliacao(self):
        self._mw.canal = self._combo_canal.currentText()
        self._mw.data  = self._campo_data.text()
        self._mw.show_tela(1)
```

- [ ] **Step 4: Rodar testes**

```bash
pytest tests/app/test_tela_operacao.py -v
```

Esperado: todos PASS.

- [ ] **Step 5: Commit**

```bash
git add app/tela_operacao.py tests/app/test_tela_operacao.py
git commit -m "feat(app): tela_operacao — 4 botões, log em tempo real e indicadores"
```

---

### Task 7: Tela 2 — Conciliação + Dashboard

**Files:**
- Create: `app/tela_conciliacao.py`
- Create: `tests/app/test_tela_conciliacao.py`

**Interfaces:**
- Consumes: `build_command`, `WorkerThread` de `app.worker`; `get_paths` de `app.config`; `MainWindow.canal`, `MainWindow.data`
- Produces: `TelaConciliacao(QWidget)` com construtor `TelaConciliacao(main_window)`

O caminho do HTML é detectado a partir da linha `Dashboard do dia (CANAL): <caminho>` que `interativo.py` imprime no stdout — sem precisar reconstruir o nome do arquivo.

- [ ] **Step 1: Escrever testes**

```python
# tests/app/test_tela_conciliacao.py
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
    # QWebEngineView precisa de display; usar mock nos testes
    with patch("app.tela_conciliacao.QWebEngineView", MagicMock()):
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
        mock_build.assert_called_once_with("conciliar", "FULL", "22.09.2026", False)
        inst.start.assert_called_once()


def test_linha_dashboard_detecta_html(tela):
    caminho = r"C:\Conciliacoes\Conciliacao_FULL.html"
    tela._processar_linha(f"Dashboard do dia (FULL): {caminho}")
    assert tela._html_path == caminho


def test_linha_sem_dashboard_nao_altera_html_path(tela):
    tela._html_path = ""
    tela._processar_linha("Processando pedido 123")
    assert tela._html_path == ""
```

- [ ] **Step 2: Rodar para confirmar falha**

```bash
pytest tests/app/test_tela_conciliacao.py -v
```

- [ ] **Step 3: Implementar `app/tela_conciliacao.py`**

```python
import re
import shutil
from pathlib import Path
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                              QPushButton, QCheckBox, QFileDialog, QLabel)
from PyQt6.QtCore import Qt, QUrl
try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
except ImportError:
    # fallback para ambientes sem WebEngine (CI, testes)
    from PyQt6.QtWidgets import QTextEdit as QWebEngineView  # type: ignore

from app.worker import build_command, WorkerThread
from app.config import get_paths

_DASHBOARD_RE = re.compile(r"^Dashboard do dia \([^)]+\):\s*(.+)$")


class TelaConciliacao(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self._mw       = main_window
        self._rb_path  = ""
        self._rbx_path = ""
        self._html_path = ""
        self._worker   = None

        layout = QVBoxLayout(self)

        # Cabeçalho
        hdr = QHBoxLayout()
        btn_voltar = QPushButton("← Voltar")
        btn_voltar.clicked.connect(lambda: self._mw.show_tela(0))
        self._titulo = QLabel()
        hdr.addWidget(btn_voltar)
        hdr.addWidget(self._titulo)
        hdr.addStretch()
        layout.addLayout(hdr)

        # Seletores de arquivo
        rb_row = QHBoxLayout()
        btn_rb  = QPushButton("📂 rb.xlsx")
        btn_rbx = QPushButton("📂 rbx.xlsx")
        btn_rb.clicked.connect(lambda: self._selecionar("rb"))
        btn_rbx.clicked.connect(lambda: self._selecionar("rbx"))
        self._lbl_rb  = QLabel("rb.xlsx: —")
        self._lbl_rbx = QLabel("rbx.xlsx: —")
        rb_row.addWidget(btn_rb);  rb_row.addWidget(self._lbl_rb)
        rb_row.addWidget(btn_rbx); rb_row.addWidget(self._lbl_rbx)
        rb_row.addStretch()
        layout.addLayout(rb_row)

        # Controles
        ctrl = QHBoxLayout()
        self._chk_scraper   = QCheckBox("Usar scraper ML/MP")
        self._btn_conciliar = QPushButton("▶ Iniciar Conciliação")
        self._btn_conciliar.setEnabled(False)
        self._btn_conciliar.clicked.connect(self._conciliar)
        ctrl.addWidget(self._chk_scraper)
        ctrl.addStretch()
        ctrl.addWidget(self._btn_conciliar)
        layout.addLayout(ctrl)

        # Dashboard embutido
        self._webview = QWebEngineView()
        layout.addWidget(self._webview, stretch=1)

    def showEvent(self, event):
        super().showEvent(event)
        self._titulo.setText(f"Conciliação — {self._mw.canal} — {self._mw.data}")

    def _selecionar(self, qual: str):
        path, _ = QFileDialog.getOpenFileName(self, f"Selecionar {qual}.xlsx", "", "Excel (*.xlsx)")
        if not path:
            return
        if qual == "rb":
            self._rb_path = path
            self._lbl_rb.setText(f"rb.xlsx: ✓ {Path(path).name}")
        else:
            self._rbx_path = path
            self._lbl_rbx.setText(f"rbx.xlsx: ✓ {Path(path).name}")
        self._atualizar_btn()

    def _atualizar_btn(self):
        self._btn_conciliar.setEnabled(bool(self._rb_path and self._rbx_path))

    def _conciliar(self):
        paths   = get_paths()
        aux_dir = Path(paths["pasta_saida"]) / self._mw.canal / "AUXILIAR"
        aux_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self._rb_path,  aux_dir / "rb.xlsx")
        shutil.copy2(self._rbx_path, aux_dir / "rbx.xlsx")

        self._html_path = ""
        cmd = build_command("conciliar", self._mw.canal, self._mw.data,
                            usar_scraper=self._chk_scraper.isChecked())
        self._worker = WorkerThread(cmd)
        self._worker.line_emitted.connect(self._processar_linha)
        self._worker.finished.connect(self._ao_terminar)
        self._worker.start()

    def _processar_linha(self, line: str):
        m = _DASHBOARD_RE.match(line.strip())
        if m:
            self._html_path = m.group(1).strip()

    def _ao_terminar(self, ok: bool):
        if ok and self._html_path:
            self._webview.load(QUrl.fromLocalFile(self._html_path))
```

- [ ] **Step 4: Rodar testes**

```bash
pytest tests/app/test_tela_conciliacao.py -v
```

Esperado: todos PASS.

- [ ] **Step 5: Commit**

```bash
git add app/tela_conciliacao.py tests/app/test_tela_conciliacao.py
git commit -m "feat(app): tela_conciliacao — RBs, scraper toggle, dashboard embutido"
```

---

### Task 8: Integração e build

**Files:**
- Create: `app/build.bat`

**Interfaces:**
- Produces: `dist\Conciliacao Mercado.exe`

- [ ] **Step 1: Rodar todos os testes**

```bash
pytest tests/app/ -v
```

Esperado: todos PASS.

- [ ] **Step 2: Testar o app completo (percurso do ouro)**

```bash
python app/main.py
```

1. Tela 1 abre — selecionar FULL e `22.09.2026` — 4 botões habilitam  
2. Clicar ⚙ — Tela 3 abre — preencher credenciais — Salvar — Voltar  
3. Clicar `4. Conciliar →` — Tela 2 abre com "Conciliação — FULL — 22.09.2026" no título  
4. Clicar `📂 rb.xlsx` e `📂 rbx.xlsx` — botão `▶ Iniciar Conciliação` habilita  
5. Voltar — Tela 1 — log ainda mostra output anterior  

- [ ] **Step 3: Criar `app/build.bat`**

```bat
@echo off
cd /d "%~dp0.."
pip install pyinstaller
pyinstaller --onefile --windowed ^
    --name "Conciliacao Mercado" ^
    --add-data "AB;AB" ^
    app/main.py
echo.
echo Build concluido: dist\Conciliacao Mercado.exe
pause
```

- [ ] **Step 4: Gerar o .exe**

```bash
app\build.bat
```

Esperado: `dist\Conciliacao Mercado.exe` gerado (~250 MB com WebEngine). Executar o `.exe` e repetir o percurso do passo 2.

- [ ] **Step 5: Commit e push**

```bash
git add app/build.bat
git commit -m "feat(app): build.bat — PyInstaller gera Conciliacao Mercado.exe"
git push
```
