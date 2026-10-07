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
