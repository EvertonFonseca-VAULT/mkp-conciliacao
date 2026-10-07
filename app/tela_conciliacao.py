import re
import shutil
from pathlib import Path
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                              QPushButton, QCheckBox, QFileDialog, QLabel)
from PyQt6.QtCore import Qt, QUrl
try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
except ImportError:
    from PyQt6.QtWidgets import QTextEdit as QWebEngineView  # type: ignore

from app.worker import build_command, WorkerThread
from app.config import get_paths

_DASHBOARD_RE = re.compile(r"^Dashboard do dia \([^)]+\):\s*(.+)$")


class TelaConciliacao(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self._mw        = main_window
        self._rb_path   = ""
        self._rbx_path  = ""
        self._html_path = ""
        self._worker    = None

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
