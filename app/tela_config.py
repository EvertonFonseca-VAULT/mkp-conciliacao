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

        grp = QGroupBox("Koncili")
        form = QFormLayout(grp)
        self._campo_usuario = QLineEdit()
        self._campo_senha   = QLineEdit()
        self._campo_senha.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("Usuário:", self._campo_usuario)
        form.addRow("Senha:",   self._campo_senha)
        self._btn_salvar = QPushButton("Salvar credenciais")
        self._btn_salvar.clicked.connect(self._salvar)
        form.addRow(self._btn_salvar)
        layout.addWidget(grp)

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
