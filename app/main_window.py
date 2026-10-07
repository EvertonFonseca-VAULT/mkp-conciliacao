from PyQt6.QtWidgets import QMainWindow, QStackedWidget

try:
    from app.tela_operacao    import TelaOperacao
except ImportError:
    from PyQt6.QtWidgets import QWidget as TelaOperacao  # type: ignore

try:
    from app.tela_conciliacao import TelaConciliacao
except ImportError:
    from PyQt6.QtWidgets import QWidget as TelaConciliacao  # type: ignore

try:
    from app.tela_config      import TelaConfig
except ImportError:
    from PyQt6.QtWidgets import QWidget as TelaConfig  # type: ignore


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
