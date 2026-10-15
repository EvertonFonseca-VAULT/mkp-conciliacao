"""PROTOTIPO DESCARTAVEL -- nao e codigo de producao.

Responde uma pergunta de UI: um launcher Tkinter com botoes (canal, data,
scraper, atualizar, arquivar RBX, rebate fora da regra) no lugar do menu
digitado do interativo.py, com as mesmas cores/paleta do dashboard HTML
(dashboard._CSS, tema claro) pra parecer a mesma familia visual.

NAO reimplementa nada: so chama as funcoes que ja existem em interativo.py /
rodar_conciliacao_rb.py / rebate_fora_da_regra.py. O dashboard HTML (tabela,
fila VERIFICAR, checkboxes de natureza) continua abrindo no navegador --
replicar aquilo em Tkinter seria outro projeto, nao um prototipo.

Rodar: duplo-clique em abrir_prototipo_gui.bat (ou "py prototype_gui_launcher.py").

Decisao a validar: o layout (botao por canal + lista de datas + botoes de
acao) resolve o fluxo do dia a dia? Se sim, vira GUI de verdade (e os .bat
gerados por gerar_bat_* podem sair); se nao, descarta sem dor.
"""

import contextlib
import io
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import cmm_writer
import interativo
import mov_writer
from canais import CANAIS
from rebate_cmm import COLUNAS as COLUNAS_REBATE_CMM

# Mapa canal -> nome usado no arquivo CMM real (rede, fora deste repo) --
# ACHADO POR INSPECAO DA PASTA (ARQUIVOS AUXILIARES\CMM_<ano>\), NAO
# CONFIRMADO COM O EVERTON. WEB virou "WEB" a partir de 2025 (antes "WE") --
# so cobre 2025+ aqui, de proposito (pra falhar alto se o ano for antigo
# em vez de escrever no arquivo errado).
NOME_CMM_POR_CANAL = {
    "FULL": "FULL", "ES": "COLETAS_ES", "PRIME": "PRIME", "WEB": "WEB",
    "MOVEIS": "MOVEIS", "ELETRO": "ELETROD", "CENTER": "CE", "CO05": "COLETAS_05",
}
# CMM_PASTA_REDE_TESTE: aponta pra uma pasta de teste (mesma estrutura
# CMM_<ano>\CMM_MELI_<nome>_...xlsx) em vez do arquivo de rede de verdade --
# pra testar os botoes sem risco. Sem essa variavel, usa o caminho real.
PASTA_CMM_REDE = os.environ.get(
    "CMM_PASTA_REDE_TESTE",
    r"X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\ARQUIVOS AUXILIARES",
)


def caminho_cmm_real(canal, ano):
    nome = NOME_CMM_POR_CANAL[canal]
    return os.path.join(PASTA_CMM_REDE, f"CMM_{ano}",
                         f"CMM_MELI_{nome}_Controle de Movimentos Manual {ano}.xlsx")

# Paleta tema claro, copiada de dashboard._CSS (:root) -- so pra "parecer a
# mesma familia visual" do dashboard, sem puxar CSS/JS de verdade.
BG = "#f6f6f4"
SURFACE = "#ffffff"
INK = "#16181a"
INK2 = "#575d63"
LINE = "#e5e5e1"
ACCENT = "#2a56c6"
OK = "#0d7a4d"
BAD = "#b0261c"


class _FilaParaLog(io.TextIOBase):
    """print() dentro da thread de trabalho cai aqui -- a GUI so le a fila."""

    def __init__(self, fila):
        self.fila = fila

    def write(self, texto):
        if texto:
            self.fila.put(texto)
        return len(texto)


class App:
    def __init__(self, root):
        self.root = root
        self.canal = None
        self.data = None
        root.title("Conciliacao ML/TOTVS -- PROTOTIPO")
        root.configure(bg=BG)
        root.geometry("760x620")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Card.TFrame", background=SURFACE)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=INK)
        style.configure("Card.TLabel", background=SURFACE, foreground=INK)
        style.configure("TCheckbutton", background=SURFACE, foreground=INK)
        style.configure("Canal.TButton", padding=8)
        style.configure("Accent.TButton", padding=10, foreground="white", background=ACCENT)
        style.map("Accent.TButton", background=[("active", ACCENT)])

        tk.Label(root, text="PROTOTIPO -- descartavel, so pra validar o layout",
                 bg=BG, fg=INK2, font=("Segoe UI", 9, "italic")).pack(anchor="w", padx=16, pady=(10, 0))

        # Canal
        bloco_canal = ttk.Frame(root, style="Card.TFrame", padding=12)
        bloco_canal.pack(fill="x", padx=16, pady=10)
        ttk.Label(bloco_canal, text="Canal", style="Card.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        linha_canais = ttk.Frame(bloco_canal, style="Card.TFrame")
        linha_canais.pack(fill="x", pady=(6, 0))
        self.botoes_canal = {}
        for nome in CANAIS:
            b = ttk.Button(linha_canais, text=nome, style="Canal.TButton",
                            command=lambda c=nome: self.selecionar_canal(c))
            b.pack(side="left", padx=4)
            self.botoes_canal[nome] = b

        # Data
        bloco_data = ttk.Frame(root, style="Card.TFrame", padding=12)
        bloco_data.pack(fill="x", padx=16, pady=(0, 10))
        ttk.Label(bloco_data, text="Data (\"Para Processar\" encontrados)",
                  style="Card.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.lista_datas = tk.Listbox(bloco_data, height=5, bg=SURFACE, fg=INK,
                                       highlightbackground=LINE, selectbackground=ACCENT)
        self.lista_datas.pack(fill="x", pady=(6, 0))
        self.lista_datas.bind("<<ListboxSelect>>", self._on_selecionar_data)

        # Scraper + executar
        bloco_acao = ttk.Frame(root, style="Card.TFrame", padding=12)
        bloco_acao.pack(fill="x", padx=16, pady=(0, 10))
        self.var_scraper = tk.BooleanVar(value=False)
        ttk.Checkbutton(bloco_acao, text="Rodar busca no ML/MP (precisa do Chrome aberto e logado)",
                         variable=self.var_scraper).pack(anchor="w")
        self.btn_executar = ttk.Button(bloco_acao, text="Executar conciliacao", style="Accent.TButton",
                                        command=self.executar, state="disabled")
        self.btn_executar.pack(anchor="w", pady=(8, 0))

        # Acoes secundarias (equivalentes aos .bat gerados)
        bloco_sec = ttk.Frame(root, style="Card.TFrame", padding=12)
        bloco_sec.pack(fill="x", padx=16, pady=(0, 10))
        ttk.Label(bloco_sec, text="Depois de rodar", style="Card.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        linha_sec = ttk.Frame(bloco_sec, style="Card.TFrame")
        linha_sec.pack(fill="x", pady=(6, 0))
        self.btn_atualizar = ttk.Button(linha_sec, text="Atualizar dashboard", state="disabled",
                                         command=self.atualizar)
        self.btn_atualizar.pack(side="left", padx=(0, 6))
        self.btn_arquivar = ttk.Button(linha_sec, text="Arquivar RBX", state="disabled",
                                        command=self.arquivar_rbx)
        self.btn_arquivar.pack(side="left", padx=(0, 6))
        self.btn_rebate = ttk.Button(linha_sec, text="Rebate fora da regra...", state="disabled",
                                      command=self.rebate_fora_da_regra)
        self.btn_rebate.pack(side="left", padx=(0, 6))

        bloco_cmm = ttk.Frame(root, style="Card.TFrame", padding=12)
        bloco_cmm.pack(fill="x", padx=16, pady=(0, 10))
        ttk.Label(bloco_cmm, text="Incluir no CMM (Excel real, sem copiar a mao)",
                  style="Card.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        linha_cmm = ttk.Frame(bloco_cmm, style="Card.TFrame")
        linha_cmm.pack(fill="x", pady=(6, 0))
        self.btn_rebate_cmm = ttk.Button(linha_cmm, text="Rebate automatico -> CMM", state="disabled",
                                          command=self.rebate_automatico_cmm)
        self.btn_rebate_cmm.pack(side="left", padx=(0, 6))
        self.btn_marcacoes_cmm = ttk.Button(linha_cmm, text="Marcacoes do dashboard (clipboard) -> CMM",
                                             state="disabled", command=self.marcacoes_dashboard_cmm)
        self.btn_marcacoes_cmm.pack(side="left")

        # Log (print() das funcoes reais)
        ttk.Label(root, text="Saida", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16)
        self.log = scrolledtext.ScrolledText(root, height=12, bg=SURFACE, fg=INK,
                                              font=("Consolas", 9), state="disabled")
        self.log.pack(fill="both", expand=True, padx=16, pady=(4, 16))

    # -- selecao -----------------------------------------------------------

    def selecionar_canal(self, canal):
        self.canal = canal
        self.data = None
        for nome, b in self.botoes_canal.items():
            b.state(["pressed"] if nome == canal else ["!pressed"])
        self.lista_datas.delete(0, "end")
        datas = interativo._listar_datas_disponiveis(canal)
        if not datas:
            self._log(f"Nenhum arquivo 'Para Processar' encontrado pra {canal} em {CANAIS[canal]['pasta']}.\n")
        for d in datas:
            self.lista_datas.insert("end", d)
        self._atualizar_estado_botoes()

    def _on_selecionar_data(self, _evt):
        sel = self.lista_datas.curselection()
        self.data = self.lista_datas.get(sel[0]) if sel else None
        self._atualizar_estado_botoes()

    def _atualizar_estado_botoes(self):
        pronto = "normal" if (self.canal and self.data) else "disabled"
        self.btn_executar.configure(state=pronto)
        self.btn_atualizar.configure(state=pronto)
        self.btn_arquivar.configure(state=pronto)
        self.btn_rebate.configure(state=pronto)
        self.btn_rebate_cmm.configure(state=pronto)
        self.btn_marcacoes_cmm.configure(state=pronto)

    # -- log + thread --------------------------------------------------------

    def _log(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _rodar_em_thread(self, fn, *args):
        fila = queue.Queue()

        def trabalho():
            with contextlib.redirect_stdout(_FilaParaLog(fila)):
                try:
                    fn(*args)
                except SystemExit as e:
                    print(f"\nErro: {e}")
                except Exception as e:  # prototipo: mostra e segue, nao trava a GUI
                    print(f"\nErro inesperado: {e}")
            fila.put(None)

        threading.Thread(target=trabalho, daemon=True).start()

        def drenar():
            try:
                while True:
                    item = fila.get_nowait()
                    if item is None:
                        return
                    self._log(item)
            except queue.Empty:
                self.root.after(100, drenar)

        drenar()

    # -- acoes ---------------------------------------------------------------

    def executar(self):
        self._log(f"\n=== Executando {self.canal} {self.data} (scraper={self.var_scraper.get()}) ===\n")
        usar_scraper = self.var_scraper.get()
        if usar_scraper and not interativo.scraper_disponivel():
            self._rodar_em_thread(
                lambda: interativo.relancar_com_scraper(self.canal, self.data, interativo.PASTA_REDE_PADRAO)
            )
        else:
            self._rodar_em_thread(interativo.executar, self.canal, self.data,
                                   interativo.PASTA_REDE_PADRAO, usar_scraper)

    def atualizar(self):
        self._log(f"\n=== Atualizando {self.canal} {self.data} (sem scraper) ===\n")
        self._rodar_em_thread(interativo.executar, self.canal, self.data,
                               interativo.PASTA_REDE_PADRAO, False)

    def arquivar_rbx(self):
        def tarefa():
            from rodar_conciliacao_rb import _arquivar_rbx
            caminho_rbx = os.path.join(interativo.PASTA_REDE_PADRAO, "AUXILIAR", "rbx.xlsx")
            destino = _arquivar_rbx(self.canal, self.data, CANAIS[self.canal]["raiz"], caminho_rbx)
            print(f"RBX arquivado: {destino}")
        self._log(f"\n=== Arquivando RBX de {self.canal} {self.data} ===\n")
        self._rodar_em_thread(tarefa)

    def rebate_fora_da_regra(self):
        caminho = filedialog.askopenfilename(
            title="Arquivo de resultado do TOTVS (marketplace...xls)",
            filetypes=[("Excel/HTML TOTVS", "*.xls *.xlsx *.html"), ("Todos", "*.*")],
        )
        if not caminho:
            return

        def tarefa():
            from parser_rb import metadata_rb
            from rebate_fora_da_regra import gerar_rebate_fora_da_regra_cmm
            data_ref, responsavel = metadata_rb(interativo.CAMINHO_RB_PADRAO)
            raiz = CANAIS[self.canal]["raiz"]
            df = gerar_rebate_fora_da_regra_cmm(caminho, raiz, data_ref, responsavel)
            saida = caminho.rsplit(".", 1)[0] + "_cmm.xlsx"
            df.to_excel(saida, index=False)
            print(f"{len(df)} pedido(s) fora da regra -> {saida}")

        self._log(f"\n=== Rebate fora da regra ({os.path.basename(caminho)}) ===\n")
        self._rodar_em_thread(tarefa)

    def _data_ref_do_rb(self):
        """DATA_LACTO do RB exportado do TOTVS (dd/mm/yyyy, ver
        parser_rb.metadata_rb) -- nao usa a data de hoje, usa a do RB (pode
        ser virada de ano/exportar atrasado). Usada tanto pro ano do CMM
        (CMM_<ano>\\...) quanto pro ddmmaaaa do nome do arquivo MOV."""
        from parser_rb import metadata_rb
        data_ref, _ = metadata_rb(interativo.CAMINHO_RB_PADRAO)
        return data_ref

    def _caminho_mov_ou_none(self, data_ref):
        """None quando o canal ainda nao tem pasta MOV confirmada (ver
        mov_writer.PASTA_POR_CANAL) -- nesse caso so grava no CMM, nao
        trava a acao toda por causa do MOV."""
        try:
            return mov_writer.caminho_mov(self.canal, data_ref.replace("/", ""))
        except ValueError as e:
            self._log(f"\nAviso: {e} (MOV nao sera gravado, so o CMM)\n")
            return None

    def rebate_automatico_cmm(self):
        from rodar_conciliacao_rb import processar
        try:
            r = processar(self.canal, self.data, interativo.PASTA_REDE_PADRAO, False)
        except Exception as e:
            self._log(f"\nErro ao reler a conciliacao: {e}\n")
            return
        df = r["rebate_cmm"]
        if not len(df):
            self._log(f"\nNenhum REBATE em {self.canal} {self.data} -- nada pra incluir.\n")
            return
        data_ref = r["data_ref"]
        ano = data_ref.split("/")[-1]
        caminho_cmm = caminho_cmm_real(self.canal, ano)
        if not os.path.exists(caminho_cmm):
            self._log(f"\nArquivo CMM nao encontrado: {caminho_cmm}\n"
                       "(mapa canal->arquivo em NOME_CMM_POR_CANAL nao confirmado -- confira o nome)\n")
            return
        caminho_mov = self._caminho_mov_ou_none(data_ref)
        destinos = f"CMM:\n{caminho_cmm}" + (f"\n\nMOV:\n{caminho_mov}" if caminho_mov else "")
        if not messagebox.askyesno("Confirma inclusao", f"Incluir {len(df)} linha(s) de rebate em:\n\n{destinos}"):
            return
        self._log(f"\n=== Rebate automatico: {len(df)} linha(s) -> CMM"
                   + (" + MOV" if caminho_mov else "") + " ===\n")
        self._rodar_em_thread(self._escrever_cmm_e_mov, caminho_cmm, df.to_dict("records"), caminho_mov)

    def marcacoes_dashboard_cmm(self):
        try:
            texto = self.root.clipboard_get()
        except tk.TclError:
            self._log("\nArea de transferencia vazia -- clique 'Copiar' na gaveta do dashboard primeiro.\n")
            return
        linhas = cmm_writer.linhas_do_clipboard_cmm(texto, COLUNAS_REBATE_CMM)
        if not linhas:
            self._log("\nNao achei linha de CMM no clipboard (clique 'Copiar CMM' na gaveta do dashboard).\n")
            return
        try:
            data_ref = self._data_ref_do_rb()
        except Exception as e:
            self._log(f"\nErro ao achar a data do RB: {e}\n")
            return
        ano = data_ref.split("/")[-1]
        caminho_cmm = caminho_cmm_real(self.canal, ano)
        if not os.path.exists(caminho_cmm):
            self._log(f"\nArquivo CMM nao encontrado: {caminho_cmm}\n")
            return
        caminho_mov = self._caminho_mov_ou_none(data_ref)
        destinos = f"CMM:\n{caminho_cmm}" + (f"\n\nMOV:\n{caminho_mov}" if caminho_mov else "")
        if not messagebox.askyesno(
            "Confirma inclusao",
            f"Incluir {len(linhas)} linha(s) marcada(s) no dashboard em:\n\n{destinos}\n\n1a linha: {linhas[0]}"
        ):
            return
        self._log(f"\n=== Marcacoes do dashboard: {len(linhas)} linha(s) -> CMM"
                   + (" + MOV" if caminho_mov else "") + " ===\n")
        self._rodar_em_thread(self._escrever_cmm_e_mov, caminho_cmm, linhas, caminho_mov)

    @staticmethod
    def _escrever_cmm_e_mov(caminho_cmm, linhas, caminho_mov):
        ultima, backup = cmm_writer.incluir_linhas_cmm(caminho_cmm, linhas)
        print(f"Incluido no CMM: {len(linhas)} linha(s), ate a linha {ultima}. Backup: {backup}")
        if not caminho_mov:
            return
        try:
            linhas_mov = mov_writer.linhas_cmm_para_mov(linhas)
            ultima_mov = mov_writer.incluir_linhas_mov(caminho_mov, linhas_mov)
            print(f"Incluido no MOV: {len(linhas)} linha(s), ate a linha {ultima_mov}. Arquivo: {caminho_mov}")
        except Exception as e:
            print(f"CMM ok, mas MOV falhou ({caminho_mov}): {e}")


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    janela = tk.Tk()
    App(janela)
    janela.mainloop()
