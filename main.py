import customtkinter as ctk
from funcoes.navegador import acessar_eproc, fechar_navegador
from funcoes.eproc import acessar_localizadores, migrador, migrador_sem_cpf, registrar_cancelamento
import funcoes.eproc as eproc_module
from funcoes.logger import logger
import threading
import logging


class UILogHandler(logging.Handler):
    """Exibe mensagens importantes diretamente no painel da janela."""

    _FILTROS = [
        "--- [",
        "Total de processos encontrados",
        "Lendo arquivo:",
        "Iniciando:",
        "✅ Processo migrado",
        "✅ Processamento concluído",
        "❌ Erros de validação",
        "Código identificado:",
        "❌ TIMEOUT",
        "❌ ERRO:",
        "❌ Falha na recuperação",
        "⚠️ Execução cancelada",
        "⚠️ EXECUÇÃO INTERROMPIDA",
        "Aplicação iniciada",
    ]

    def __init__(self, textbox, app):
        super().__init__()
        self._tb = textbox
        self._app = app

    def emit(self, record):
        msg = record.getMessage()
        if not any(p in msg for p in self._FILTROS):
            return
        if "✅" in msg:
            tag = "verde"
        elif "❌" in msg:
            tag = "vermelho"
        elif "⚠️" in msg:
            tag = "laranja"
        elif "--- [" in msg:
            tag = "azul"
        else:
            tag = "normal"
        self._app.after(0, self._inserir, msg, tag)

    def _inserir(self, msg, tag):
        try:
            raw = self._tb._textbox
            raw.configure(state="normal")
            raw.insert("end", msg + "\n", (tag,))
            raw.see("end")
            raw.configure(state="disabled")
        except Exception:
            pass

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

MODOS_MIGRADOR = {
    "Parte COM CPF": migrador,
    "Parte SEM CPF": migrador_sem_cpf,
}

TIMEOUT_OPTIONS = {
    "3 min": 180,
    "5 min": 300,
    "9 min": 540,
}

FUNCOES = {
    "Acessar Localizadores": {
        "descricao": "Abre a seção 'Meus Localizadores' no e-Proc.",
    },
    "Migrador": {
        "descricao": "Migra processos em lote a partir de uma planilha Excel.",
    }
}


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automação TJSP")
        self.geometry("900x680")
        self.resizable(False, False)

        # Estado de execução
        self.is_running = False
        self.worker_thread = None
        self.cancel_event = None

        acessar_eproc()

        # ── Sidebar ───────────────────────────────────────────────────────────
        self.sidebar = ctk.CTkFrame(self, width=250, fg_color="#3B4A5A", corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(
            self.sidebar, text="Minhas Funções",
            text_color="white", font=("Arial", 22, "bold")
        ).pack(pady=(30, 20))

        self.color_btn_normal = "#2980B9"
        self.color_btn_hover = "#3498DB"
        self.color_btn_selecionado = "#5DADE2"

        self.sidebar_buttons = {}
        button_width = 220
        button_height = 40
        button_padx = (250 - button_width) // 2

        btn_acessar = ctk.CTkButton(
            self.sidebar, text="📂   Acessar Localizadores",
            width=button_width, height=button_height,
            fg_color=self.color_btn_normal, hover_color=self.color_btn_hover,
            corner_radius=10, font=("Arial", 14, "bold"), text_color="white", anchor="w",
            command=lambda: self.selecionar_funcao("Acessar Localizadores")
        )
        btn_acessar.pack(pady=(0, 8), padx=button_padx)
        self.sidebar_buttons["Acessar Localizadores"] = btn_acessar

        btn_migrador = ctk.CTkButton(
            self.sidebar, text="🔄   Migrador",
            width=button_width, height=button_height,
            fg_color=self.color_btn_normal, hover_color=self.color_btn_hover,
            corner_radius=10, font=("Arial", 14, "bold"), text_color="white", anchor="w",
            command=lambda: self.selecionar_funcao("Migrador")
        )
        btn_migrador.pack(pady=(0, 16), padx=button_padx)
        self.sidebar_buttons["Migrador"] = btn_migrador

        # Modo do migrador
        ctk.CTkLabel(
            self.sidebar, text="Modo:", text_color="#AABBCC", font=("Arial", 11)
        ).pack(pady=(0, 2))
        self.modo_var = ctk.StringVar(value=list(MODOS_MIGRADOR.keys())[0])
        ctk.CTkOptionMenu(
            self.sidebar,
            values=list(MODOS_MIGRADOR.keys()),
            variable=self.modo_var,
            width=button_width,
        ).pack(padx=button_padx, pady=(0, 10))

        # Timeout
        ctk.CTkLabel(
            self.sidebar, text="Timeout por processo:", text_color="#AABBCC", font=("Arial", 11)
        ).pack(pady=(0, 2))
        self.timeout_var = ctk.StringVar(value="3 min")
        ctk.CTkOptionMenu(
            self.sidebar,
            values=list(TIMEOUT_OPTIONS.keys()),
            variable=self.timeout_var,
            width=button_width,
            command=self._aplicar_timeout,
        ).pack(padx=button_padx)

        # Botão Sair
        ctk.CTkButton(
            self.sidebar, text="🚪  Sair",
            width=button_width, height=button_height,
            fg_color="#C0392B", hover_color="#E74C3C", corner_radius=10,
            font=("Arial", 14, "bold"), text_color="white",
            command=self.on_close, anchor="center"
        ).pack(side="bottom", pady=30, padx=button_padx)

        # ── Área principal ────────────────────────────────────────────────────
        self.main_area = ctk.CTkFrame(self, fg_color="#ECECEC")
        self.main_area.pack(side="left", fill="both", expand=True)
        self.main_area.pack_propagate(False)

        self.label_titulo = ctk.CTkLabel(
            self.main_area, text="", font=("Arial", 26, "bold"), text_color="#222"
        )
        self.label_titulo.pack(pady=(22, 4))

        self.label_desc = ctk.CTkLabel(
            self.main_area, text="", font=("Arial", 14), text_color="#444"
        )
        self.label_desc.pack(pady=(0, 14))

        self.btn_executar = ctk.CTkButton(
            self.main_area, text="Executar", width=200, height=46,
            fg_color="#2980B9", hover_color="#3498DB", corner_radius=10,
            font=("Arial", 17, "bold"), text_color="white",
            command=self.executar_funcao
        )
        self.btn_executar.pack(pady=(0, 6))

        self.btn_parar = ctk.CTkButton(
            self.main_area, text="Parar", width=200, height=36,
            fg_color="#C0392B", hover_color="#E74C3C", corner_radius=10,
            font=("Arial", 14, "bold"), text_color="white",
            command=self.parar_execucao, state="disabled"
        )
        self.btn_parar.pack(pady=(0, 12))

        # ── Painel de andamento ───────────────────────────────────────────────
        ctk.CTkLabel(
            self.main_area, text="Andamento:", font=("Arial", 12, "bold"),
            text_color="#555", anchor="w"
        ).pack(fill="x", padx=16, pady=(0, 3))

        self.log_box = ctk.CTkTextbox(
            self.main_area,
            fg_color="#1C2833",
            text_color="#D5D8DC",
            font=("Courier New", 11),
            state="disabled",
            wrap="word",
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        # Configurar cores por tipo de mensagem
        _raw = self.log_box._textbox
        _raw.tag_configure("verde",    foreground="#58D68D")
        _raw.tag_configure("vermelho", foreground="#EC7063")
        _raw.tag_configure("laranja",  foreground="#F0B27A")
        _raw.tag_configure("azul",     foreground="#85C1E9", font=("Courier New", 11, "bold"))
        _raw.tag_configure("normal",   foreground="#D5D8DC")

        # Registrar handler para exibir logs no painel
        self._ui_handler = UILogHandler(self.log_box, self)
        logger.addHandler(self._ui_handler)

        # ─────────────────────────────────────────────────────────────────────
        self.funcao_selecionada = None
        self.selecionar_funcao("Acessar Localizadores")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._aplicar_timeout("3 min")
        logger.info(f"Aplicação iniciada. Log salvo em: logs\\migracao.log")

    def _aplicar_timeout(self, _valor: str = None):
        segundos = TIMEOUT_OPTIONS.get(self.timeout_var.get(), 180)
        eproc_module.TIMEOUT_MIGRACAO = segundos
        logger.info(f"Timeout configurado para {self.timeout_var.get()} ({segundos}s)")

    def selecionar_funcao(self, funcao):
        self.funcao_selecionada = funcao
        self.label_titulo.configure(text=funcao)
        self.label_desc.configure(text=FUNCOES[funcao]["descricao"])
        self.btn_executar.configure(state="normal" if not self.is_running else "disabled")

        for nome, button in self.sidebar_buttons.items():
            cor = self.color_btn_selecionado if nome == funcao else self.color_btn_normal
            button.configure(fg_color=cor)

    def executar_funcao(self):
        if not self.funcao_selecionada or self.is_running:
            return

        if self.funcao_selecionada == "Migrador":
            self._aplicar_timeout()
            try:
                self.iconify()
            except Exception:
                pass
            acao = MODOS_MIGRADOR[self.modo_var.get()]
        else:
            acao = acessar_localizadores

        self.is_running = True
        self.cancel_event = threading.Event()
        try:
            registrar_cancelamento(self.cancel_event)
        except Exception:
            pass
        logger.info(f"Iniciando: {self.funcao_selecionada}")
        self.btn_executar.configure(state="disabled", text="Executando...")
        self.btn_parar.configure(state="normal")

        def _run():
            try:
                acao()
            finally:
                self.after(0, self._finalizar_execucao)

        self.worker_thread = threading.Thread(target=_run, daemon=True)
        self.worker_thread.start()

    def parar_execucao(self):
        if self.is_running and self.cancel_event is not None:
            self.cancel_event.set()
            self.btn_parar.configure(state="disabled")
            try:
                self.btn_executar.configure(text="Cancelando...")
            except Exception:
                pass

    def _finalizar_execucao(self):
        self.is_running = False
        self.btn_executar.configure(state="normal", text="Executar")
        self.btn_parar.configure(state="disabled")

    def on_close(self):
        if self.is_running and self.cancel_event is not None:
            try:
                self.cancel_event.set()
            except Exception:
                pass
        fechar_navegador()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
