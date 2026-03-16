# ui_utils.py
# Contém componentes de UI customizados para o aplicativo.

import customtkinter as ctk

class CustomDialog(ctk.CTkToplevel):
    """
    Cria uma janela de diálogo modal customizada com tamanho padronizado.
    Suporta os tipos: 'info', 'warning', 'error', e 'yesno'.
    """
    def __init__(self, title: str, message: str, dialog_type: str = "info"):
        super().__init__()

        self.dialog_type = dialog_type
        self.result = None  # Para armazenar o resultado de 'yesno'

        self._configure_window(title)
        self._create_widgets(message)

    def _configure_window(self, title):
        """Configura as propriedades da janela."""
        self.title(title)
        self.geometry("550x180")  # Altura reduzida
        self.resizable(False, False)
        
        # Faz a janela ser modal (bloqueia a janela principal)
        self.transient()
        self.grab_set()
        
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _create_widgets(self, message):
        """Cria os widgets (label e botões) dentro da janela."""
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        message_label = ctk.CTkLabel(
            main_frame,
            text=message,
            wraplength=500,  # Quebra de linha automática
            justify="left",
            font=("Arial", 14)
        )
        message_label.pack(expand=True, fill="both")

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=(0, 20))

        if self.dialog_type in ["info", "warning", "error"]:
            ok_button = ctk.CTkButton(button_frame, text="OK", width=100, command=self._on_ok)
            ok_button.pack()
        elif self.dialog_type == "yesno":
            yes_button = ctk.CTkButton(button_frame, text="Sim", width=100, command=self._on_yes)
            yes_button.pack(side="left", padx=10)
            no_button = ctk.CTkButton(button_frame, text="Não", width=100, command=self._on_no)
            no_button.pack(side="right", padx=10)

    def _on_ok(self):
        self.result = True
        self.destroy()

    def _on_yes(self):
        self.result = True
        self.destroy()

    def _on_no(self):
        self.result = False
        self.destroy()
        
    def _on_cancel(self):
        # Trata o fechamento da janela (pelo 'X') como "Não" ou "OK"
        self.result = False
        self.destroy()

    def show(self):
        """Mostra a janela e espera até que ela seja fechada."""
        self.wait_window()
        return self.result