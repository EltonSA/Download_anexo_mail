import customtkinter as tk
from tkinter import messagebox, Toplevel, IntVar

class ConfigWindow:
    def __init__(self, master, backend):
        self.master = master
        self.backend = backend
        self.config_window = Toplevel(master)
        self.config_window.title("Configurações")

        # Variáveis de controle
        self.access_key_var = tk.StringVar()
        self.server_address_var = tk.StringVar()
        self.sender_filter_var = tk.StringVar()
        self.want_details_var = IntVar()

        # Configuração da interface de configurações
        tk.CTkLabel(self.config_window, text="Chave de Acesso:").grid(row=0, column=0, sticky="e")
        tk.CTkEntry(self.config_window, textvariable=self.access_key_var, show="*").grid(row=0, column=1)

        tk.CTkLabel(self.config_window, text="Servidor de E-mail:").grid(row=1, column=0, sticky="e")
        tk.CTkEntry(self.config_window, textvariable=self.server_address_var).grid(row=1, column=1)

        tk.CTkLabel(self.config_window, text="Filtro de Remetente:").grid(row=2, column=0, sticky="e")
        tk.CTkEntry(self.config_window, textvariable=self.sender_filter_var).grid(row=2, column=1)

        tk.CTkCheckBox(self.config_window, text="Quero os Detalhes", variable=self.want_details_var).grid(row=3, column=1, pady=5)

        tk.CTkButton(self.config_window, text="Salvar Configurações", command=self.save_config).grid(row=4, column=1, pady=10)

    def save_config(self):
        access_key = self.access_key_var.get()
        server_address = self.server_address_var.get()
        sender_filter = self.sender_filter_var.get()
        want_details = bool(self.want_details_var.get())

        if access_key and server_address:
            self.backend.set_access_key(access_key)
            self.backend.set_server_address(server_address)
            self.backend.set_sender_filter(sender_filter)
            self.backend.set_want_details(want_details)
            self.backend.save_config()
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            self.config_window.destroy()
        else:
            messagebox.showerror("Erro", "Preencha todos os campos.")
