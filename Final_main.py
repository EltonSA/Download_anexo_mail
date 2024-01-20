import customtkinter as tk
from tkinter import filedialog, messagebox
import imaplib
import email
import os
import re
import logging
from datetime import datetime

# Configuração do logging
logging.basicConfig(filename='log_downloader.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class EmailDownloaderBackend:
    def __init__(self):
        self.email = ''
        self.password = ''
        self.save_location = ''
        self.access_key = None
        self.attachment_extension = ''

    def set_credentials(self, email, password):
        self.email = email
        self.password = password

    def set_save_location(self, save_location):
        self.save_location = save_location

    def set_access_key(self, access_key):
        self.access_key = access_key

    def set_attachment_extension(self, extension):
        self.attachment_extension = extension.lower()

    def validate_access_key(self, key):
        return key == '1'  # Substitua 'sua_chave_de_acesso' pela chave desejada

    def download_attachments(self):
        # Verifica se a chave de acesso foi fornecida
        if self.access_key is None:
            messagebox.showerror("Erro de Acesso", "Por favor, forneça a chave de acesso antes de continuar.")
            return

        # Verifica se as credenciais, o local de salvamento e a extensão foram fornecidos
        if not self.email or not self.password or not self.save_location:
            messagebox.showerror("Ops!!!", "Por favor, forneça email, senha e local de salvamento.")
            return

        # Verifica a chave de acesso
        if not self.validate_access_key(self.access_key):
            messagebox.showerror("Erro de Acesso", "Chave de acesso inválida. Acesso negado.")
            return

        # Conecta-se ao servidor IMAP
        server = imaplib.IMAP4_SSL("imap.gmail.com")
        server.login(self.email, self.password)

        # Seleciona a caixa de entrada
        server.select(mailbox='inbox', readonly=True)

        # Obtém os IDs dos e-mails
        _, email_ids = server.search(None, 'ALL')
        email_ids = email_ids[0].split()

        # Mapeia os nomes dos meses em inglês para português
        meses_portugues = {
            'January': 'Janeiro',
            'February': 'Fevereiro',
            'March': 'Março',
            'April': 'Abril',
            'May': 'Maio',
            'June': 'Junho',
            'July': 'Julho',
            'August': 'Agosto',
            'September': 'Setembro',
            'October': 'Outubro',
            'November': 'Novembro',
            'December': 'Dezembro'
        }

        for email_id in email_ids:
            _, msg_data = server.fetch(email_id, '(RFC822)')
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            # Verifica se o e-mail contém anexo com a extensão desejada
            if self.email_contains_attachment(msg):
                sender = self.decode_sender(msg["From"])
                sender_folder = os.path.join(self.save_location, self.clean_folder_name(sender))

                # Adiciona pasta do ano e do mês
                today = datetime.today()
                year_folder = os.path.join(sender_folder, str(today.year))
                
                # Nome do mês em português
                month_folder_name = meses_portugues[today.strftime('%B')]
                month_folder = os.path.join(year_folder, month_folder_name)

                # Cria a estrutura de pastas se não existir
                for folder in [sender_folder, year_folder, month_folder]:
                    if not os.path.exists(folder):
                        os.makedirs(folder)

                self._download_attachments(msg, month_folder)

                # Log
                attachments = [part.get_filename() for part in msg.walk() if part.get('Content-Disposition') is not None]
                log_message = f"Anexo(s) baixado(s) do e-mail {sender}: {attachments} em: {month_folder}"
                logging.info(log_message)
                print(log_message)

        # Fecha a conexão com o servidor
        server.close()
        server.logout()

        # Limpa campos de e-mail e senha
        self.email = ''
        self.password = ''

        # Exibe mensagem de conclusão
        messagebox.showinfo("Mensagem", "Os arquivos foram baixados com sucesso!")

    def _download_attachments(self, msg, sender_folder):
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            filename = part.get_filename()
            if filename:
                filepath = os.path.join(sender_folder, filename)
                with open(filepath, 'wb') as file:
                    file.write(part.get_payload(decode=True))
                # Log
                log_message = f"Anexo baixado do e-mail {self.decode_sender(msg['From'])}: {filename} em: {sender_folder}"
                logging.info(log_message)
                print(log_message)

    def decode_sender(self, sender):
        decoded, encoding = email.header.decode_header(sender)[0]
        if isinstance(decoded, bytes):
            return decoded.decode(encoding or 'utf-8')
        else:
            return str(decoded)

    def clean_folder_name(self, name):
        return re.sub(r'[\/:*?"<>|]', '_', name)

    def email_contains_attachment(self, msg):
        if msg.is_multipart():
            for part in msg.walk():
                if part.get('Content-Disposition') is not None:
                    if self.attachment_extension == 'all' or (part.get_filename() and part.get_filename().lower().endswith(self.attachment_extension)):
                        return True
        return False


class EmailDownloaderFrontend:
    def __init__(self, master, backend):
        self.master = master
        self.backend = backend
        master.title("Download de E-mails")

        # Variáveis de controle
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.save_location_var = tk.StringVar()
        self.access_key_var = tk.StringVar()
        self.extension_var = tk.StringVar()

        # Configuração da interface
        tk.CTkLabel(master, text="Chave de Acesso:").grid(row=0, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.access_key_var, show="*").grid(row=0, column=1)

        tk.CTkLabel(master, text="Email:").grid(row=1, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.email_var).grid(row=1, column=1)

        tk.CTkLabel(master, text="Senha:").grid(row=2, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.password_var, show="*").grid(row=2, column=1)

        tk.CTkLabel(master, text="Local de Salvamento:").grid(row=3, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.save_location_var).grid(row=3, column=1)
        tk.CTkButton(master, text="Escolher Pasta", command=self.browse_folder).grid(row=3, column=2)

        tk.CTkLabel(master, text="Extensão do Anexo:").grid(row=4, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.extension_var).grid(row=4, column=1)

        tk.CTkButton(master, text="Baixar Anexos", command=self.download_attachments).grid(row=5, column=1, pady=10)

        tk.CTkButton(master, text="Baixar Todos os Anexos", command=self.download_all_attachments).grid(row=5, column=2, pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.save_location_var.set(folder_selected)

    def download_attachments(self):
        self.backend.set_access_key(self.access_key_var.get())
        self.backend.set_credentials(self.email_var.get(), self.password_var.get())
        self.backend.set_save_location(self.save_location_var.get())
        self.backend.set_attachment_extension(self.extension_var.get())
        self.backend.download_attachments()

    def download_all_attachments(self):
        self.backend.set_access_key(self.access_key_var.get())
        self.backend.set_credentials(self.email_var.get(), self.password_var.get())
        self.backend.set_save_location(self.save_location_var.get())
        self.backend.set_attachment_extension('all')
        self.backend.download_attachments()


if __name__ == "__main__":
    root = tk.CTk()
    backend = EmailDownloaderBackend()
    frontend = EmailDownloaderFrontend(root, backend)
    root.mainloop()
