import customtkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
import os
import re
import logging
import email
import imaplib
import time

from config import ConfigWindow

# Configuração do logging
logging.basicConfig(filename='email_downloader.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class EmailDownloaderBackend:
    def __init__(self):
        self.email = ''
        self.password = ''
        self.save_location = ''
        self.access_key = None
        self.attachment_extension = ''
        self.server_address = ''
        self.sender_filter = ''
        self.config_file = 'config.txt'
        self.want_details = False
        self.load_config()

    def set_credentials(self, email, password):
        self.email = email
        self.password = password

    def set_save_location(self, save_location):
        self.save_location = save_location

    def set_access_key(self, access_key):
        self.access_key = access_key

    def set_attachment_extension(self, extension):
        self.attachment_extension = extension.lower()

    def set_server_address(self, server_address):
        self.server_address = server_address

    def set_sender_filter(self, sender_filter):
        self.sender_filter = sender_filter

    def set_want_details(self, want_details):
        self.want_details = want_details

    def validate_access_key(self, key):
        return key == '1'

    def download_attachments(self, start_date, end_date):
        if self.access_key is None:
            messagebox.showerror("Erro de Acesso", "Por favor, forneça a chave de acesso antes de continuar.")
            return

        if not self.email or not self.password or not self.save_location:
            messagebox.showerror("Ops!!!", "Por favor, forneça email, senha e local de salvamento.")
            return

        if not self.validate_access_key(self.access_key):
            messagebox.showerror("Erro de Acesso", "Chave de acesso inválida. Acesso negado.")
            return

        server = imaplib.IMAP4_SSL(self.server_address)
        server.login(self.email, self.password)
        server.select(mailbox='inbox', readonly=True)

        start_date = self.convert_to_imap_date_format(start_date)
        end_date = self.convert_to_imap_date_format(end_date)

        search_criteria = f'(SINCE "{start_date}" BEFORE "{end_date}")'
        _, email_ids = server.search(None, search_criteria)
        email_ids = email_ids[0].split()

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
            print(msg)

            if self.email_contains_attachment(msg):
                sender = self.decode_sender(msg["From"])
                sender_folder = os.path.join(self.save_location, self.clean_folder_name(sender))

                if self.email_contains_attachment(msg):
                    senderDate = self.decode_sender(msg["Date"])
                    senderDate = datetime.strptime(msg['Date'],"%a, %d %b %Y %H:%M:%S %z")

                    today = datetime.today()
                    year_folder = os.path.join(sender_folder, str(today.year))

                    month_folder_name = meses_portugues[today.strftime('%B')]
                    month_folder = os.path.join(year_folder, month_folder_name)

                    for folder in [sender_folder, year_folder, month_folder]:
                        if not os.path.exists(folder):
                            os.makedirs(folder)

                    details = {
                        'sender': sender,
                        'date': senderDate,
                        'subject': self.decode_sender(msg["Subject"]),
                        'body': msg.get_payload()
                    }

                    self.set_sender_filter(sender)
                    self.set_want_details(True)
                    self.want_details = True
                    self._download_attachments(msg, month_folder, details)

                    attachments = [part.get_filename() for part in msg.walk() if part.get('Content-Disposition') is not None]
                    log_message = f"Detalhes e anexo(s) baixado(s) do e-mail {sender}: {attachments} em: {month_folder}"
                    logging.info(log_message)
                    print(log_message)

        server.close()
        server.logout()

        self.email = ''
        self.password = ''

        messagebox.showinfo("Mensagem", "Os arquivos foram baixados com sucesso!")

    def _download_attachments(self, msg, sender_folder, details):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_number = int(time.time() * 1000) % 1000
        file_name = f"{timestamp}_{unique_number}_attachment.txt"
        file_path = os.path.join(sender_folder, file_name)

        if self.want_details:
            with open(file_path, 'w', encoding='utf-8') as details_file:
                details_file.write(f"Remetente: {details['sender']}\n")
                details_file.write(f"Data: {details['date'].strftime('%d/%m/%Y %H:%M:%S')}\n")
                details_file.write(f"Assunto: {details['subject']}\n")
                details_file.write("\nCorpo do e-mail:\n")
                details_file.write(str(details['body']))

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

                log_message = f"Anexo baixado do e-mail {details['sender']}: {filename} em: {sender_folder}"
                logging.info(log_message)
                print(log_message)


    def convert_to_imap_date_format(self, user_date):
        # Converte a data fornecida pelo usuário para o formato necessário para o filtro IMAP
        user_date = datetime.strptime(user_date, '%d/%m/%Y')
        imap_date_format = user_date.strftime('%d-%b-%Y')
        return imap_date_format

    def decode_sender(self, sender):
        decoded, encoding = email.header.decode_header(sender)[0]
        if isinstance(decoded, bytes):
            return decoded.decode(encoding or 'utf-8')
        else:
            return str(decoded)

    def clean_folder_name(self, name):
        invalid_chars = r'[\/:*?"<>|]'
        return re.sub(invalid_chars, '_', name)


    def email_contains_attachment(self, msg):
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            if self.attachment_extension == 'all' or (part.get_filename() and part.get_filename().lower().endswith(self.attachment_extension)):
                return True
        return False

    def save_config(self):
        # Salva as configurações em um arquivo de texto
        with open(self.config_file, 'w') as config_file:
            config_file.write(f"AccessKey={self.access_key}\n")
            config_file.write(f"ServerAddress={self.server_address}\n")

    def load_config(self):
        # Tenta carregar as configurações do arquivo, se existir
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as config_file:
                lines = config_file.readlines()
                for line in lines:
                    key, value = line.strip().split('=')
                    if key == 'AccessKey':
                        self.access_key = value
                    elif key == 'ServerAddress':
                        self.server_address = value

class EmailDownloaderFrontend:
    def __init__(self, master, backend):
        self.master = master
        self.backend = backend
        master.title("Download de E-mails")

        # Variáveis de controle
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.save_location_var = tk.StringVar()
        self.extension_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()

        # Configuração da interface
        tk.CTkLabel(master, text="Email:").grid(row=0, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.email_var).grid(row=0, column=1)

        tk.CTkLabel(master, text="Senha:").grid(row=1, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.password_var, show="*").grid(row=1, column=1)

        tk.CTkLabel(master, text="Extensão do Anexo:").grid(row=2, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.extension_var).grid(row=2, column=1)

        tk.CTkLabel(master, text="Data Inicial:").grid(row=3, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.start_date_var).grid(row=3, column=1)

        tk.CTkLabel(master, text="Data Final:").grid(row=4, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.end_date_var).grid(row=4, column=1)

        tk.CTkLabel(master, text="Local de Salvamento:").grid(row=5, column=0, sticky="e")
        tk.CTkEntry(master, textvariable=self.save_location_var).grid(row=5, column=1)
        tk.CTkButton(master, text="Escolher Pasta", command=self.browse_folder).grid(row=5, column=2)

        tk.CTkButton(master, text="Baixar Anexos Específicos", command=self.download_attachments).grid(row=6, column=1, pady=10)

        tk.CTkButton(master, text="Baixar Todos os Anexos", command=self.download_all_attachments).grid(row=6, column=2, pady=10)

        tk.CTkButton(master, text="Configurações", command=self.open_config_window).grid(row=6, column=0, pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.save_location_var.set(folder_selected)

    def download_attachments(self):
        self.backend.set_credentials(self.email_var.get(), self.password_var.get())
        self.backend.set_save_location(self.save_location_var.get())
        self.backend.set_attachment_extension(self.extension_var.get())
        self.backend.download_attachments(self.start_date_var.get(), self.end_date_var.get())

    def download_all_attachments(self):
        start_date = self.start_date_var.get()
        end_date = self.end_date_var.get()

        # Verifica se as datas estão preenchidas
        if not start_date or not end_date:
            messagebox.showerror("Erro", "Preencha as datas antes de baixar os anexos.")
            return

        self.backend.set_credentials(self.email_var.get(), self.password_var.get())
        self.backend.set_save_location(self.save_location_var.get())
        self.backend.set_attachment_extension('all')
        self.backend.download_attachments(start_date, end_date)


    def open_config_window(self):
        config_window = ConfigWindow(self.master, self.backend)

if __name__ == "__main__":
    root = tk.CTk()
    backend = EmailDownloaderBackend()
    frontend = EmailDownloaderFrontend(root, backend)
    root.mainloop()
