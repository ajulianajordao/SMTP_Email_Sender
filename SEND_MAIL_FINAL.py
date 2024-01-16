import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Enviar E-mails")

        # Variável para armazenar a senha de forma segura
        self.smtp_senha_var = tk.StringVar()

        self.criar_widgets_autenticacao()

    def criar_widgets_autenticacao(self):
        autenticacao_frame = ttk.Frame(self.master, padding=(20, 10))
        autenticacao_frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(autenticacao_frame, text="E-mail (Outlook/Hotmail):").grid(column=0, row=0, pady=5, sticky=tk.W)
        self.smtp_usuario_entry = ttk.Entry(autenticacao_frame, width=40)
        self.smtp_usuario_entry.grid(column=1, row=0, pady=5, sticky=tk.W)

        ttk.Label(autenticacao_frame, text="Senha:").grid(column=0, row=1, pady=5, sticky=tk.W)
        self.smtp_senha_entry = ttk.Entry(autenticacao_frame, show="*", textvariable=self.smtp_senha_var, width=40)
        self.smtp_senha_entry.grid(column=1, row=1, pady=5, sticky=tk.W)

        self.mostrar_senha_button = ttk.Button(autenticacao_frame, text="👁", command=self.mostrar_senha, width=2)
        self.mostrar_senha_button.grid(column=2, row=1, pady=5, sticky=tk.W)

        ttk.Button(autenticacao_frame, text="OK", command=self.iniciar_sessao).grid(column=1, row=2, pady=10, sticky=(tk.W, tk.E))

    def criar_widgets_envio_email(self, smtp_usuario, smtp_senha, servidor_smtp):
        envio_frame = ttk.Frame(self.master, padding=(20, 10))
        envio_frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(envio_frame, text="Destinatários (separados por vírgula):").grid(column=0, row=0, pady=5, sticky=tk.W)
        destinatarios_entry = ttk.Entry(envio_frame, width=40)
        destinatarios_entry.grid(column=1, row=0, pady=5, sticky=tk.W)

        ttk.Label(envio_frame, text="Assunto:").grid(column=0, row=1, pady=5, sticky=tk.W)
        assunto_entry = ttk.Entry(envio_frame, width=40)
        assunto_entry.grid(column=1, row=1, pady=5, sticky=tk.W)

        ttk.Label(envio_frame, text="Corpo:").grid(column=0, row=2, pady=5, sticky=tk.W)
        corpo_entry = ttk.Entry(envio_frame, width=40)
        corpo_entry.grid(column=1, row=2, pady=5, sticky=tk.W)

        ttk.Label(envio_frame, text="Quantidade de e-mails:").grid(column=0, row=3, pady=5, sticky=tk.W)
        qtd_emails_entry = ttk.Entry(envio_frame, width=40)
        qtd_emails_entry.grid(column=1, row=3, pady=5, sticky=tk.W)

        ttk.Button(envio_frame, text="Enviar", command=lambda: self.iniciar_envio(smtp_usuario, smtp_senha, servidor_smtp, destinatarios_entry, assunto_entry, corpo_entry, qtd_emails_entry)).grid(column=0, row=4, pady=10, sticky=tk.W)

        ttk.Button(envio_frame, text="Cancelar", command=self.cancelar).grid(column=1, row=4, pady=10, sticky=tk.W)

        # Adiciona a barra de progresso
        self.progress_frame = ttk.Frame(envio_frame)
        self.progress_frame.grid(column=0, row=5, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))

        style = ttk.Style()
        style.configure("green.Horizontal.TProgressbar", troughcolor="white", background="green")
        self.progressbar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=300, mode='determinate', style="green.Horizontal.TProgressbar")
        self.progressbar.grid(column=0, row=0, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.porcentagem_label = ttk.Label(self.progress_frame, text="0%", style="Bold.TLabel")
        self.porcentagem_label.grid(column=1, row=0, pady=5, sticky=tk.W)

    def mostrar_senha(self):
        # Alterna entre mostrar e ocultar a senha
        if self.smtp_senha_entry.cget("show") == "":
            self.smtp_senha_entry.configure(show="*")
            self.mostrar_senha_button.configure(text="👁")
        else:
            self.smtp_senha_entry.configure(show="")
            self.mostrar_senha_button.configure(text="👁️‍🗨")

    def iniciar_sessao(self):
        smtp_usuario = self.smtp_usuario_entry.get()
        smtp_senha = self.smtp_senha_var.get()

        # Verifica se os campos obrigatórios não estão vazios
        if not smtp_usuario or not smtp_senha:
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return

        try:
            contexto = ssl.create_default_context()
            
            # Configurações específicas do servidor SMTP do Outlook/Hotmail
            servidor_smtp = smtplib.SMTP("smtp-mail.outlook.com", 587)
            servidor_smtp.starttls(context=contexto)
            
            # Use a senha da conta principal
            servidor_smtp.login(smtp_usuario, smtp_senha)

            # Se a autenticação for bem-sucedida, criar widgets para envio de e-mails
            self.criar_widgets_envio_email(smtp_usuario, smtp_senha, servidor_smtp)
        except smtplib.SMTPAuthenticationError as auth_error:
            erro_msg = f"Erro ao autenticar: {auth_error.smtp_error.decode()}"
            messagebox.showerror("Erro de Autenticação", erro_msg)
        except Exception as e:
            erro_msg = f"Erro desconhecido: {str(e)}"
            messagebox.showerror("Erro", erro_msg)

    def iniciar_envio(self, smtp_usuario, smtp_senha, servidor_smtp, destinatarios_entry, assunto_entry, corpo_entry, qtd_emails_entry):
        destinatarios = destinatarios_entry.get().split(",")
        assunto = assunto_entry.get()
        corpo = corpo_entry.get()
        qtd_emails = int(qtd_emails_entry.get())

        # Verifica se os campos obrigatórios não estão vazios
        if not destinatarios or not assunto or not corpo or not qtd_emails:
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return

        try:
            contexto = ssl.create_default_context()

            for i in range(qtd_emails):
                destinatario = destinatarios[i % len(destinatarios)].strip()
                mensagem = MIMEMultipart()
                mensagem['From'] = smtp_usuario
                mensagem['To'] = destinatario
                mensagem['Subject'] = assunto
                mensagem.attach(MIMEText(corpo, 'plain'))

                servidor_smtp.sendmail(smtp_usuario, destinatario, mensagem.as_string())

                # Atualiza a barra de progresso
                progresso_atual = (i + 1) / qtd_emails * 100
                self.progressbar['value'] = progresso_atual
                self.porcentagem_label["text"] = f"{int(progresso_atual)}%"
                self.master.update_idletasks()

            # Exibir mensagem quando o envio estiver completo
            messagebox.showinfo("Concluído", "Todos os e-mails foram enviados com sucesso!")

        except Exception as e:
            erro_msg = f"Erro ao enviar e-mails: {str(e)}"
            messagebox.showerror("Erro", erro_msg)

        finally:
            # Fechar a conexão SMTP após o envio
            servidor_smtp.quit()

    def cancelar(self):
        # Ação ao pressionar o botão Cancelar
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()