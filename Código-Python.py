import imaplib
import email
import os
import smtplib
import datetime
from email import policy
from email.message import EmailMessage

# Configurações de e-mail
IMAP_SERVER = "imap.gmail.com"  # Alterar conforme o provedor
IMAP_PORT = 993
EMAIL_USER = "email_origem@gmail.com"
EMAIL_PASS = "abcd efgh ijkl mnop"  # É necessário gerar senha de app

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

DESTINATARIO = "email_de@gmail.com"  # Para quem será enviado o e-mail

# Criar pasta para salvar os arquivos XML
SAVE_FOLDER = "xml_files"
os.makedirs(SAVE_FOLDER, exist_ok=True)

def conectar_imap():
    """ Conecta ao servidor IMAP e faz login """
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    return mail

def obter_emails_ultimos_dias(mail, dias=2):
    """ Obtém os e-mails da caixa de entrada dos últimos 'dias' dias """
    mail.select("inbox")

    data = (datetime.date.today() - datetime.timedelta(days=dias)).strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(SINCE {data})')

    return messages[0].split()

def baixar_anexos(mail, email_ids):
    """ Baixa arquivos XML dos e-mails selecionados """
    arquivos_baixados = []

    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1], policy=policy.default)

                for part in msg.walk():
                    if part.get_content_disposition() == "attachment":
                        filename = part.get_filename()
                        if filename and filename.lower().endswith(".xml"):
                            filepath = os.path.join(SAVE_FOLDER, filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            arquivos_baixados.append(filepath)

    return arquivos_baixados

def enviar_email_com_anexos(arquivos):
    """ Envia um e-mail com os arquivos XML baixados """
    if not arquivos:
        print("Nenhum arquivo para enviar.")
        return

    msg = EmailMessage()
    msg["Subject"] = "Arquivos XML Recebidos"
    msg["From"] = EMAIL_USER
    msg["To"] = DESTINATARIO
    msg.set_content("Segue em anexo os arquivos XML recebidos nos últimos dois dias.")

    # Adicionar anexos
    for arquivo in arquivos:
        with open(arquivo, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="xml", filename=os.path.basename(arquivo))

    # Enviar e-mail
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)

    print("E-mail enviado com sucesso!")

def main():
    mail = conectar_imap()
    email_ids = obter_emails_ultimos_dias(mail, dias=2)
    arquivos = baixar_anexos(mail, email_ids)
    mail.logout()

    enviar_email_com_anexos(arquivos)

if __name__ == "__main__":
    main()
