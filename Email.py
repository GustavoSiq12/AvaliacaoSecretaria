from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib

def enviar_email(assunto, corpo, destinatario, remetente, senha, arquivo_anexo=None):
    # Configuração do servidor SMTP
    smtp = "smtp.gmail.com" 
    port = '587'  # Porta para TLS

    # Criação do e-mail
    message = MIMEMultipart()
    message["From"] = remetente
    message["To"] = destinatario
    message["Subject"] = assunto
    message.attach(MIMEText(corpo, "plain"))

    # Anexando arquivo, se fornecido
    if arquivo_anexo:
        try:
            with open(arquivo_anexo, "rb") as file:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={arquivo_anexo.split('/')[-1]}",
            )
            message.attach(part)
        except Exception as e:
            print(f"Erro ao anexar arquivo: {e}")

    try:
        # Conexão com o servidor SMTP
        server = smtplib.SMTP(smtp, port)
        server.starttls()  # Inicializa TLS
        server.login(remetente, senha)  # Login no servidor SMTP
        server.sendmail(remetente, destinatario, message.as_string())  # Envia o e-mail
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")
    finally:
        server.quit()  # Fecha a conexão com o servidor
