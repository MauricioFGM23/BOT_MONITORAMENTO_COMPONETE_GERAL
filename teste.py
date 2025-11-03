import os
import json
import shutil
import smtplib
from datetime import datetime
import papermill as pm
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import webbrowser

# ======================================================
# CARREGAR VARIÁVEIS DO .env
# ======================================================
load_dotenv()

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_APP = os.getenv("SENHA_APP")
DESTINATARIOS = [email.strip() for email in os.getenv("DESTINATARIOS").split(",")]
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")
GOOGLE_PROJECT_ID = os.getenv("GOOGLE_PROJECT_ID")

# ======================================================
# CONFIGURAÇÕES GERAIS
# ======================================================
NOTEBOOK_PATH = 'credito_modalidade.ipynb'
SAIDA_DIR = 'saida'
ARQUIVO_NOME = f"{datetime.today().strftime('%Y%m%d')}_MONITORAMENTO DE COMPONENTE.xlsx"
RELATORIO_PATH = os.path.join(SAIDA_DIR, ARQUIVO_NOME)
ASSINATURA_IMG = os.path.abspath("img/assinatura.jpg")

# ======================================================
# FUNÇÕES
# ======================================================

def gerar_arquivo_credenciais():
    credenciais = {
        "installed": {
            "client_id": GOOGLE_CLIENT_ID,
            "project_id": GOOGLE_PROJECT_ID,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_secret": GOOGLE_CLIENT_SECRET,
            "redirect_uris": ["http://localhost:8080/"]
        }
    }
    with open("credentials.json", "w") as f:
        json.dump(credenciais, f)

def executar_notebook():
    print("🚀 Executando notebook...")
    try:
        pm.execute_notebook(NOTEBOOK_PATH, NOTEBOOK_PATH)
        print("✅ Notebook executado com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao executar notebook: {e}")

def limpar_arquivos_em_uso(pasta):
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                print(f"🗑️ Arquivo removido: {caminho_arquivo}")
            except PermissionError:
                print(f"⚠️ Arquivo em uso: {caminho_arquivo}")

def upload_para_google_drive(caminho_arquivo, nome_arquivo, pasta_id):
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    gerar_arquivo_credenciais()

    # Força o navegador Chrome para autenticação OAuth
    chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe %s'
    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))

    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=8080, open_browser=False)

    # Abre manualmente no Chrome
    auth_url, _ = flow.authorization_url(prompt='consent')
    webbrowser.get('chrome').open(auth_url)

    flow.fetch_token(authorization_response=input("Cole a URL de redirecionamento aqui: "))
    creds = flow.credentials

    with open('token.json', 'w') as token:
        token.write(creds.to_json())

    service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': nome_arquivo, 'parents': [pasta_id]}
    media = MediaFileUpload(caminho_arquivo, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"✅ Arquivo enviado para o Google Drive (ID: {file.get('id')})")

def enviar_email():
    if not os.path.exists(RELATORIO_PATH):
        print(f"❌ Arquivo para envio não encontrado: {RELATORIO_PATH}")
        return

    print("📧 Preparando e-mail para envio via Gmail...")

    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["Subject"] = f"Relatório Diário - {datetime.today().strftime('%d/%m/%Y')}"

    corpo_html = f"""
    <html>
      <body>
        <p>Prezados,</p>
        <p>Segue em anexo o relatório diário gerado automaticamente.<br>
        Temos novidades... agora com CNES nas propostas que tinham apenas CNPJ.</p>
        <p>Atenciosamente,<br>Otavio Augusto - BOT</p>
        <img src="cid:assinatura_img">
      </body>
    </html>
    """
    msg.attach(MIMEText(corpo_html, "html"))

    if os.path.exists(ASSINATURA_IMG):
        with open(ASSINATURA_IMG, "rb") as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header("Content-ID", "<assinatura_img>")
            msg.attach(mime_img)

    with open(RELATORIO_PATH, "rb") as f:
        parte = MIMEBase("application", "octet-stream")
        parte.set_payload(f.read())
    encoders.encode_base64(parte)
    parte.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(RELATORIO_PATH)}"')
    msg.attach(parte)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(EMAIL_REMETENTE, SENHA_APP)
            servidor.send_message(msg)
        print("📤 E-mail enviado com sucesso via Gmail.")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

# ======================================================
# EXECUÇÃO PRINCIPAL
# ======================================================

if __name__ == "__main__":
    limpar_arquivos_em_uso(r"C:\Users\Datasus\Downloads")
    executar_notebook()
    upload_para_google_drive(RELATORIO_PATH, ARQUIVO_NOME, GOOGLE_DRIVE_FOLDER_ID)
    enviar_email()
