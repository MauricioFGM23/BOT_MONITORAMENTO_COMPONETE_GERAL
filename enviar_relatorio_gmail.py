import os
import shutil
import smtplib
from datetime import datetime
import papermill as pm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

# ======================================================
# CONFIGURAÇÕES GERAIS
# ======================================================
NOTEBOOK_PATH = 'credito_modalidade.ipynb'
SAIDA_DIR = 'saida'
ARQUIVO_NOME = f"{datetime.today().strftime('%Y%m%d')}_MONITORAMENTO DE COMPONENTE.xlsx"
RELATORIO_PATH = os.path.join(SAIDA_DIR, ARQUIVO_NOME)

DESTINO_PUBLICO = r"C:\Users\Datasus\OneDrive - Ministério da Saúde\Coordenação de Gestão da Informação - Documentos\BOT'S\PUBLICO"
DESTINO_FINAL = os.path.join(DESTINO_PUBLICO, ARQUIVO_NOME)

# ======================================================
# CONFIGURAÇÃO DO GMAIL
# ======================================================
EMAIL_REMETENTE = "seu.email@gmail.com"
SENHA_APP = "sua-senha-de-aplicativo"
DESTINATARIOS = [
    "otavio.santos@saude.gov.br",
]
ASSINATURA_IMG = os.path.abspath("img/assinatura.jpg")

# ======================================================
# FUNÇÕES
# ======================================================

def executar_notebook():
    print("🚀 Executando notebook...")
    try:
        pm.execute_notebook(NOTEBOOK_PATH, NOTEBOOK_PATH)
        print("✅ Notebook executado com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao executar notebook: {e}")

def copiar_para_publico():
    if os.path.exists(RELATORIO_PATH):
        try:
            shutil.copy(RELATORIO_PATH, DESTINO_FINAL)
            print(f"📁 Relatório copiado para pasta pública:\n{DESTINO_FINAL}")
        except PermissionError:
            print(f"⚠️ Arquivo em uso: {RELATORIO_PATH}")
    else:
        print(f"❌ Relatório não encontrado em: {RELATORIO_PATH}")

def limpar_arquivos_em_uso(pasta):
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                print(f"🗑️ Arquivo removido: {caminho_arquivo}")
            except PermissionError:
                print(f"⚠️ Arquivo em uso: {caminho_arquivo}")

def enviar_email():
    if not os.path.exists(RELATORIO_PATH):
        print(f"❌ Arquivo para envio não encontrado: {RELATORIO_PATH}")
        return

    print("📧 Preparando e-mail para envio via Gmail...")

    # Monta o e-mail
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

    # Adiciona imagem da assinatura
    if os.path.exists(ASSINATURA_IMG):
        with open(ASSINATURA_IMG, "rb") as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header("Content-ID", "<assinatura_img>")
            msg.attach(mime_img)

    # Anexa o relatório
    with open(RELATORIO_PATH, "rb") as f:
        parte = MIMEBase("application", "octet-stream")
        parte.set_payload(f.read())
    encoders.encode_base64(parte)
    parte.add_header(
        "Content-Disposition", f'attachment; filename="{os.path.basename(RELATORIO_PATH)}"'
    )
    msg.attach(parte)

    # Envia o e-mail via Gmail SMTP
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
    copiar_para_publico()
    enviar_email()
