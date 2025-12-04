import os
import shutil
import json
from datetime import datetime
import papermill as pm
import win32com.client
from dotenv import load_dotenv

# ======================================================
# CARREGAR VARIÁVEIS DO .env
# ======================================================
load_dotenv()

NOTEBOOK_PATH = os.getenv("NOTEBOOK_PATH")
SAIDA_DIR = os.getenv("SAIDA_DIR")
DESTINO_PUBLICO = os.getenv("DESTINO_PUBLICO")
EMAIL_DESTINATARIOS = os.getenv("EMAIL_DESTINATARIOS")
ASSINATURA_IMG = os.path.abspath(os.getenv("ASSINATURA_IMG"))
SHAREPOINT_LINK = os.getenv("SHAREPOINT_LINK")

# ======================================================
# CONFIGURAÇÕES DE ARQUIVO
# ======================================================
ARQUIVO_NOME = f"{datetime.today().strftime('%Y%m%d')}_MONITORAMENTO DE COMPONENTE.xlsx"
RELATORIO_PATH = os.path.join(SAIDA_DIR, ARQUIVO_NOME)
DESTINO_FINAL = os.path.join(DESTINO_PUBLICO, ARQUIVO_NOME)
METRICAS_PATH = os.path.join(SAIDA_DIR, 'whatsapp_metrics.json')

# ======================================================
# FUNÇÕES
# ======================================================

def formatar_modalidade(modalidade):
    nome = modalidade['nome']
    status = modalidade['status_propostas']
    ufs = modalidade['ufs_aprovadas_count']
    municipios = modalidade['municipios_aprovados_count']
    estabelecimento = modalidade['entidade_aprovadas']

    texto = (
        f"<strong>{nome}</strong><br>"
        f"Aprovado: {status.get('Aprovado', 0)}<br>"
        f"PRE-Aprovado: {status.get('PRE-Aprovado', 0)}<br>"
        f"Em Diligência: {status.get('Em Diligência', 0)}<br>"
        f"Em Preenchimento: {status.get('Em Preenchimento', 0)}<br>"
        f"Aguardando Validação do Gestor: {status.get('Aguardando Validação do Gestor', 0)}<br>"
        f"Em Análise: {status.get('Em Análise', 0)}<br>"
        f"Enviado para Análise: {status.get('Enviado para Análise', 0)}<br>"
        f"UFs aprovadas: {ufs}<br>"
        f"Municípios aprovados: {municipios}<br>"
        f"Estabelecimento aprovados: {estabelecimento}<br>"
    )
    return texto

def executar_notebook():
    print('🚀 Executando notebook...')
    try:
        pm.execute_notebook(NOTEBOOK_PATH, NOTEBOOK_PATH)
        print('✅ Notebook executado.')
    except Exception as e:
        print(f'❌ Erro ao executar notebook: {e}')

def copiar_para_publico():
    if os.path.exists(RELATORIO_PATH):
        try:
            shutil.copy(RELATORIO_PATH, DESTINO_FINAL)
            print(f'📁 Relatório copiado para pasta pública:\n{DESTINO_FINAL}')
        except PermissionError:
            print(f'⚠️ Permissão negada ao copiar o arquivo. Verifique se ele está aberto: {RELATORIO_PATH}')
    else:
        print(f'❌ Relatório não encontrado em: {RELATORIO_PATH}')

def enviar_email():
    if not os.path.exists(RELATORIO_PATH):
        print(f'❌ Arquivo para envio não encontrado: {RELATORIO_PATH}')
        return

    try:
        with open(METRICAS_PATH, encoding='utf-8') as f:
            metricas = json.load(f)
    except Exception as e:
        print(f'❌ Erro ao carregar métricas: {e}')
        return

    resumo_html = (
        f"<p><strong>Resumo por modalidade:</strong></p>"
        f"<ul>"
        f"<li>{formatar_modalidade(metricas['credito_financeiro'])}</li><br>"
        f"<li>{formatar_modalidade(metricas['modalidade_1'])}</li>"
        f"</ul>"
        f"<p>📎 Para mais detalhes, acesse o relatório completo em anexo<br>"
    )

    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)

        email.To = EMAIL_DESTINATARIOS
        email.Subject = f"Relatório Diário - {datetime.today().strftime('%d/%m/%Y')}"
        email.HTMLBody = (
            '<p>Prezados,</p>'
            '<p>Segue em anexo o relatório diário atualizado com os dados mais recentes.</p>'
            '<p>NOVIDADES!: "Modalidade 1" agora com solicitações PRÉ APROVADAS (lista disponibilizada pelo sr. Marcus)'
            + resumo_html +
            '<p>Atenciosamente,<br>Otavio Augusto - BOT</p>'
            '<img src="cid:assinatura_img">'
        )

        email.Attachments.Add(os.path.abspath(RELATORIO_PATH))

        assinatura = email.Attachments.Add(ASSINATURA_IMG)
        assinatura.PropertyAccessor.SetProperty(
            'http://schemas.microsoft.com/mapi/proptag/0x3712001F',
            'assinatura_img',
        )

        email.Send()
        print('📤 E-mail enviado com sucesso com resumo humanizado.')
    except Exception as e:
        print(f'❌ Erro ao enviar e-mail: {e}')

def limpar_arquivos_em_uso(pasta):
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                print(f'🗑️ Arquivo removido: {caminho_arquivo}')
            except PermissionError:
                print(f'⚠️ Arquivo em uso, não foi possível excluir: {caminho_arquivo}')

# ======================================================
# EXECUÇÃO PRINCIPAL
# ======================================================

if __name__ == '__main__':
    limpar_arquivos_em_uso(r'C:\Users\Datasus\Downloads')
    executar_notebook()
    copiar_para_publico()
    enviar_email()
