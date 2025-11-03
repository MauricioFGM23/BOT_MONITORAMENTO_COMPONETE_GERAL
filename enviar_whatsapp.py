import json
import logging
import os
import time
import urllib.parse
import webbrowser
from datetime import datetime

import papermill as pm
import pyautogui

# Desativar logs desnecessários
logging.basicConfig(level=logging.ERROR)

# --- Caminhos principais ---
notebook_path = 'credito_modalidade.ipynb'
saida_dir = 'saida'
metrics_path = os.path.join(saida_dir, 'whatsapp_metrics.json')

# --- Contatos ---
WHATSAPP_CONTATOS = [
    {'nome': 'Otavio Augusto', 'numero': '+55...'},
    {'nome': 'Gabriella', 'numero': '+55...'},
]

SHAREPOINT_LINK = 'https://saudegov.sharepoint.com......'

# --- Caminho das imagens ---
CAMINHO_IMAGEM_BOTAO_ENVIAR = os.path.join('img', 'btn_enviar.png')

# ---------------- FUNÇÕES BASE ----------------
def obter_saudacao():
    hora = datetime.now().hour
    if 5 <= hora < 12:
        return 'BOM DIA'
    elif 12 <= hora < 18:
        return 'BOA TARDE'
    return 'BOA NOITE'


def executar_notebook():
    print('🚀 1/3: Executando notebook...')
    try:
        pm.execute_notebook(notebook_path, notebook_path)
        print('✅ Notebook executado e métricas geradas.')
    except Exception as e:
        raise RuntimeError(f'Erro ao executar notebook: {e}')


def carregar_metricas():
    print('📊 2/3: Lendo métricas...')
    if not os.path.exists(metrics_path):
        print('⚠️ Métricas não encontradas. Usando dados N/A.')
        return {
            'data': datetime.today().strftime('%d/%m/%Y'),
            'credito_financeiro': {
                'nome': 'Crédito Financeiro',
                'status_propostas': {},
                'ufs_aprovadas_count': 'N/A',
                'municipios_aprovados_count': 'N/A',
            },
            'modalidade_1': {
                'nome': 'Modalidade 1',
                'status_propostas': {},
                'ufs_aprovadas_count': 'N/A',
                'municipios_aprovados_count': 'N/A',
            },
        }
    with open(metrics_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def criar_mensagem_detalhada(metricas, nome_contato):
    saudacao = obter_saudacao()

    def formatar_modalidade(data):
        msg = f"  *Modalidade: {data['nome']}*\n"
        if data['status_propostas']:
            for status, count in data['status_propostas'].items():
                msg += f'  -> {status}: {count} propostas\n'
        else:
            msg += '  -> Status não disponíveis.\n'
        msg += f"  📍 {data['ufs_aprovadas_count']} UFs e {data['municipios_aprovados_count']} municípios com propostas aprovadas.\n"
        return msg

    msg = (
        f'{saudacao}, {nome_contato.upper()}!\n\n'
        f"Segue o Relatório Diário - {metricas['data']}.\n\n"
        f'**RESUMO DE MONITORAMENTO POR MODALIDADE**:\n'
        f'----------------------------------------------------\n'
        f"{formatar_modalidade(metricas['credito_financeiro'])}"
        f'----------------------------------------------------\n'
        f"{formatar_modalidade(metricas['modalidade_1'])}"
        f'----------------------------------------------------\n'
        f'📎 Acesso ao relatório completo:\n{SHAREPOINT_LINK}\n\n'
        'Atenciosamente,\nOtavio Augusto - BOT'
    )
    return msg


# ---------------- FUNÇÃO DE ENVIO ----------------
def enviar_whatsapp_nao_interativo_automatico_visual():
    print('📢 3/3: ENVIANDO WHATSAPP via PyAutoGUI (Reconhecimento Visual)...')

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 1.0

    metricas = carregar_metricas()

    for idx, contato in enumerate(WHATSAPP_CONTATOS, 1):
        nome = contato['nome']
        numero = contato['numero'].replace('+', '')
        mensagem_final = criar_mensagem_detalhada(metricas, nome)
        mensagem_codificada = urllib.parse.quote(mensagem_final)
        url = f'https://web.whatsapp.com/send?phone={numero}&text={mensagem_codificada}'

        print(
            f'\n📤 ({idx}/{len(WHATSAPP_CONTATOS)}) Enviando para {nome} ({numero})...'
        )
        webbrowser.open_new_tab(url)

        # Aguarda o carregamento da página
        print('⏳ Aguardando carregamento do WhatsApp Web...')
        time.sleep(15)

        # Aguarda o botão "Enviar" aparecer na tela (até 10 tentativas)
        print('🔎 Procurando o botão de envio...')
        send_btn = None
        for _ in range(10):
            send_btn = pyautogui.locateOnScreen(
                CAMINHO_IMAGEM_BOTAO_ENVIAR, confidence=0.6, grayscale=True
            )
            if send_btn:
                break
            time.sleep(2)

        if send_btn:
            pyautogui.click(send_btn)
            print(f'🚀 Mensagem enviada com sucesso para {nome}!')
        else:
            print(
                f'⚠️ Botão de envio não encontrado para {nome}. Verifique a imagem de referência.'
            )

        # Espera antes do próximo envio
        time.sleep(5)

        # Fecha aba atual para evitar múltiplas abertas
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(2)

    print('\n🎉 PROCESSO CONCLUÍDO COM SUCESSO!')


# ---------------- MAIN ----------------
if __name__ == '__main__':
    try:
        print('🤖 INICIANDO ORQUESTRAÇÃO DE ENVIO AUTOMÁTICO WHATSAPP')
        print('=' * 50)
        executar_notebook()
        enviar_whatsapp_nao_interativo_automatico_visual()
    except Exception as e:
        print(f'❌ PROCESSO INTERROMPIDO: {e}')
