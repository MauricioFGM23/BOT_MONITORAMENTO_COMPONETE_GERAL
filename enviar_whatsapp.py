import json
import logging
import os
import time
import urllib.parse
import pyautogui
import papermill as pm
from dotenv import load_dotenv
import ast
from datetime import datetime
import subprocess

# ==============================================================
# CONFIGURAÇÕES GERAIS
# ==============================================================
load_dotenv()
logging.basicConfig(level=logging.ERROR)

NOTEBOOK_PATH = 'credito_modalidade.ipynb'
SAIDA_DIR = 'saida'
METRICS_PATH = os.path.join(SAIDA_DIR, 'whatsapp_metrics.json')
CAMINHO_IMAGEM_BOTAO_ENVIAR = os.path.join('img', 'btn_enviar.png')

WHATSAPP_CONTATOS = ast.literal_eval(os.getenv("WHATSAPP_CONTATOS"))
SHAREPOINT_LINK = os.getenv("SHAREPOINT_LINK")


# ==============================================================
# FUNÇÕES AUXILIARES
# ==============================================================
def obter_saudacao():
    hora = datetime.now().hour
    if 5 <= hora < 12:
        return 'BOM DIA'
    elif 12 <= hora < 18:
        return 'BOA TARDE'
    else:
        return 'BOA NOITE'


def executar_notebook():
    print('🚀 1/3: Executando notebook...')
    try:
        pm.execute_notebook(NOTEBOOK_PATH, NOTEBOOK_PATH)
        print('✅ Notebook executado e métricas geradas.')
    except Exception as e:
        raise RuntimeError(f'Erro ao executar notebook: {e}')


def carregar_metricas():
    print('📊 2/3: Lendo métricas...')
    if not os.path.exists(METRICS_PATH):
        print('⚠️ Métricas não encontradas. Usando valores padrão.')
        return {
            'data': datetime.today().strftime('%d/%m/%Y'),
            'credito_financeiro': {'nome': 'Crédito Financeiro', 'status_propostas': {}, 'ufs_aprovadas_count': 'N/A', 'municipios_aprovados_count': 'N/A'},
            'modalidade_1': {'nome': 'Modalidade 1', 'status_propostas': {}, 'ufs_aprovadas_count': 'N/A', 'municipios_aprovados_count': 'N/A'}
        }
    with open(METRICS_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def criar_mensagem_detalhada(metricas, nome_contato):
    saudacao = obter_saudacao()

    def formatar_modalidade(mod):
        msg = f"*Modalidade: {mod['nome']}*\n"
        if mod['status_propostas']:
            for status, count in mod['status_propostas'].items():
                msg += f"-> {status}: {count} propostas\n"
        else:
            msg += "-> Nenhum dado disponível.\n"
        msg += f"📍 {mod['ufs_aprovadas_count']} UFs e {mod['municipios_aprovados_count']} municípios aprovados.\n"
        return msg

    return (
        f"{saudacao}, {nome_contato.upper()}!\n\n"
        f"Segue o Relatório Diário - {metricas['data']}.\n\n"
        
        "*RESUMO DE MONITORAMENTO POR MODALIDADE:*\n"
        "--------------------------------------\n"
        f"{formatar_modalidade(metricas['credito_financeiro'])}"
        "--------------------------------------\n"
        "--------------------------------------\n"
        f"📎 Acesso ao relatório completo:\n{SHAREPOINT_LINK}\n\n"
        "Atenciosamente,\nOtavio Augusto - BOT 🤖"
    )


# ==============================================================
# FUNÇÃO PRINCIPAL DE ENVIO (COM JANELA MAXIMIZADA)
# ==============================================================
def enviar_whatsapp_nao_interativo_automatico_visual():
    print('📢 3/3: ENVIANDO WHATSAPP via PyAutoGUI + Chrome (janela maximizada)...')

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.7

    metricas = carregar_metricas()

    for idx, contato in enumerate(WHATSAPP_CONTATOS, 1):
        nome = contato['nome']
        numero = contato['numero'].replace('+', '')
        mensagem_final = criar_mensagem_detalhada(metricas, nome)
        mensagem_codificada = urllib.parse.quote(mensagem_final)
        url = f'https://web.whatsapp.com/send?phone={numero}&text={mensagem_codificada}'

        print(f'\n📤 ({idx}/{len(WHATSAPP_CONTATOS)}) Enviando para {nome} ({numero})...')

        # ✅ Abre Chrome em nova janela **maximizada**
        cmd = f'powershell -Command "Start-Process chrome \'{url}\' -WindowStyle Maximized"'
        subprocess.Popen(cmd, shell=True)

        print('⏳ Aguardando carregamento do WhatsApp Web (12s)...')
        time.sleep(12)

        # 🔹 Garante foco e força renderização visual
        pyautogui.hotkey('alt', 'tab')
        time.sleep(1)

        screen_w, screen_h = pyautogui.size()
        pyautogui.moveTo(screen_w // 2, screen_h // 2, duration=0.5)
        pyautogui.moveRel(80, 0, duration=0.3)
        pyautogui.moveRel(-160, 0, duration=0.3)
        pyautogui.scroll(-400)
        time.sleep(1)

        # 🔹 Localiza botão "Enviar" por imagem
        print("🔎 Procurando o botão 'Enviar' (até 30s)...")
        send_center = None
        start_time = time.time()

        while time.time() - start_time < 30:
            try:
                send_center = (
                    pyautogui.locateCenterOnScreen(CAMINHO_IMAGEM_BOTAO_ENVIAR, confidence=0.9, grayscale=True)
                    or pyautogui.locateCenterOnScreen(CAMINHO_IMAGEM_BOTAO_ENVIAR, confidence=0.9, grayscale=True)
                )
                if send_center:
                    break
            except Exception as e:
                print(f'(debug locateOnScreen) erro: {e}')
            time.sleep(1)

        if send_center:
            x, y = send_center
            print(f'🟢 Botão encontrado em ({x}, {y}). Clicando...')
            pyautogui.moveTo(x, y, duration=0.3)
            pyautogui.click()
            print(f'✅ Mensagem enviada para {nome}.')
        else:
            print('⚠️ Botão não encontrado. Usando fallback (ENTER)...')
            pyautogui.press('enter')
            print(f'✅ Mensagem enviada para {nome} (via ENTER).')

        # 🔹 Fecha aba
        time.sleep(4)
        pyautogui.hotkey('alt', 'f4')
        print(f'🪟 Aba de {nome} fechada.')
        time.sleep(3)

    print('\n🎉 PROCESSO CONCLUÍDO!')


# ==============================================================
# EXECUÇÃO PRINCIPAL
# ==============================================================
if __name__ == '__main__':
    try:
        print("🤖 INICIANDO ORQUESTRAÇÃO DE ENVIO AUTOMÁTICO WHATSAPP")
        print("=" * 50)
        executar_notebook()
        enviar_whatsapp_nao_interativo_automatico_visual()
    except Exception as e:
        print(f"❌ PROCESSO INTERROMPIDO: {e}")
