import pandas as pd
import pywhatkit as kit
import pyautogui
import keyboard
import time

# Função para extrair o primeiro nome e formatar com a primeira letra maiúscula
def formatar_nome_completo(nome_completo):
    primeiro_nome = nome_completo.split()[0].capitalize()
    return primeiro_nome

# Função para enviar a mensagem pelo WhatsApp
def enviar_mensagem(numero, nome):
    mensagem = f"""Bom dia, {nome}. Tudo bem?

Prazer em conhecê-lo. Meu nome é Yuri Carneiro e sou da W1 Consultoria Financeira, a primeira e maior consultoria financeira do Brasil. 📊

Atualmente, estamos em parceria com a Secretaria da Educação, e tenho conversado com alguns de seus colegas professores sobre temas importantes como organização financeira, aquisição de imóveis e veículos, planejamento de aposentadoria, gestão de milhas, dívidas, entre outros. 🏠🚗💼

Gostaria de convidá-lo para uma reunião por videoconferência 📹, onde faremos uma análise detalhada da sua saúde financeira e verificaremos como podemos auxiliá-lo. A reunião é gratuita e tem duração de aproximadamente 1h30 a 2h. Durante nossa conversa, traremos também algumas orientações iniciais para aprimorar sua organização financeira. 💡💰

Por gentileza, poderia me informar um dia e horário que lhe seja conveniente para essa reunião?

Fico à disposição e no aguardo de seu retorno!"""
    
    hora = int(time.strftime('%H')) + 1
    minuto = int(time.strftime('%M'))
    kit.sendwhatmsg(f"{numero}", mensagem, hora, minuto)
    time.sleep(15)
    pyautogui.hotkey('ctrl', 'w')  # Fecha a guia do WhatsApp Web
    time.sleep(2)

# Carregar a planilha ignorando a primeira linha
df = pd.read_excel(r'C:\Users\AMD\OneDrive\Documentos\RemuneracaoAtivos.xlsx', header=1)

# Imprime os nomes das colunas para verificar
print(df.columns)

# Agora, ajuste o código com base nos nomes das colunas identificados
for index, row in df.iterrows():
    status = row['STATUS LIGAÇÃO']  # Ajuste com o nome correto da coluna
    if status in ['Caixa Postal', 'Desligou instantaneo']:
        nome_cliente = formatar_nome_completo(row['NOME'])  # Ajuste com o nome correto da coluna
        numero_cliente = row['CONTATO']  # Ajuste com o nome correto da coluna

        # Verificar se o número não é nulo ou 'N/C'
        if pd.notna(numero_cliente) and numero_cliente != 'N/C':
            enviar_mensagem(numero_cliente, nome_cliente)

print("Mensagens enviadas com sucesso!")
