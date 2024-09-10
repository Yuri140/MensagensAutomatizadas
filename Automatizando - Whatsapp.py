import pywhatkit as kit
import keyboard
import time

def enviar_mensagem(numero, nome, primeira_execucao):
    mensagem = f"""Bom dia, {nome}. Tudo bem?

Prazer em conhecê-lo. Meu nome é Yuri Carneiro e sou da W1 Consultoria Financeira, a primeira e maior consultoria financeira do Brasil. 📊

Atualmente, estamos em parceria com a Secretaria da Educação, e tenho conversado com alguns de seus colegas professores sobre temas importantes como organização financeira, aquisição de imóveis e veículos, planejamento de aposentadoria, gestão de milhas, dívidas, entre outros. 🏠🚗💼

Gostaria de convidá-lo para uma reunião por videoconferência 📹, onde faremos uma análise detalhada da sua saúde financeira e verificaremos como podemos auxiliá-lo. A reunião é gratuita e tem duração de aproximadamente 1h30 a 2h. Durante nossa conversa, traremos também algumas orientações iniciais para aprimorar sua organização financeira. 💡💰

Por gentileza, poderia me informar um dia e horário que lhe seja conveniente para essa reunião?

Fico à disposição e no aguardo de seu retorno!"""

    # Tempo de espera diferenciado para a primeira execução
    if primeira_execucao:
        wait_time = 15  # Aumente o tempo para a primeira execução
    else:
        wait_time = 5  # Tempo reduzido para as execuções subsequentes

    # Enviar mensagem com tempo ajustado
    kit.sendwhatmsg_instantly(f"+{numero}", mensagem, wait_time=wait_time, tab_close=True)

    # Aguarde um pouco para garantir que a mensagem seja enviada
    time.sleep(wait_time)

    # Fecha a guia aberta (Ctrl + W)
    keyboard.press_and_release('ctrl+w')

# Loop para cadastro de mensagens
contatos = []
while True:
    numero_cliente = input("Digite o número do WhatsApp do cliente (com código do país e DDD): ")
    nome_cliente = input("Digite o nome do cliente: ")
    contatos.append((numero_cliente, nome_cliente))
    
    continuar = input("Deseja adicionar outro contato? (s/n): ")
    if continuar.lower() != 's':
        break

# Enviar mensagens para todos os contatos cadastrados
primeira_execucao = True
for numero_cliente, nome_cliente in contatos:
    enviar_mensagem(numero_cliente, nome_cliente, primeira_execucao)
    primeira_execucao = False  # Após a primeira execução, ajuste o tempo para as seguintes


#pip install pywhatkit
#pip install pyautogui
# pip install keyboard


