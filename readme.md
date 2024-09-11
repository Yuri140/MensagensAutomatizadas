# Automatização de Envio de Mensagens via WhatsApp

Este projeto foi desenvolvido para automatizar o envio de mensagens personalizadas via WhatsApp, afim de automatizar o processo para a empresa W1, utilizando dados de uma planilha Excel. Ele é ideal para casos onde há necessidade de enviar mensagens em massa de forma eficiente e personalizada.

## Funcionalidades

- Leitura de dados de uma planilha Excel (.xlsx)
- Extração e formatação automática de nomes e números de telefone
- Envio automatizado de mensagens pelo WhatsApp
- Fechamento automático de guias abertas no WhatsApp Web após o envio da mensagem

## Tecnologias Utilizadas

- **Python**: Linguagem de programação principal.
- **Pandas**: Para leitura e manipulação dos dados da planilha.
- **PyWhatKit**: Biblioteca utilizada para enviar mensagens pelo WhatsApp.
- **PyAutoGUI**: Para automatizar ações no navegador, como fechar guias.

## Como Utilizar

1. Clone este repositório para sua máquina local.
   ```bash
   git clone https://github.com/Yuri140/MensagensAutomatizadas.git
   ```
2. Instale as dependências necessárias.
   ```bash
   pip install -r requirements.txt
   ```
3. Coloque sua planilha no diretório do projeto e ajuste o código para apontar para o caminho correto.
4. Execute o script.
   ```bash
   python Automatizando_Whatsapp.py
   ```

## Observações Importantes

- **Privacidade**: Certifique-se de que os dados utilizados estejam em conformidade com as leis de privacidade aplicáveis, como a LGPD no Brasil. **Não compartilhe dados pessoais em repositórios públicos.**
- **Customização**: Adapte a mensagem padrão e os detalhes da planilha conforme necessário para o seu caso de uso.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
