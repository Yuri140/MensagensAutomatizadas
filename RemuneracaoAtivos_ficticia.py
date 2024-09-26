import openpyxl

# Função para converter valores monetários para números
def converter_para_numero(valor):
    try:
        return float(valor.replace('.', '').replace(',', '.'))
    except ValueError:
        return None

# Carregar a planilha existente
caminho_planilha = 'RemuneracaoAtivos.xlsx'
workbook = openpyxl.load_workbook(caminho_planilha)

# Selecionar a aba ativa (ou uma específica se necessário)
sheet = workbook.active

# Dados fictícios com remuneração em formato de texto
dados_falsos = [
    ("João Silva", "Analista Administrativo", "Ministério da Saúde", "8.200,00", "+55 11 91234-5678", "Caixa Postal"),
    ("Maria Oliveira", "Técnico de Informática", "Secretaria de Educação", "4.500,00", "+55 21 97654-3210", "Desligou Instantâneo"),
    ("Thiago Santos", "Analista de Sistemas", "Tribunal de Contas", "8.400,00", "+55 61 98765-4321", "Não Atendeu"),
    ("Ana Costa", "Engenheira de Produção", "Ministério do Trabalho", "10.500,00", "+55 31 92345-6789", "Caixa Postal"),
    ("Pedro Almeida", "Gerente de TI", "Prefeitura de São Paulo", "12.300,00", "+55 41 97654-3210", "Desligou Instantâneo"),
    ("Luana Pereira", "Coordenadora Pedagógica", "Secretaria de Saúde", "7.800,00", "+55 51 98765-4321", "Não Atendeu"),
    ("Carlos Souza", "Consultor Jurídico", "Tribunal de Justiça", "9.000,00", "+55 71 91234-5678", "Caixa Postal"),
    ("Juliana Lima", "Analista Financeiro", "Banco Central", "8.700,00", "+55 81 92345-6789", "Desligou Instantâneo"),
    ("Ricardo Mendes", "Diretor de Recursos Humanos", "Instituto Federal", "11.200,00", "+55 91 97654-3210", "Não Atendeu"),
    ("Fernanda Silva", "Especialista em Marketing", "Ministério da Educação", "6.900,00", "+55 85 98765-4321", "Caixa Postal"),
    ("Gustavo Costa", "Engenheiro Elétrico", "Secretaria de Infraestrutura", "10.800,00", "+55 43 91234-5678", "Desligou Instantâneo"),
    ("Beatriz Rocha", "Médica", "Hospital das Clínicas", "15.000,00", "+55 27 92345-6789", "Não Atendeu"),
    ("Eduardo Martins", "Arquivista", "Arquivo Nacional", "5.600,00", "+55 62 97654-3210", "Caixa Postal"),
    ("Tatiane Silva", "Analista de Qualidade", "Agência Nacional de Saúde", "7.100,00", "+55 51 98765-4321", "Desligou Instantâneo"),
    ("Marcelo Pereira", "Professor Universitário", "Universidade Federal", "9.800,00", "+55 77 91234-5678", "Não Atendeu"),
    ("Camila Fernandes", "Assistente Social", "Secretaria de Assistência Social", "6.400,00", "+55 21 92345-6789", "Caixa Postal"),
    ("Roberto Lima", "Tecnólogo em Informática", "Agência de Tecnologia", "7.500,00", "+55 11 97654-3210", "Desligou Instantâneo"),
    ("Tatiane Souza", "Analista de Projetos", "Instituto de Pesquisa", "8.200,00", "+55 91 98765-4321", "Não Atendeu"),
    ("Renan Costa", "Chefe de Gabinete", "Câmara dos Deputados", "12.000,00", "+55 35 91234-5678", "Caixa Postal"),
    ("Mariana Alves", "Gestora de Projetos", "Fundação Nacional", "8.900,00", "+55 43 92345-6789", "Desligou Instantâneo"),
    ("Jorge Mendes", "Supervisor de Operações", "Agência Reguladora", "10.500,00", "+55 81 97654-3210", "Não Atendeu"),
    ("Priscila Martins", "Coordenadora de Eventos", "Secretaria de Cultura", "7.600,00", "+55 62 98765-4321", "Caixa Postal"),
    ("Lucas Silva", "Analista de Comunicação", "Ministério da Justiça", "6.800,00", "+55 91 91234-5678", "Desligou Instantâneo"),
    ("Patrícia Fernandes", "Consultora de Vendas", "Empresa de Consultoria", "7.900,00", "+55 43 92345-6789", "Não Atendeu"),
    ("Ricardo Barbosa", "Especialista em TI", "Secretaria de Tecnologia", "8.500,00", "+55 11 97654-3210", "Caixa Postal"),
    ("Sofia Lima", "Analista de Riscos", "Banco de Desenvolvimento", "9.200,00", "+55 31 98765-4321", "Desligou Instantâneo"),
    ("Bruno Costa", "Gerente de Projetos", "Instituto de Engenharia", "11.000,00", "+55 21 92345-6789", "Não Atendeu"),
    ("Juliana Santos", "Diretora de Marketing", "Universidade de São Paulo", "13.000,00", "+55 77 91234-5678", "Caixa Postal"),
    ("Felipe Alves", "Desenvolvedor de Software", "Empresa de Tecnologia", "9.500,00", "+55 35 97654-3210", "Desligou Instantâneo"),
    ("Amanda Pereira", "Engenheira de Segurança", "Instituto de Segurança", "10.200,00", "+55 62 98765-4321", "Não Atendeu")
]

# Inserir os dados na planilha, começando da linha 2 (assumindo que a linha 1 são os cabeçalhos)
for i, dados in enumerate(dados_falsos, start=3):
    for j, valor in enumerate(dados, start=1):
        if j == 4:  # Coluna de remuneração
            valor = converter_para_numero(valor)
        sheet.cell(row=i, column=j, value=valor)

# Salvar a planilha com as alterações
workbook.save(caminho_planilha)

print("Dados fictícios inseridos com sucesso!")
