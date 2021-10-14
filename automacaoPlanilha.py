"""
Desafio Python - Dev Aprender

Descrição: Crie uma planilha com as seguintes informaçãoes
abaixo:

Nome da planilha: Meus computadores.
Nome da página: Computadores.
Nome da coluna: eletrônica, memória ram e preço.

Dados:

Computador 1, 8gb Ram, R$ 2500
Computador 2, 16gb Ram, R$ 5500
Computador 3, 32gb Ram, R$ 2600

Autor (a): Alexandre Gonçalo

Data atual: 03/10/2021

"""

# Importando biblioteca

import openpyxl

# Classe do programa

class AutomacaoExcel:

    # Construtor da classe

    def __init__(self):
        self.file_excel = ""
        self.computer_page = []

    # Criando arquivo excel

    def criate_file(self):
        try:
            self.file_excel = openpyxl.Workbook()
        except:
            print(f'Erro na criação do arquivo.')

    # Criar uma página no arquivo

    def criate_page(self):
        try:
            self.file_excel.create_sheet('Computadores')
        except (ValueError, TypeError):
            print(f'Erro na criação de página.')

    # Selecionando a página

    def choose_page(self):
        try:
            self.computer_page = self.file_excel['Computadores']
        except (ValueError, TypeError):
            print(f'Erro ao seleciona a página determinada.')

    # Adicionando dados em colunas e linhas

    def add_data(self):
        try:
            self.computer_page.append(['Eletrônica', 'Memórira Ram', 'Preço'])
            self.computer_page.append(['Computador 1', '8gb Ram', 'R$ 2500'])
            self.computer_page.append(['Computador 2', '16gb Ram', 'R$ 5500'])
            self.computer_page.append(['Comptuador 3', '32gb Ram', 'R$ 2600'])
        except (ValueError, TypeError):
            print(f'Erro ao inserir dados nas colunas e linhas.')

    # Salvando arquivo

    def save_file(self):
        try:
            self.file_excel.save('Meus computadores.xlsx')
        except (ValueError, TypeError):
            print(f'Erro ao salvar o arquivo.')

# Instanciando a classe

computador = AutomacaoExcel()

# Chamando métodos da classe

computador.criate_file()
computador.criate_page()
computador.choose_page()
computador.add_data()
computador.save_file()
            
            







