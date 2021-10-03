"""
Desafio Python - Dev Aprender

Crie uma planilha com os seguintes dados:

Nome da planilha: Meus computadores.
Nome da página: Computadores.
Nome da coluna: eletrônica, memória ram e preço.

Dados:

Computador 1, 8gb Ram, R$ 2500
Computador 2, 16gb Ram, R$ 5500
Computador 3, 32gb Ram, R$ 2600
"""

# Importando biblioteca

import openpyxl

# Classe do programa

class AutomacaoExcel:

    # Construtor da classe

    def __init__(self):
        self.fileExcel = ""
        self.computerPage = []

    # Criando arquivo excel

    def criateFile(self):
        try:
            self.fileExcel = openpyxl.Workbook()
        except:
            print(f'Erro na criação do arquivo.')

    # Criar uma página no arquivo

    def criatePage(self):
        try:
            self.fileExcel.create_sheet('Computadores')
        except (ValueError, TypeError):
            print(f'Erro na criação de página.')

    # Selecionando a página

    def choosePage(self):
        try:
            self.computerPage = self.fileExcel['Computadores']
        except (ValueError, TypeError):
            print(f'Erro ao seleciona a página determinada.')

    # Adicionando dados em colunas e linhas

    def addData(self):
        try:
            self.computerPage.append(['Eletrônica', 'Memórira Ram', 'Preço'])
            self.computerPage.append(['Computador 1', '8gb Ram', 'R$ 2500'])
            self.computerPage.append(['Computador 2', '16gb Ram', 'R$ 5500'])
            self.computerPage.append(['Comptuador 3', '32gb Ram', 'R$ 2600'])
        except (ValueError, TypeError):
            print(f'Erro ao inserir dados nas colunas e linhas.')

    # Salvando arquivo

    def saveFile(self):
        try:
            self.fileExcel.save('Meus computadores.xlsx')
        except (ValueError, TypeError):
            print(f'Erro ao salvar o arquivo.')

# Instanciando a classe

Computador = AutomacaoExcel()

# Chamando métodos da classe

Computador.criateFile()
Computador.criatePage()
Computador.choosePage()
Computador.addData()
Computador.saveFile()
            
            







