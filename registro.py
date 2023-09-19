"""
    criado em 19/09/2023
    autor: Eduardo Santos
    ultima modificacao: 19/09/2023
"""
import pymysql # Biblioteca para fazer interacao com BD
import openpyxl # Biblioteca para ler planilha xlsx


# Conexao com BD
banco = pymysql.connect(    
    host="localhost",
    user="root",
    passwd="",
    database="login"
)

# Definir Cursor
cursor = banco.cursor()

# Carregando nosso workbook (Imagine ele como um livro físico para editar as planilhas)

wb = openpyxl.load_workbook(filename='dados_login.xlsx') # Abrindo o Arquivo.xlsx
sheet = wb['dados_01'] # Selecionando o nome da planilha

# Array auxiliar que vai guardar os dados para serem registrados no BD
arr=[]

# Percorre a planilha
for x in range(2,32): # Percorre as linhas
    for y in range (2,6): # Percorre as colunas
        arr.append(sheet.cell(row=x,column=y).value)

    # Definir os valores que serão registrados no banco
    registro = """
        insert into usuario(nome,sobrenome,email,senha)
            values('{}','{}','{}','{}');
        """.format(arr[0],arr[1],arr[2],int(arr[3]))

    # Realiza o registro
    cursor.execute(registro)
    arr.clear()

#Fechar conexoes
cursor.close()    
banco.close()

'''
 Fim do codigo até 19/09/2023
 modificado por: Eduardo Santos
'''