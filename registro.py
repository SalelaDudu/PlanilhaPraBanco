import pymysql # Biblioteca para fazer interacao com BD
import openpyxl # Biblioteca para ler planilha xlsx



# Conexao com BD
banco = pymysql.connect(    
    host="localhost",
    user="root",
    passwd="root",
    database="login"
)

# Definir Cursor
cursor = banco.cursor()

# Carregando nosso workbook (Imagine ele como um livro físico para editar as planilhas)

wb = openpyxl.load_workbook(filename='dados_login.xlsx') # Abrindo o Arquivo.xlsx
sheet = wb['dados_01'] # Selecionando o nome da planilha

arr=[]

for x in range(2,5):
    for y in range(1,6):
        arr.append(sheet.cell(row=x,column=y).value)

print(arr)
# # Definir que sera consulta  feita
# consulta = """
#   select actor_id,first_name, last_name from actor where actor_id in 
#   (select actor_id from film_actor where film_id in (select film_id from film where title = "TRAP GUYS"));
# """

# # Realiza a consulta
# cursor.execute(consulta)

# #Fechar conexoes
cursor.close()    
banco.close()

'''
 Fim do codigo até 18/09/2023
 modificado por : Eduardo Santos
'''