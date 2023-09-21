import os 

pasta = './' # Torna a pasta atual o diretório de trabalho
arqs = os.listdir(pasta) # Lista os arquivos locais

planilhas = [] # Array auxiliar
format_visual = 50
# Verifica se há alguma planilha '.xlsx' no diretório local
for(arq) in arqs: 
    if os.path.splitext(arq)[1] == ".xlsx":
        planilhas.append(arq)

if len(planilhas) <=0:
    print('Nenhum arquivo ".xlsx" foi encontrado!')
else:
    print('-'*format_visual)
    print("As seguintes planilhas foram encontradas:")
    print('-'*format_visual)
    for i in planilhas:
        print(i +'\n')    
    print("-"*format_visual)    