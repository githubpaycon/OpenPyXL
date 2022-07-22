"""
https://openpyxl.readthedocs.io/en/stable/
pip install openpyxl
pipenv install openpyxl
"""
from openpyxl_funcs import get_colun_data, len_columns

import openpyxl
import os

# nome da planilha (abre a planilha)
pedidos = openpyxl.load_workbook(os.path.abspath('pedidos.xlsx'))
nome_planilhas = pedidos.sheetnames # saber quais são as planilhas que tem no arquivo excel

# print(nome_planilhas) # ['Página1']

# pegar a sheet da planilha pedidos
plan1 = pedidos['Página1']

# acessando a coluna e a linha
# print(plan1['b4'].value) # ver o que tem na coluna B na linha 4

# print(plan1['b']) # ver só o que tem na coluna B  # retorna uma tupla com vários valores

# for campo in plan1['b']: # iterando sobre para ver só o que tem na coluna B
#     print(campo.value) 
    
# acessando com um range() # digamos que queremos acessar da colua a linha1 até a coluna c linha2
# for linha in plan1['a1:c2']:
#     for coluna in linha:
#         print(coluna.value) 

# dar for direto na planilha:
# for lin in plan1:  # o lado negativo é que fica desorganizado
#     for col in lin:
#         print(col.value)


# pegar os dados de modo organizado
# for linha in plan1:
#     print(len(linha)) # pega quantas linhas tem em cada coluna
# existem 4 linhas -> [0, 1,2,3]


# ver de modo organizado # sem remover valores none
# for linha in plan1:
#     print(linha[0].value, # 0 é o mesmo que coluna a
#           linha[1].value, # 1 é o mesmo que coluna b
#           linha[2].value, # 2 é o mesmo que coluna c
#           linha[3].value) # 3 é o mesmo que coluna d

# pegar somente colunas com valores

# pega daddos somente de uma coluna
dados_coluna_pedido = []


print(get_colun_data(plan1, 1, 1))
len_columns(plan1, 1)

for linha in plan1:
    if linha[0].value is not None:
        print(linha[0].value, end=' ')
    if linha[1].value is not None:
        print(linha[1].value, end=' ')
    if linha[2].value is not None:
        print(linha[2].value)

