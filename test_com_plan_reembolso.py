import os
import openpyxl
from openpyxl_funcs import *

planilha_reembolso = openpyxl.load_workbook(os.path.abspath('Lançamento Judicial Modelo_ Maio.xlsx'), data_only=True)
get_names_worksheets(planilha_reembolso, 1)

planilha_cobrados_total = planilha_reembolso['Cobrados Total']



len_colun = len_columns(planilha_cobrados_total)



ids = get_colun_data(plan=planilha_cobrados_total, column=0, convert_tuple=True, print_values=False)
numeros_de_processo = get_colun_data(plan=planilha_cobrados_total, column=1, convert_tuple=True, print_values=False)
parceiros = get_colun_data(plan=planilha_cobrados_total, column=2, convert_tuple=True, print_values=False)
pagamento_realizado = get_colun_data(plan=planilha_cobrados_total, column=3, convert_tuple=True, print_values=False)


print(ids)
print()
print(numeros_de_processo)
print()
print(parceiros)
print()
print(pagamento_realizado)
    
    