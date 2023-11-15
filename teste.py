import openpyxl

caminho_arquivo = r'C:/Users/felip/Desktop/Mascaras (1)/Fios_e_cabos_NBR_16612_E_50618.xlsm'

workbook = openpyxl.load_workbook(caminho_arquivo)

sheet = workbook['Grau de acidez']

sheet.protection.password = 'sisten21'

celula = sheet['C11']

print(celula.value)
