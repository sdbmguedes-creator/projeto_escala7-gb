from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
B = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(B)
for c in ["AL39","AL40","AL53"]:
    v = wb["2ªSB"][c].value
    t = v.text if isinstance(v, ArrayFormula) else v
    print(c, repr(t))
