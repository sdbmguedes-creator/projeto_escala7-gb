from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
wb = load_workbook(V)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
for sh,cells in [("1ªSB",["G30","G31","G32","G36"]),("4ªSB",["F32","F33","F34"])]:
    print(f"===== {sh} =====")
    for c in cells:
        print(f"{c}: {t(wb[sh][c].value)}")
        print()
