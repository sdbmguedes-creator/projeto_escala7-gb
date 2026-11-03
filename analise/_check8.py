from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P); wbv = load_workbook(P, data_only=True)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
print("=== militares ADM fora da linha 3 / do bloco ADM ===")
for sh,col,row in [("1ªSB","AL",58),("2ªSB","AK",42),("4ªSB","AK",40),("1ªSB","AL",39),("2ªSB","AK",3)]:
    f = t(wb[sh][f"{col}{row}"].value)
    print(f" {sh}!{col}{row} ({wbv[sh].cell(row,4).value}):")
    print(f"     valor={wbv[sh][f'{col}{row}'].value} | formula={str(f)[:120]}")
