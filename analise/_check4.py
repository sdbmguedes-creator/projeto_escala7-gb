from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P); wbv = load_workbook(P, data_only=True)
for sh,cells in [("ACO",["A54","B54","D54","A53","B53"]),("1ªSB",["A92","A91","A55"])]:
    ws=wb[sh]; wsv=wbv[sh]
    for c in cells:
        v=ws[c].value
        t=v.text if isinstance(v,ArrayFormula) else v
        print(f"{sh}!{c}: valor={wsv[c].value!r} | formula={str(t)[:120]}")
print()
print("4ªSB AK3:", wb["4ªSB"]["AK3"].value)
print("1ªSB AL3 (bloco proprio):", wb["1ªSB"]["AL3"].value)
print("1ªSB AL39 (bloco ADM):", wb["1ªSB"]["AL39"].value)
