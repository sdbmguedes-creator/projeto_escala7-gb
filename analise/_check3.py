from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P); wbv = load_workbook(P, data_only=True)
def txt(c):
    v = wb[c.rsplit("!",1)[0]][c.rsplit("!",1)[1]].value
    return v.text if isinstance(v, ArrayFormula) else v
print("=== formulas diarias: quais codigos entram no peso 0,5 ===")
for sh,row in [("1ªSB",30),("2ªSB",31),("3ªSB",19),("4ªSB",32)]:
    ws=wb[sh]
    v=ws.cell(row,7).value
    t=v.text if isinstance(v,ArrayFormula) else str(v)
    import re
    m=re.findall(r"\{[^}]*\}", t)
    print(f" {sh} linha {row} (col G): {m[:4]}")
print()
print("=== colunas de saldo: codigos contados ===")
for sh,cells in [("1ªSB",["AL3","AM3"]),("2ªSB",["AK3","AL3"]),("3ªSB",["AK3","AL3"]),("4ªSB",["AK3","AL3"])]:
    ws=wb[sh]
    for c in cells:
        print(f" {sh}!{c}: {str(ws[c].value)[:150]}")
print()
print("=== celulas #N/A remanescentes ===")
for sh in ["1ªSB","ACO"]:
    ws=wbv[sh]
    for row in ws.iter_rows():
        for c in row:
            if c.value=="#N/A":
                f=wb[sh][c.coordinate].value
                ft=f.text if isinstance(f,ArrayFormula) else f
                print(f" {sh}!{c.coordinate}: valor=#N/A | formula={str(ft)[:110]}")
