from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
V3 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
def t(v): return v.text if isinstance(v,ArrayFormula) else v
a, b = load_workbook(V2), load_workbook(V3)
ch={}
for sh in b.sheetnames:
    d=[c.coordinate for row in b[sh].iter_rows() for c in row if t(a[sh][c.coordinate].value)!=t(c.value)]
    if d: ch[sh]=len(d)
print("alteradas vs v2:", ch, "| total:", sum(ch.values()))
for n in b.sheetnames:
    assert len(a[n].data_validations.dataValidation)==len(b[n].data_validations.dataValidation)
    assert len(a[n].merged_cells.ranges)==len(b[n].merged_cells.ranges)
print("DV e merges: identicos")
print()
print("1aSB!AL3 :", t(b["1ªSB"]["AL3"].value))
print("4aSB!AK3 :", t(b["4ªSB"]["AK3"].value))
print("RESUMO!P1:", b["RESUMO"]["P1"].value)
print("ERALDIR P:", b["RESUMO"]["P9"].value, "| E:", b["RESUMO"]["E9"].value)
print("um nao-ADM P:", b["RESUMO"]["P3"].value)
