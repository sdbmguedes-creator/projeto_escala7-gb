import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V3 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
V4 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v4.xlsx"
a, b = load_workbook(V3), load_workbook(V4)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
tot=0; ch={}
for sh in b.sheetnames:
    d=[c.coordinate for row in b[sh].iter_rows() for c in row if t(a[sh][c.coordinate].value)!=t(c.value)]
    if d: ch[sh]=len(d); tot+=len(d)
print("alteradas vs v3:", ch, "| total:", tot)
restos=[]
for sh in b.sheetnames:
    for row in b[sh].iter_rows():
        for c in row:
            x=t(c.value)
            if isinstance(x,str) and re.search(r"RESUMO!?\$?F\$?\d*:\$?H\$?\d*", x, re.I):
                if "H$400" not in x: restos.append(f"{sh}!{c.coordinate}")
print("faixas antigas remanescentes:", len(restos), restos[:5])
for n in b.sheetnames:
    assert len(a[n].data_validations.dataValidation)==len(b[n].data_validations.dataValidation), n
    ca=sum(len(r) for r in a[n].conditional_formatting._cf_rules.values())
    cb=sum(len(r) for r in b[n].conditional_formatting._cf_rules.values())
    assert ca==cb, n
print("DV e CF: identicos")
print()
print("1aSB!B3 :", t(b["1ªSB"]["B3"].value))
print("RESUMO!I2:", t(b["RESUMO"]["I2"].value)[:190])
