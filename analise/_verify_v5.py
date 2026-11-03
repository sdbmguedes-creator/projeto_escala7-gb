import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V4 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v4.xlsx"
V5 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
a, b = load_workbook(V4), load_workbook(V5)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
ch={}
for sh in b.sheetnames:
    d=[c.coordinate for row in b[sh].iter_rows() for c in row if t(a[sh][c.coordinate].value)!=t(c.value)]
    if d: ch[sh]=len(d)
print("alteradas vs v4:", ch, "| total:", sum(ch.values()))
for n in b.sheetnames:
    assert len(a[n].data_validations.dataValidation)==len(b[n].data_validations.dataValidation)
    assert len(a[n].merged_cells.ranges)==len(b[n].merged_cells.ranges)
    ca=sum(len(r) for r in a[n].conditional_formatting._cf_rules.values())
    cb=sum(len(r) for r in b[n].conditional_formatting._cf_rules.values())
    assert ca==cb
print("DV, merges e CF: identicos")
print()
print("=== varredura: alguma faixa de COUNTIF fora do padrao do bloco? ===")
esp = {"1ªSB":"G:AK","2ªSB":"F:AJ","3ªSB":"F:AJ","4ªSB":"F:AJ"}
falhas=[]
for sh,colidx in [("1ªSB",38),("2ªSB",37),("2ªSB",38),("3ªSB",37),("3ªSB",38),("4ªSB",37),("4ªSB",38)]:
    ws=b[sh]
    for r in range(1, ws.max_row+1):
        x=t(ws.cell(r,colidx).value)
        if not isinstance(x,str): continue
        for m in re.finditer(r"COUNTIF\(\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+", x):
            if f"{m.group(1)}:{m.group(2)}" != esp[sh]:
                falhas.append((sh, ws.cell(r,colidx).coordinate, f"{m.group(1)}:{m.group(2)}"))
print("faixas fora do padrao:", len(falhas))
for f in falhas[:10]: print("   ", f)
print()
print("RESUMO!I2:", t(b["RESUMO"]["I2"].value)[-60:])
