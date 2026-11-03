from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
A = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"
B = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
O = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
def t(v): return v.text if isinstance(v,ArrayFormula) else v
wa, wb = load_workbook(A), load_workbook(B)
ch={}
for sh in wb.sheetnames:
    d=[]
    for row in wb[sh].iter_rows():
        for c in row:
            if t(wa[sh][c.coordinate].value)!=t(c.value): d.append(c.coordinate)
    if d: ch[sh]=len(d)
print("alteradas vs Bloco A:", ch, "| total:", sum(ch.values()))
# integridade de recursos
for n in wb.sheetnames:
    a,b = wa[n], wb[n]
    assert len(a.data_validations.dataValidation)==len(b.data_validations.dataValidation)
    ca=sum(len(r) for r in a.conditional_formatting._cf_rules.values())
    cb=sum(len(r) for r in b.conditional_formatting._cf_rules.values())
    assert ca==cb, n
    assert len(a.merged_cells.ranges)==len(b.merged_cells.ranges), n
print("DV / CF / merges: identicos em todas as abas")
# residuos
diario=sum(1 for sh in ["1ªSB","2ªSB","3ªSB","4ªSB"] for row in wb[sh].iter_rows() for c in row
           if isinstance(t(c.value),str) and '"PJ12";"S1"' in t(c.value))
ref=sum(1 for sh in wb.sheetnames for row in wb[sh].iter_rows() for c in row
        if isinstance(t(c.value),str) and "VOLUNTÁRIOS PJES" in t(c.value))
print("metas diarias sem PJ12H:", diario, "| refs a aba extinta:", ref)
print("A55:", t(wb["1ªSB"]["A55"].value))
print("orfas B29/C29/A54/B54:", [wb[s][c].value for s,c in [("1ªSB","B29"),("1ªSB","C29"),("ACO","A54"),("ACO","B54")]])
