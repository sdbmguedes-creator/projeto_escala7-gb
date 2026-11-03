from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
wb = load_workbook(F); wv = load_workbook(F, data_only=True)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
print("abas:", wb.sheetnames)
print()
for c in ["I2","J2","K2"]:
    print(f"RESUMO!{c}: {str(t(wb['RESUMO'][c].value))[:200]}")
print()
for sh,c in [("1ªSB","AM3"),("1ªSB","AL3"),("2ªSB","D3"),("4ªSB","D3"),("1ªSB","D3")]:
    print(f"{sh}!{c}: {str(t(wb[sh][c].value))[:180]}  | valor={wv[sh][c].value!r}")
print()
# erros
err={}
for ws in wv.worksheets:
    n=sum(1 for row in ws.iter_rows() for cc in row if isinstance(cc.value,str) and cc.value.startswith("#"))
    if n: err[ws.title]=n
print("erros por aba:", err)
print("RESUMO!S9 =", wv["RESUMO"]["S9"].value)
