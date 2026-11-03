from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
wb = load_workbook(F); wv = load_workbook(F, data_only=True)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
print("abas:", wb.sheetnames)
print()
for sh,c in [("RESUMO","I2"),("RESUMO","J2"),("RESUMO","K2")]:
    print(f"{sh}!{c}: {str(t(wb[sh][c].value))[:170]}")
print()
for sh,c in [("1ªSB","AM3"),("1ªSB","AL3"),("2ªSB","D3"),("3ªSB","D3"),("4ªSB","D3"),("1ªSB","D3")]:
    v=wv[sh][c].value
    print(f"{sh}!{c}: {str(t(wb[sh][c].value))[:120]} | valor={v!r}")
print()
err={}
for ws in wv.worksheets:
    n=sum(1 for row in ws.iter_rows() for cc in row if isinstance(cc.value,str) and cc.value.startswith("#"))
    if n: err[ws.title]=n
print("erros:", err)
print("S9 =", wv["RESUMO"]["S9"].value, "| S4:S8 =", [wv["RESUMO"][f"S{r}"].value for r in range(4,9)])
print("1ªSB BM dia1 (G31?) :", wv["1ªSB"]["G31"].value, "| G30:", wv["1ªSB"]["G30"].value)
