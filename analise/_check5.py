import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P); wbv = load_workbook(P, data_only=True)
print("=== rotulos das colunas de saldo ===")
for sh, cols in [("1ªSB",["AJ","AK","AL","AM","AN","AO"]),("2ªSB",["AI","AJ","AK","AL","AM"]),
                 ("4ªSB",["AI","AJ","AK","AL","AM"])]:
    ws=wb[sh]
    print(f" {sh}: " + " | ".join(f"{c}{r}={ws[f'{c}{r}'].value!r}" for c in cols for r in (1,2) if ws[f'{c}{r}'].value))
print()
print("=== quem consome RESUMO coluna I (ORD) ===")
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v=c.value; t=v.text if isinstance(v,ArrayFormula) else v
            if isinstance(t,str) and re.search(r"RESUMO!\$?I", t):
                print(f"  {ws.title}!{c.coordinate}: {t[:110]}")
print("(vazio = ninguem usa)")
print()
print("=== exemplo: 8 servicos X em cada SB -> quanto cada coluna devolve ===")
for sh,col,row,nome in [("1ªSB","AL",3,"TELES"),("2ªSB","AK",4,"VINICIUS"),("3ªSB","AK",4,"(2a linha 3aSB)"),("4ªSB","AK",4,"(2a linha 4aSB)")]:
    ws=wbv[sh]
    xs=sum(1 for c in range(7 if sh=="1ªSB" else 6, 38 if sh=="1ªSB" else 37)
           if ws.cell(row,c).value=="X")
    print(f"  {sh}!{col}{row} ({nome}): X no mes={xs} | valor da coluna={ws[f'{col}{row}'].value}")
