from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P); wbv = load_workbook(P, data_only=True)
res = wbv["RESUMO"]
sec = {}
for r in range(2,112):
    nome = res.cell(r,6).value
    if isinstance(nome,str): sec[nome] = (res.cell(r,5).value, res.cell(r,8).value)
print("=== pessoas das linhas 3 (formula 'diferente') ===")
for sh,col in [("2ªSB","AK"),("3ªSB","AK"),("4ªSB","AK")]:
    nome = wbv[sh].cell(3,4).value
    print(f" {sh}!D3 = {nome!r} -> secao={sec.get(nome)}")
print()
print("=== bloco ADM da 1aSB: secoes ===")
for r in range(39,50):
    nome = wbv["1ªSB"].cell(r,4).value
    if nome: print(f"  linha {r}: {nome!r} -> secao={sec.get(nome)}")
print()
print("=== contagem de ADM por aba (coluna D) ===")
for sh,(a,b) in {"1ªSB":(3,92),"2ªSB":(3,53),"3ªSB":(3,41),"4ªSB":(3,55)}.items():
    adm=[r for r in range(a,b+1) if sec.get(wbv[sh].cell(r,4).value,("",""))[0]=="ADM"]
    print(f"  {sh}: linhas {adm}")
