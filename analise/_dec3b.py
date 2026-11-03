from openpyxl import load_workbook
ORIG = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wv = load_workbook(ORIG, data_only=True)
ws = wv["2ªSB"]
def cnt(r,c1,c2,cods):
    return sum(1 for c in range(c1,c2+1) if ws.cell(r,c).value in cods)
alvos = [("AK",17,7,37),("AK",19,7,37),("AK",24,7,37),("AK",30,6,35),
         ("AL",5,6,35),("AL",14,6,34),("AL",17,7,37),("AL",19,7,37),
         ("AL",24,7,37),("AL",26,6,35),("AL",30,6,35)]
print("col  linha  militar                     hoje  correto  delta")
tot=0
for col,r,c1,c2 in alvos:
    nome = ws.cell(r,4).value
    if col=="AK":
        hoje = cnt(r,c1,c2,{"X"})*2 + cnt(r,c1,c2,{"SPJ"}) + cnt(r,c1,c2,{"S1"})
        certo = cnt(r,6,36,{"X"})*2 + cnt(r,6,36,{"SPJ"}) + cnt(r,6,36,{"S1"})
    else:
        hoje = cnt(r,c1,c2,{"PJ"})*2 + cnt(r,c1,c2,{"SPJ"}) + cnt(r,c1,c2,{"PJ12H"})
        certo = cnt(r,6,36,{"PJ"})*2 + cnt(r,6,36,{"SPJ"}) + cnt(r,6,36,{"PJ12H"}) + cnt(r,6,36,{"PJ12"})
    cache = ws[f"{col}{r}"].value
    print(f" {col}   {r:3}  {str(nome)[:26]:28} {cache}      {certo}      {certo-(cache or 0)}")
    tot += certo-(cache or 0)
print("soma dos deltas:", tot)
