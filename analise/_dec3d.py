from openpyxl import load_workbook
ORIG = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
ws = load_workbook(ORIG, data_only=True)["2ªSB"]
alvos = [("AK",17,7,36),("AK",19,7,36),("AK",24,7,36),("AK",30,6,35),
         ("AL",5,6,35),("AL",14,6,34),("AL",17,7,36),("AL",19,7,36),
         ("AL",24,7,36),("AL",26,6,35),("AL",30,6,35)]
def contar(r,c1,c2,cods):
    return sum(1 for c in range(c1,c2+1) if ws.cell(r,c).value in cods)
print("col linha militar                    dias fora   valor")
mudam=0
for col,r,c1,c2 in alvos:
    nome = ws.cell(r,4).value
    faltam = [c-5 for c in range(6,37) if not (c1 <= c <= c2)]
    if col=="AK":
        antes = contar(r,c1,c2,{"X"})*2 + contar(r,c1,c2,{"SPJ"}) + contar(r,c1,c2,{"S1"})
        depois = contar(r,6,36,{"X"})*2 + contar(r,6,36,{"SPJ"}) + contar(r,6,36,{"S1"})
    else:
        antes = contar(r,c1,c2,{"PJ"})*2 + contar(r,c1,c2,{"SPJ"}) + contar(r,c1,c2,{"PJ12"}) + contar(r,c1,c2,{"PJ12H"})
        depois = contar(r,6,36,{"PJ"})*2 + contar(r,6,36,{"SPJ"}) + contar(r,6,36,{"PJ12"}) + contar(r,6,36,{"PJ12H"})
    marca = "  <== muda" if antes != depois else ""
    if antes != depois: mudam += 1
    print(f" {col}  {r:3}  {str(nome)[:24]:26} {str(faltam):11} {antes} -> {depois}{marca}")
print("celulas cujo valor muda:", mudam, "de", len(alvos))
