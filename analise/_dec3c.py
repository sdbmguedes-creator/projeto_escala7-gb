from openpyxl import load_workbook
ORIG = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
ws = load_workbook(ORIG, data_only=True)["2ªSB"]
alvos = [("AK",17,7,"AJ",17),("AK",19,7,"AJ",19),("AK",24,7,"AJ",24),("AK",30,6,"AI",30),
         ("AL",5,6,"AI",5),("AL",14,6,"AH",14),("AL",17,7,"AJ",17),("AL",19,7,"AJ",19),
         ("AL",24,7,"AJ",24),("AL",26,6,"AI",26),("AL",30,6,"AI",30)]
def contar(r,c1,c2,cods):
    return sum(1 for c in range(c1,c2+1) if ws.cell(r,c).value in cods)
print("col linha militar                  dias_fora   impacto no valor")
mudam = 0
for col,r,c1,c2 in alvos:
    nome = ws.cell(r,4).value
    # colunas faltantes entre F(6) e AJ(36)
    faltam = [c for c in range(6,37) if not (c1 <= c <= c2)]
    dias = [c-5 for c in faltam]
    if col=="AK":
        antes = contar(r,c1,c2,{"X"})*2 + contar(r,c1,c2,{"SPJ"}) + contar(r,c1,c2,{"S1"})
        depois = contar(r,6,36,{"X"})*2 + contar(r,6,36,{"SPJ"}) + contar(r,6,36,{"S1"})
    else:
        antes = contar(r,c1,c2,{"PJ"})*2 + contar(r,c1,c2,{"SPJ"}) + contar(r,c1,c2,{"PJ12"}) + contar(r,c1,c2,{"PJ12H"})
        depois = contar(r,6,36,{"PJ"})*2 + contar(r,6,36,{"SPJ"}) + contar(r,6,36,{"PJ12"}) + contar(r,6,36,{"PJ12H"})
    if antes != depois:
        mudam += 1
        print(f" {col}  {r:3}  {str(nome)[:24]:26} dias {dias}   {antes} -> {depois}")
print("celulas cujo valor muda:", mudam, "de", len(alvos))
