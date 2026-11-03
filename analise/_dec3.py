import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v4.xlsx"
wb = load_workbook(V); wbv = load_workbook(V, data_only=True)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
def conta(ws, r, c1, c2, cod):
    return sum(1 for c in range(c1,c2+1) if ws.cell(r,c).value==cod)
print("=== (3) 2aSB: celulas com faixa deslocada ===")
ws, wsv = wb["2ªSB"], wbv["2ªSB"]
rows=[]
for ci,col in [(37,"AK"),(38,"AL")]:
    for r in range(3,31):
        f = t(ws.cell(r,ci).value)
        if not isinstance(f,str) or not f.startswith("=COUNTIF"): continue
        ranges = re.findall(r"COUNTIF\(\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+", f)
        if not ranges: continue
        c1,c2 = ranges[0]
        if (c1,c2) != ("F","AJ"):
            atual = wsv.cell(r,ci).value
            # recalcula com a faixa correta
            if col=="AK":
                novo = conta(ws,r,6,36,"X")*2 + conta(ws,r,6,36,"SPJ") + conta(ws,r,6,36,"S1")
            else:
                novo = conta(ws,r,6,36,"PJ")*2 + conta(ws,r,6,36,"SPJ") + conta(ws,r,6,36,"PJ12") + conta(ws,r,6,36,"PJ12H") + conta(ws,r,6,36,"PJ12")
            rows.append((f"{col}{r}", ws.cell(r,4).value, f"{c1}:{c2}", atual, novo))
for x in rows:
    print(f"   {x[0]:5} {str(x[1])[:28]:30} faixa {x[2]:8} hoje={x[3]}  correto={x[4]}")
print("total de celulas com faixa errada:", len(rows))
print("celulas em que o valor MUDA:", sum(1 for x in rows if x[3]!=x[4]))
print()
print("=== (1) ERALDIR — marcacao do mes ===")
ws4 = wbv["4ªSB"]
marc = [(c, ws4.cell(3,c).value) for c in range(6,37) if isinstance(ws4.cell(3,c).value,str)]
print("  marcacoes:", marc)
print("  X:", sum(1 for _,v in marc if v=="X"), "| PJ:", sum(1 for _,v in marc if v=="PJ"),
      "| SPJ:", sum(1 for _,v in marc if v=="SPJ"), "| S1:", sum(1 for _,v in marc if v=="S1"))
print("  RESUMO: secao =", wb["RESUMO"].cell(9,5).value, "| escala =", wb["RESUMO"].cell(9,8).value)
