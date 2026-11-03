import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
V5 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
ORIG = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(V5); wv = load_workbook(ORIG, data_only=True)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
print("=== RESUMO J..N (formulas) ===")
for c in ["J2","K2","L2","M2","N2"]:
    print(f"{c}: {t(wb['RESUMO'][c].value)}")
print()
print("=== soma dos PJES por aba (minha contagem, dados de outubro) ===")
TERM = re.compile(r'COUNTIF\(\$?([A-Z]{1,3})\$?\d+:\$?[A-Z]{1,3}\$?\d+,"([A-Z0-9]+)"\)(\*2)?')
def avalia(ws, formula):
    total=0
    for m in TERM.finditer(formula):
        cod, dobro = m.group(2), m.group(3)
        col = m.group(1)
        n = sum(1 for r in range(1, ws.max_row+1) if ws.cell(r, {
            "A":1,"B":2,"C":3,"D":4,"E":5,"F":6,"G":7}.get("F",6)).value is None)
        total += 0
    return None
# contagem direta: percorre linhas e soma pelo nome
res = wb["RESUMO"]
total = 0
for r in range(2,112):
    for colidx, sh in [(10,"ACO"),(11,"1ªSB"),(12,"2ªSB"),(13,"3ªSB"),(14,"4ªSB")]:
        pass
print("(contagem direta abaixo)")
for sh,colidx in [("1ªSB",39),("2ªSB",38),("3ªSB",38),("4ªSB",38)]:
    ws=wb[sh]; s=0
    for r in range(3,120):
        f=t(ws.cell(r,colidx).value)
        if not isinstance(f,str): continue
        for m in re.finditer(r'COUNTIF\(\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+,"([A-Z0-9]+)"\)(\*2)?', f):
            c1,c2,cod,dob = m.group(1),m.group(2),m.group(3),m.group(4)
            def ci(l):
                n=0
                for ch in l: n=n*26+(ord(ch)-64)
                return n
            n=sum(1 for rr in range(3,120) for cc in range(ci(c1),ci(c2)+1) if ws.cell(rr,cc).value==cod)
            s += n*(2 if dob else 1)
    print(f"  {sh}: {s}")
    total += s
print("  TOTAL SBs:", total)
