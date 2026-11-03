import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
ORIG = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
V5 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
wa, wb = load_workbook(ORIG), load_workbook(V5)
wv = load_workbook(ORIG, data_only=True)
def t(v): return v.text if isinstance(v,ArrayFormula) else v
TERM = re.compile(r'COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"([A-Z0-9]+)"\)(\*2)?')
def avalia(ws, formula):
    if not isinstance(formula,str): return None
    total=0
    for m in TERM.finditer(formula):
        rng, cod, dobro = m.group(1), m.group(2), m.group(3)
        cells = ws[rng.replace("$","")]
        vals = [c.value for row in cells for c in row]
        total += vals.count(cod) * (2 if dobro else 1)
    return total
print("=== PJES: valor original vs v5 (somente onde muda) ===")
delta_total = 0
for sh, colidx in [("1ªSB",39),("2ªSB",38),("3ªSB",38),("4ªSB",38)]:
    for r in range(3, 120):
        fo = t(wa[sh].cell(r,colidx).value); fn = t(wb[sh].cell(r,colidx).value)
        if not isinstance(fo,str) or not isinstance(fn,str): continue
        vo, vn = avalia(wa[sh], fo), avalia(wb[sh], fn)
        if vo is None or vn is None: continue
        if vo != vn:
            delta_total += vn-vo
            print(f"  {sh}!{wb[sh].cell(r,colidx).coordinate} {str(wb[sh].cell(r,4).value)[:26]:28} {vo} -> {vn}")
print("delta total:", delta_total)
