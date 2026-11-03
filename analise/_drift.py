import re
from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P)
spec = {"1ªSB": ("AL", "AM", "G", "AK", [(3,29)]), "2ªSB": ("AK","AL","F","AJ",[(3,30)]),
        "3ªSB": ("AK","AL","F","AJ",[(3,41)]), "4ªSB": ("AK","AL","F","AJ",[(3,31)])}
for sh,(cal,cal2,c1,c2,blocks) in spec.items():
    ws = wb[sh]; bad = []
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not (isinstance(v,str) and v.startswith("=COUNTIF(") and 'G' not in v[:12]): pass
            if isinstance(v,str) and v.startswith("=COUNTIF(") and ("$"+cal) not in v and '\"X\"' in v or (isinstance(v,str) and v.startswith("=COUNTIF(") and "PJ" in v):
                pass
    # simpler: read explicit ranges in COUNTIF for total columns
    for c in ws[cal]:
        v = c.value
        if isinstance(v,str) and v.startswith("=COUNTIF("):
            m = re.match(r"=COUNTIF\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", v)
            if m and (m.group(1),m.group(3)) != (c1,c2):
                bad.append((c.coordinate, f"{m.group(1)}{m.group(2)}:{m.group(3)}{m.group(4)}"))
    for c in ws[cal2]:
        v = c.value
        if isinstance(v,str) and v.startswith("=COUNTIF("):
            m = re.match(r"=COUNTIF\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", v)
            if m and (m.group(1),m.group(3)) != (c1,c2):
                bad.append((c.coordinate, f"{m.group(1)}{m.group(2)}:{m.group(3)}{m.group(4)}"))
    print(f"{sh}: colunas {cal}/{cal2} esperado {c1}:{c2} -> desvios: {len(bad)}")
    for b in bad[:12]: print("   ", b)
