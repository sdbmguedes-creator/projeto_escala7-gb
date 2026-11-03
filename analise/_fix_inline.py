from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(F)
alvos = [("1ªSB", 39), ("2ªSB", 38), ("3ªSB", 38), ("4ªSB", 38)]  # AM=39, AL=38
feitos = []
for sh, colidx in alvos:
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, colidx)
        v = c.value
        t = v.text if isinstance(v, ArrayFormula) else v
        if not isinstance(t, str):
            continue
        s = t.strip()
        if s.startswith("=COUNTIF(") and s.endswith('"PJ")*2'):
            rng = s[len("=COUNTIF("):s.index(",")]
            novo = s + '+COUNTIF(' + rng + ',"PJ12")+COUNTIF(' + rng + ',"PJ12H")'
            c.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo
            feitos.append(sh + "!" + c.coordinate)
wb.save(F)
print("completados:", len(feitos), feitos[:5])
wb2 = load_workbook(F)
ok = inc = 0
for sh, colidx in alvos:
    ws = wb2[sh]
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, colidx).value
        t = v.text if isinstance(v, ArrayFormula) else v
        if isinstance(t, str) and t.startswith("=COUNTIF") and "*2" in t and '"PJ"' in t:
            ok += 1 if ('"PJ12"' in t and '"PJ12H"' in t) else 0
            inc += 0 if ('"PJ12"' in t and '"PJ12H"' in t) else 1
print("PJES completos:", ok, "| incompletos:", inc)
