"""Base 2.0 refinada: dedupe por matricula e blocos CONTIGUOS de N na aba de origem."""
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_base.xlsx"

# ---------- PESSOAS + MES (novembro) ----------
pessoas = {}
for p in ARQS.values():
    wv = load_workbook(p, data_only=True)
    res = wv["RESUMO"]
    for r in range(2, 160):
        posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
        if not all(isinstance(x, str) for x in (posto, mat, guerra)):
            continue
        pessoas.setdefault(mat.strip(), (posto.strip(), guerra.strip()))

wv = load_workbook(ARQS["NOV"], data_only=True)
res = wv["RESUMO"]
mes, vistos = [], set()
for r in range(2, 160):
    posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
    if not all(isinstance(x, str) for x in (posto, mat, guerra)):
        continue
    mat = mat.strip()
    if mat in vistos:
        continue
    vistos.add(mat)
    esc = res.cell(r, 8).value
    mot = res.cell(r, 7).value
    adm = "S" if (isinstance(esc, str) and "ADM" in esc) else ""
    mes.append((mat, res.cell(r, 5).value, esc, adm, mot))

# ---------- AFASTAMENTOS: blocos contiguos de N na propria aba ----------
COLS = {"1ªSB": ("G", "AK"), "2ªSB": ("F", "AJ"), "3ªSB": ("F", "AJ"), "4ªSB": ("F", "AJ")}
sec_de = {m[0]: m[1] for m in mes}
afast = []
for mat, secao, esc, adm, mot in mes:
    if secao not in COLS:
        continue
    sh = secao
    c1, c2 = COLS[sh]
    ws = wv[sh]
    for r in range(3, 120):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str):
            continue
        m = re.search(r"\d{6}-\d", nome)
        if not m or m.group(0) != mat:
            continue
        dias = {}
        for ci in range(ws[c1 + "1"].column, ws[c2 + "1"].column + 1):
            d = ws.cell(2, ci).value
            if d is not None:
                dias[d.day] = ws.cell(r, ci).value
        # blocos contiguos de N
        blocos, ini = [], None
        for dia in sorted(dias):
            if dias[dia] == "N":
                if ini is None:
                    ini = dia
                fim = dia
            else:
                if ini is not None:
                    blocos.append((ini, fim))
                    ini = None
        if ini is not None:
            blocos.append((ini, fim))
        for a, b in blocos:
            if b - a + 1 >= 3:          # ignora N isolado (folga), guarda periodos
                afast.append((mat, nome, a, b, b - a + 1))
        break

print("pessoas:", len(pessoas), "| linhas no MES:", len(mes), "| periodos de N (>=3 dias):", len(afast))
print()
for a in sorted(afast, key=lambda x: -x[4])[:20]:
    print(f"   {a[1][:30]:32} {a[2]:2} a {a[3]:2}  ({a[4]} dias)")

# ---------- grava ----------
wb = Workbook()
ws = wb.active
ws.title = "PESSOAS"
ws.append(["ID", "POSTO", "NOME DE GUERRA"])
for c in ws[1]:
    c.font = Font(bold=True)
for mat, (posto, guerra) in sorted(pessoas.items()):
    ws.append([mat, posto, guerra])

ws2 = wb.create_sheet("MES")
ws2.append(["ID", "SEÇÃO", "ESCALA", "ADM", "MOTORISTA"])
for c in ws2[1]:
    c.font = Font(bold=True)
for m in mes:
    ws2.append(list(m))

ws3 = wb.create_sheet("AFASTAMENTOS")
ws3.append(["ID", "NOME", "DE", "ATÉ", "DIAS"])
for c in ws3[1]:
    c.font = Font(bold=True)
for a in sorted(afast):
    ws3.append([a[0], a[1], a[2], a[3], a[4]])

for w, largs in [(ws, (14, 12, 26)), (ws2, (14, 10, 12, 9, 12)), (ws3, (14, 30, 7, 7, 7))]:
    w.freeze_panes = "A2"
    for i, larg in enumerate(largs, start=1):
        w.column_dimensions[w.cell(1, i).column_letter].width = larg

wb.save(OUT)
print()
print("salvo:", OUT)
