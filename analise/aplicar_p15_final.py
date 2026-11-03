"""P15 final: limpa a coluna P, marca os ADM e troca a formula de regime."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"

wb = load_workbook(F)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def put(cell, novo):
    v = cell.value
    cell.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo


RES = wb["RESUMO"]

# 1) limpar residuos de texto da antiga coluna de voluntarios
limpos = 0
for r in range(2, 112):
    if RES.cell(r, 16).value is not None:
        RES.cell(r, 16).value = None
        limpos += 1
RES["P1"] = "ADM DA SB"

# 2) marcar os ADM
OVERRIDE = {"CB 711063-4 JOÃO LIANDA": "ADM 2ªSB", "3º SGT 710130-9 ANDRADE": "ADM 3ªSB"}
marc = []
for r in range(2, 112):
    p, m, g = (RES.cell(r, c).value for c in (2, 3, 4))
    if not all(isinstance(x, str) for x in (p, m, g)):
        continue
    nome = f"{p} {m} {g}"
    if nome in OVERRIDE:
        RES.cell(r, 16).value = OVERRIDE[nome]
        marc.append(nome)
    elif RES.cell(r, 5).value == "ADM":
        RES.cell(r, 16).value = "ADM 7ºGB"
        marc.append(nome)

# 3) formula de regime
B = {"1ªSB": ("AL", [(3, 29), (39, 49), (55, 92)]), "2ªSB": ("AK", [(3, 30), (39, 53)]),
     "3ªSB": ("AK", [(3, 18), (27, 41)]), "4ªSB": ("AK", [(3, 31), (40, 55)])}
M = ('=IF(XLOOKUP($D{r},RESUMO!$F$2:$F$400,RESUMO!$P$2:$P$400,"",0,0)<>"",'
     'COUNTIF({rg},"X")*2+COUNTIF({rg},"SPJ")+COUNTIF({rg},"S1"),COUNTIF({rg},"X"))')
n = 0
for sh, (col, bl) in B.items():
    ws = wb[sh]
    for a, b in bl:
        for r in range(a, b + 1):
            if not isinstance(ws.cell(r, 4).value, str):
                continue
            cell = ws[f"{col}{r}"]
            at = t(cell.value)
            if not isinstance(at, str) or "COUNTIF" not in at:
                continue
            mm = re.search(r'COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"X"\)', at)
            if not mm:
                continue
            put(cell, M.format(r=r, rg=mm.group(1)))
            n += 1

wb.save(OUT)
print("limpos:", limpos, "| marcados:", len(marc), "| formulas:", n)
