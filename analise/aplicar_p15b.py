"""P15 (corrigido): coluna P = ADM DA SB + formula de regime."""
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
OVERRIDE = {
    "CB 711063-4 JOÃO LIANDA": "ADM 2ªSB",
    "3º SGT 710130-9 ANDRADE": "ADM 3ªSB",
    "2º SGT 704203-5 ERALDIR": "ADM 4ªSB",
}

RES["P1"] = "ADM DA SB"
marcados = []
for r in range(2, 112):
    posto, mat, guerra = (RES.cell(r, c).value for c in (2, 3, 4))
    if not all(isinstance(x, str) for x in (posto, mat, guerra)):
        continue
    nome = f"{posto} {mat} {guerra}"
    secao = RES.cell(r, 5).value
    if nome in OVERRIDE:
        RES.cell(r, 16).value = OVERRIDE[nome]
        marcados.append((r, nome, OVERRIDE[nome]))
    elif secao == "ADM":
        RES.cell(r, 16).value = "ADM 7ºGB"
        marcados.append((r, nome, "ADM 7ºGB"))

BLOCOS = {"1ªSB": ("AL", [(3, 29), (39, 49), (55, 92)]),
          "2ªSB": ("AK", [(3, 30), (39, 53)]),
          "3ªSB": ("AK", [(3, 18), (27, 41)]),
          "4ªSB": ("AK", [(3, 31), (40, 55)])}

modelo = ('=IF(XLOOKUP($D{r},RESUMO!$F$2:$F$400,RESUMO!$P$2:$P$400,"",0,0)<>"",'
          'COUNTIF({rg},"X")*2+COUNTIF({rg},"SPJ")+COUNTIF({rg},"S1"),'
          'COUNTIF({rg},"X"))')

trocados, pulados = [], []
for sh, (col, blocos) in BLOCOS.items():
    ws = wb[sh]
    for a, b in blocos:
        for r in range(a, b + 1):
            if not isinstance(ws.cell(r, 4).value, str):
                continue
            cell = ws[f"{col}{r}"]
            atual = t(cell.value)
            if not isinstance(atual, str) or "COUNTIF" not in atual:
                pulados.append(f"{sh}!{cell.coordinate}")
                continue
            m = re.search(r'COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"X"\)', atual)
            if not m:
                pulados.append(f"{sh}!{cell.coordinate}")
                continue
            put(cell, modelo.format(r=r, rg=m.group(1)))
            trocados.append(f"{sh}!{cell.coordinate}")

wb.save(OUT)
print("RESUMO!P:", len(marcados))
for x in marcados:
    print("   ", x)
print("formulas trocadas:", len(trocados), "| puladas:", len(pulados), pulados[:8])
