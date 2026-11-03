"""P15: coluna P = ADM DA SB (RESUMO) + formula de regime na coluna ORDINARIO."""
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
    nome = RES.cell(r, 6).value          # coluna F = NOME (texto composto)
    secao = RES.cell(r, 5).value         # coluna E = SEÇÃO
    if not isinstance(nome, str):
        continue
    if nome in OVERRIDE:
        RES.cell(r, 16).value = OVERRIDE[nome]
        marcados.append((r, nome, OVERRIDE[nome]))
    elif secao == "ADM":
        RES.cell(r, 16).value = "ADM 7ºGB"
        marcados.append((r, nome, "ADM 7ºGB"))

# ---------- formula de regime na coluna ORDINARIO ----------
BLOCOS = {"1ªSB": ("AL", [(3, 29), (39, 49), (55, 92)]),
          "2ªSB": ("AK", [(3, 30), (39, 53)]),
          "3ªSB": ("AK", [(3, 18), (27, 41)]),
          "4ªSB": ("AK", [(3, 31), (40, 55)])}

modelo = ('=IF(XLOOKUP($D{r},RESUMO!$F$2:$F$400,RESUMO!$P$2:$P$400,"",0,0)<>"",'
          'COUNTIF({rg},"X")*2+COUNTIF({rg},"SPJ")+COUNTIF({rg},"S1"),'
          'COUNTIF({rg},"X"))')

trocados = []
for sh, (col, blocos) in BLOCOS.items():
    ws = wb[sh]
    for a, b in blocos:
        for r in range(a, b + 1):
            if not isinstance(ws.cell(r, 4).value, str):
                continue
            cell = ws[f"{col}{r}"]
            atual = t(cell.value)
            if not isinstance(atual, str) or "COUNTIF" not in atual:
                continue
            m = re.search(r'COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"X"\)', atual)
            if not m:
                continue
            put(cell, modelo.format(r=r, rg=m.group(1)))
            trocados.append(f"{sh}!{cell.coordinate}")

wb.save(OUT)
print("RESUMO!P marcados:", len(marcados))
for x in marcados:
    print("   ", x)
print("formulas de regime trocadas:", len(trocados))
print("exemplos:", trocados[:4])
