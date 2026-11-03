"""P1b + P9 sobre o arquivo do Bloco A. Nao toca no original."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"

wb = load_workbook(SRC)


def get_text(cell):
    v = cell.value
    return v.text if isinstance(v, ArrayFormula) else v


def put_text(cell, novo):
    v = cell.value
    if isinstance(v, ArrayFormula):
        cell.value = ArrayFormula(ref=v.ref, text=novo)
    else:
        cell.value = novo


log = {"p1b_fix": [], "p1b_clear": [], "p9_diario": [], "p9_pjes": []}

# ---------- P1b: corrigir a numeracao dos visitantes (1aSB A55:A92) ----------
ws = wb["1ªSB"]
novo = '=_xlfn.XLOOKUP($D{r},RESUMO!$F$2:$F$111,RESUMO!$A$2:$A$111,"",0,0)'
for r in range(55, 93):
    c = ws[f"A{r}"]
    t = get_text(c)
    if isinstance(t, str) and "XLOOKUP" in t and "RESUMO!$F$2:$F$111" in t:
        put_text(c, novo.format(r=r))
        log["p1b_fix"].append(f"A{r}")

# ---------- P1b: apagar 4 celulas orfas ----------
for sh, cell in [("1ªSB", "B29"), ("1ªSB", "C29"), ("ACO", "A54"), ("ACO", "B54")]:
    wb[sh][cell].value = None
    log["p1b_clear"].append(f"{sh}!{cell}")

# ---------- P9: meta diaria inclui PJ12H no grupo de 0,5 ----------
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    ws = wb[sh]
    for row in ws.iter_rows():
        for c in row:
            t = get_text(c)
            if isinstance(t, str) and '{"PJ12";"S1"}' in t:
                put_text(c, t.replace('{"PJ12";"S1"}', '{"PJ12";"PJ12H";"S1"}'))
                log["p9_diario"].append(f"{sh}!{c.coordinate}")

# ---------- P9: coluna de PJES aceita os dois codigos de 12h ----------
alvo = {"1ªSB": "AM", "2ªSB": "AL", "3ªSB": "AL", "4ªSB": "AL"}
pat = re.compile(r'COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"([A-Z0-9]+)"\)\s*$')
for sh, col in alvo.items():
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        c = ws[f"{col}{r}"]
        t = get_text(c)
        if not (isinstance(t, str) and t.startswith("=COUNTIF") and "*2" in t and '"SPJ"' in t):
            continue
        m = pat.search(t)
        if not m:
            continue
        rng, ultimo = m.group(1), m.group(2)
        falta = {"PJ12": "PJ12H", "PJ12H": "PJ12"}.get(ultimo)
        if not falta or f'"{falta}"' in t:
            continue
        put_text(c, t + f'+COUNTIF({rng},"{falta}")')
        log["p9_pjes"].append(f"{sh}!{c.coordinate}")

wb.save(OUT)
for k, v in log.items():
    print(f"{k}: {len(v)}")
print("exemplos P9 diario:", log["p9_diario"][:3])
print("exemplos P9 pjes  :", log["p9_pjes"][:5])
