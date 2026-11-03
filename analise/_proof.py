"""Prova estatica: nenhuma formula sobrevivente referencia as celulas limpas."""
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"

CLEARED = {
    "RESUMO": [(r, 16) for r in range(2, 113)],
    "ACO": [(r, 5) for r in range(6, 55)],
    "1ªSB": [(r, 5) for r in range(3, 93)] + [(82, c) for c in range(1, 41)],
    "2ªSB": [(r, 5) for r in range(3, 54)],
    "3ªSB": [(r, 5) for r in range(3, 42)],
    "4ªSB": [(r, 5) for r in range(3, 56)],
}
CLEARED_SETS = {k: set(v) for k, v in CLEARED.items()}


def col(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def boxes(text):
    """Devolve (aba, c1, r1, c2, r2) para cada referencia A1 do texto."""
    out = []
    pat = re.compile(
        r"(?:'([^']+)'|([A-Za-zÀ-ÿ0-9_]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?"
    )
    for m in pat.finditer(text):
        c1, r1 = m.group(3), int(m.group(4))
        c2 = m.group(5) or c1
        r2 = int(m.group(6)) if m.group(6) else r1
        out.append((m.group(1) or m.group(2), col(c1), r1, col(c2), r2))
    return out


wb = load_workbook(OUT)
sheets = set(wb.sheetnames)
hits = {}
nformulas = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            t = v.text if hasattr(v, "text") else v
            if not (isinstance(t, str) and t.startswith("=")):
                continue
            nformulas += 1
            for sh, c1, r1, c2, r2 in boxes(t):
                if sh not in CLEARED_SETS:
                    continue
                lo_c, hi_c = min(c1, c2), max(c1, c2)
                lo_r, hi_r = min(r1, r2), max(r1, r2)
                for (rr, cc) in CLEARED_SETS[sh]:
                    if lo_r <= rr <= hi_r and lo_c <= cc <= hi_c:
                        hits.setdefault(f"{ws.title}!{c.coordinate}", []).append(
                            f"{sh}!{get_column_letter(cc)}{rr}")
                        break

print("formulas no arquivo limpo:", nformulas)
print("formulas que LEEM alguma celula limpa:", len(hits))
for k, v in list(hits.items())[:25]:
    print("  ", k, "->", v[:4])
