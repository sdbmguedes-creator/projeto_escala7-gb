"""Prova estatica v2: usa o conjunto REAL de celulas limpas e agrupa por padrao."""
import re
from collections import Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"

orig = load_workbook(SRC)

# conjunto real: formulas que apontavam para a aba excluida
cleared = {}
for ws in orig.worksheets:
    if ws.title in ("FILTRO_DIAS", "Cópia de RESUMO", "Página30"):
        continue
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            t = v.text if hasattr(v, "text") else v
            if isinstance(t, str) and "'VOLUNTÁRIOS PJES'!" in t:
                cleared.setdefault(ws.title, set()).add((c.row, c.column))
# linha 82 da 1aSB
cleared.setdefault("1ªSB", set()).update({(82, c) for c in range(1, 41)})

print("celulas limpas por aba:", {k: len(v) for k, v in cleared.items()},
      "| total:", sum(len(v) for v in cleared.values()))


def col(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


pat = re.compile(r"(?:'([^']+)'|([A-Za-zÀ-ÿ0-9_]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
wb = load_workbook(OUT)
patterns = Counter()
examples = {}
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            t = v.text if hasattr(v, "text") else v
            if not (isinstance(t, str) and t.startswith("=")):
                continue
            for m in pat.finditer(t):
                sh = m.group(1) or m.group(2)
                if sh not in cleared:
                    continue
                c1, r1 = col(m.group(3)), int(m.group(4))
                c2 = col(m.group(5)) if m.group(5) else c1
                r2 = int(m.group(6)) if m.group(6) else r1
                lo_c, hi_c = min(c1, c2), max(c1, c2)
                lo_r, hi_r = min(r1, r2), max(r1, r2)
                touch = [(rr, cc) for (rr, cc) in cleared[sh] if lo_r <= rr <= hi_r and lo_c <= cc <= hi_c]
                if not touch:
                    continue
                # nome do padrao: funcao + formato do intervalo
                fn = re.match(r"=([A-Z_\.]+)\(", t.upper())
                key = f"{fn.group(1) if fn else '?'} | {sh}!{get_column_letter(lo_c)}{lo_r}:{get_column_letter(hi_c)}{hi_r}"
                patterns[key] += 1
                examples.setdefault(key, (f"{ws.title}!{c.coordinate}", t[:110]))

print("\npadroes que tocam celula limpa:")
for k, n in patterns.most_common(20):
    print(f"  x{n:<4} {k}")
    print(f"        ex: {examples[k][0]}  {examples[k][1]}")
