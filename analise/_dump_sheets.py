"""Dump each sheet as text (only rows with content) plus formula patterns."""
import os
import re
from collections import Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

PATH = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\analise\dumps"
os.makedirs(OUT, exist_ok=True)

wb = load_workbook(PATH, data_only=False)
wbv = load_workbook(PATH, data_only=True)

index = []
for i, ws in enumerate(wb.worksheets, 1):
    wsv = wbv[ws.title]
    rows = {}
    formulas = Counter()
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or (isinstance(c.value, str) and c.value.strip() == ""):
                continue
            rows.setdefault(c.row, []).append(c)
            if isinstance(c.value, str) and c.value.startswith("="):
                formulas[c.value] += 1

    safe = re.sub(r"[^A-Za-z0-9]+", "_", ws.title).strip("_")
    fname = f"{i:02d}_{safe}.txt"
    with open(os.path.join(OUT, fname), "w", encoding="utf-8") as fh:
        fh.write(f"SHEET: {ws.title}  state={ws.sheet_state}  "
                 f"max_row={ws.max_row} max_col={ws.max_column}\n")
        fh.write(f"merged: {[str(m) for m in ws.merged_cells.ranges]}\n")
        cols_used = sorted({c.column for r in rows.values() for c in r})
        fh.write("colunas usadas: " + ", ".join(get_column_letter(x) for x in cols_used) + "\n")
        fh.write("=" * 90 + "\n")
        for r in sorted(rows):
            parts = []
            for c in sorted(rows[r], key=lambda x: x.column):
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    cached = wsv.cell(row=c.row, column=c.column).value
                    v = f"{v}   <<cached: {cached!r}>>"
                parts.append(f"{c.coordinate}= {v!r}")
            fh.write(f"[{r}] " + " | ".join(parts) + "\n")

    index.append((i, ws.title, ws.sheet_state, len(rows), len(formulas), fname, dict(formulas.most_common(15))))

with open(os.path.join(OUT, "_INDEX.txt"), "w", encoding="utf-8") as fh:
    for i, title, state, nrows, nf, fname, top in index:
        fh.write(f"{i:02d} {title} [{state}] linhas={nrows} formulas={nf} -> {fname}\n")
        for f, n in top.items():
            fh.write(f"     x{n:<4} {f[:300]}\n")
        fh.write("\n")

print("ok")
