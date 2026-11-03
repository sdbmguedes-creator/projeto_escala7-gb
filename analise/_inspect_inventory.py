"""Inventory of the workbook: sheets, dimensions, formulas, validations, formats."""
import json
import re
from collections import Counter
from openpyxl import load_workbook

PATH = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"

wb = load_workbook(PATH, data_only=False)
wbv = load_workbook(PATH, data_only=True)

print("== ORDER / METADATA ==")
print("sheets:", wb.sheetnames)
for ws in wb.worksheets:
    print(f"  {ws.title!r}: state={ws.sheet_state} dims={ws.dimensions} "
          f"max_row={ws.max_row} max_col={ws.max_column} "
          f"merged={len(ws.merged_cells.ranges)} images={len(ws._images)} "
          f"charts={len(ws._charts)} tables={len(ws.tables)}")

print("\n== DEFINED NAMES ==")
try:
    for name, dn in wb.defined_names.items():
        print(f"  {name} -> {dn.attr_text}  (hidden={dn.hidden})")
except Exception as exc:  # pragma: no cover
    print("  erro:", exc)

print("\n== PER SHEET ==")
for ws in wb.worksheets:
    wsv = wbv[ws.title]
    formulas, values = [], 0
    funcs = Counter()
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            values += 1
            if isinstance(c.value, str) and c.value.startswith("="):
                formulas.append((c.coordinate, c.value))
                for fn in re.findall(r"([A-Z][A-Z0-9\.]+)\(", c.value.upper()):
                    funcs[fn] += 1
    print(f"\n--- {ws.title} ---")
    print(f"  celulas preenchidas: {values} | formulas: {len(formulas)}")
    print(f"  funcoes: {json.dumps(dict(funcs.most_common(25)), ensure_ascii=False)}")
    print(f"  freezed: {ws.freeze_panes} | autofilter: {ws.auto_filter.ref} | "
          f"sheet_protection: {bool(ws.protection.sheet)}")
    dv = ws.data_validations.dataValidation
    print(f"  data_validations: {len(dv)}")
    for d in dv:
        print(f"    range={d.sqref} type={d.type} op={d.operator} "
              f"f1={str(d.formula1)[:120]} f2={str(d.formula2)[:60]} allow_blank={d.allow_blank}")
    cf = []
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for r in rules:
            cf.append((str(rng.sqref), r.type, r.operator,
                       str(r.formula)[:100] if r.formula else None,
                       r.dxf.fill.fgColor.rgb if (r.dxf and r.dxf.fill and r.dxf.fill.fgColor) else None))
    print(f"  conditional_formats: {len(cf)}")
    for x in cf[:12]:
        print("   ", x)
    if len(cf) > 12:
        print(f"    ... +{len(cf)-12}")

print("\n== CACHED VALUES SAMPLE (first sheet) ==")
ws0 = wb.worksheets[0]
wsv0 = wbv[ws0.title]
for r in range(1, min(12, ws0.max_row) + 1):
    cells = []
    for cc in range(1, min(10, ws0.max_column) + 1):
        c = ws0.cell(row=r, column=cc)
        if c.value is not None:
            cached = wsv0.cell(row=r, column=cc).value
            cells.append(f"{c.coordinate}={str(c.value)[:60]!r}" + (f"[cached={cached!r}]" if str(c.value).startswith('=') else ""))
    if cells:
        print(f"  row {r}: " + " | ".join(cells))
