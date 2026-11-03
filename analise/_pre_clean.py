from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P, data_only=False)
wbv = load_workbook(P, data_only=True)
print("=== celulas com formula apontando para VOLUNTARIOS PJES ===")
tot=0
for ws in wb.worksheets:
    cells=[]
    for row in ws.iter_rows():
        for c in row:
            v=c.value
            t=v.text if hasattr(v,"text") else v
            if isinstance(t,str) and "VOLUNT" in t.upper():
                cells.append(c.coordinate)
    if cells:
        tot+=len(cells)
        print(f"  {ws.title}: {len(cells)} -> {cells[:6]} ... {cells[-3:]}")
print("  TOTAL:", tot)
print()
print("=== 1aSB: MONTEIRO linhas 21 e 82 ===")
ws=wb["1ªSB"]; wsv=wbv["1ªSB"]
for r in (21,82):
    marks=[f"{ws.cell(r,c).coordinate}={wsv.cell(r,c).value}" for c in range(7,38) if wsv.cell(r,c).value not in (None,"")]
    print(f" linha {r}: D={ws.cell(r,4).value!r} | B={wsv.cell(r,2).value!r} | marcas={marks}")
    print(f"     AL={wsv.cell(r,38).value} AM={wsv.cell(r,39).value} AN={wsv.cell(r,40).value}")
print()
print("=== ACO coluna E (primeiras 6) ===")
wa=wb["ACO"]
print([ (wa.cell(r,5).coordinate, str(wa.cell(r,5).value)[:70]) for r in range(6,12)])
print()
print("=== 1aSB E39:E49 (NAO apagar) ===")
print([(ws.cell(r,5).coordinate, ws.cell(r,5).value) for r in range(39,50)])
