from openpyxl import load_workbook
wv = load_workbook(r'C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx', data_only=True)
ws = wv['1ªSB']
for r in range(3, 120):
    cods = [(ws.cell(r,c).coordinate, ws.cell(r,c).value) for c in range(7,38) if ws.cell(r,c).value=='SB']
    if cods:
        print(f'  linha {r}: D={ws.cell(r,4).value!r} | {len(cods)} celulas SB | primeiras: {cods[:4]}')
