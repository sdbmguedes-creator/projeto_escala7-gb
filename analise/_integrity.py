from openpyxl import load_workbook
A = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
B = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"
wa = load_workbook(A); wb = load_workbook(B)
print(f"{'aba':22} {'DV':>10} {'CF':>10} {'merges':>10}")
for n in wb.sheetnames:
    a, b = wa[n], wb[n]
    dv_a, dv_b = len(a.data_validations.dataValidation), len(b.data_validations.dataValidation)
    cf_a = sum(len(r) for r in a.conditional_formatting._cf_rules.values())
    cf_b = sum(len(r) for r in b.conditional_formatting._cf_rules.values())
    print(f"{n:22} {dv_a:4}->{dv_b:<4} {cf_a:4}->{cf_b:<4} {len(a.merged_cells.ranges):4}->{len(b.merged_cells.ranges):<4}")
