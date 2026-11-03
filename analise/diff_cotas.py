from openpyxl import load_workbook

OLD = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
NEW = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"


def cotas(path):
    wv = load_workbook(path, data_only=True)
    res = wv["RESUMO"]
    out = {}
    for r in range(2, 112):
        nome, sec = res.cell(r, 6).value, res.cell(r, 5).value
        if not (isinstance(nome, str) and isinstance(sec, str)):
            continue
        tot = 0
        detalhe = []
        for tab, colval, a, b in [("ACO", 37, 6, 54), ("1ªSB", 39, 3, 120),
                                  ("2ªSB", 38, 3, 120), ("3ªSB", 38, 3, 120),
                                  ("4ªSB", 38, 3, 120)]:
            ws = wv[tab]
            for rr in range(a, b + 1):
                if ws.cell(rr, 4).value == nome:
                    v = ws.cell(rr, colval).value
                    v = v if isinstance(v, (int, float)) else 0
                    if v:
                        detalhe.append(f"{tab}={v}")
                    tot += v
                    break
        out[nome] = (sec, tot, detalhe)
    return out


a = cotas(OLD)
b = cotas(NEW)
print(f"total antigo: {sum(v[1] for v in a.values())} | total novo: {sum(v[1] for v in b.values())}")
print()
print("=== quem mudou ===")
for nome in b:
    va, vb = a.get(nome, (None, 0, [])), b[nome]
    if va[1] != vb[1]:
        print(f"  {nome[:32]:34} {va[1]} -> {vb[1]}  {va[2]} -> {vb[2]}")
print()
print("=== ADERBAL: detalhe ===")
for nome in b:
    if "ADERBAL" in nome:
        print(" antigo:", a.get(nome), "\n novo  :", b[nome])
