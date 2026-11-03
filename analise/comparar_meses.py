"""Compara a estrutura e os numeros dos quatro meses."""
from collections import Counter
from openpyxl import load_workbook

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}

for nome, caminho in ARQS.items():
    wb = load_workbook(caminho)
    wv = load_workbook(caminho, data_only=True)
    res = wv["RESUMO"]
    try:
        data1 = wv["ACO"]["A3"].value
    except Exception:
        data1 = None
    pessoas = 0
    secoes = Counter()
    escalas = Counter()
    for r in range(2, 160):
        sec = res.cell(r, 5).value
        nome_m = res.cell(r, 6).value
        if isinstance(nome_m, str) and isinstance(sec, str) and " " in nome_m:
            pessoas += 1
            secoes[sec] += 1
            escalas[res.cell(r, 8).value] += 1
    erros = 0
    for ws in wv.worksheets:
        erros += sum(1 for row in ws.iter_rows() for c in row
                     if isinstance(c.value, str) and c.value.startswith("#"))
    print(f"===== {nome} =====")
    print(f"  abas ({len(wb.sheetnames)}): {wb.sheetnames}")
    print(f"  data-base (ACO!A3): {data1}")
    print(f"  pessoas no RESUMO: {pessoas}")
    print(f"  por secao: {dict(secoes)}")
    print(f"  escalas: {dict(escalas)}")
    print(f"  erros (#): {erros}")
    print(f"  S9 (cotas): {res['S9'].value}  | S4:S8 = {[res[f'S{r}'].value for r in range(4,9)]}")
    print()
