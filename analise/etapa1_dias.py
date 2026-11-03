"""Etapa 1: lista de dias ocupados na origem, no bloco de visitantes de cada SB."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.formula import ArrayFormula

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx"
wb = load_workbook(P)

hdr = Font(bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="8B4513")

VIS_S, VIS_N = 52, 12
OP_N = 32

ABAS = ["GRADE 1ªSB", "GRADE 2ªSB", "GRADE 3ªSB", "GRADE 4ªSB"]

for aba in ABAS:
    ws = wb[aba]
    # cabecalho da coluna AL
    c = ws["AL51"]
    c.value = "DIAS NA ORIGEM"
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["AL"].width = 26

    for i in range(VIS_N):
        r = VIS_S + i
        # AM: secao do militar no MES
        sec = (f'IFERROR(INDEX(MES!$D$2:$D$200,'
               f'MATCH(INDEX(PESSOAS!$A$2:$A$200,MATCH($D{r},PESSOAS!$E$2:$E$200,0)),'
               f'MES!$A$2:$A$200,0)),"")')
        ws[f"AM{r}"].value = f"={sec}"
        # AN: linha dele na grade de origem
        ws[f"AN{r}"].value = (
            f'=IFERROR(MATCH($D{r},INDIRECT("\'"&$AM{r}&"\'!$D$3:$D${OP_N+2}"),0),0)'
        )
        # AL: dias ocupados na origem
        texto = (
            f'=IFERROR(TEXTJOIN(", ",TRUE,'
            f'IF(INDEX(INDIRECT("\'"&$AM{r}&"\'!$G$3:$AK${OP_N+2}"),$AN{r},0)<>"",'
            f'INDIRECT("\'"&$AM{r}&"\'!$G$2:$AK$2"),"")),"")'
        )
        ws[f"AL{r}"] = ArrayFormula(ref=f"AL{r}", text=texto)

    ws.column_dimensions["AM"].hidden = True
    ws.column_dimensions["AN"].hidden = True

wb.save(P)
print("coluna DIAS NA ORIGEM criada nas 4 grades")
