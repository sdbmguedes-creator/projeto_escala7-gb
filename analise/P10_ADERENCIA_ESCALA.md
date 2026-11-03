# P10 — Aderência à escala do dia + ESCALAS 2026 por fórmula

> Objetivo: sinalizar marcação de serviço em dia fora da escala do militar, e acabar com a remontagem manual da tabela anual de escalas.

---

## Parte 1 — Gerar `ESCALAS 2026` por fórmula

Substituir os valores estáticos da tabela (colunas B a AF, linhas 6 a 17) por fórmula derivada de uma **célula-âncora**.

### Célula-âncora
Criar em `PARAMETROS!B1` (ou similar) a data-âncora do ciclo, ex.: `01/11/2026` (dia em que a escala é BRAVO).

### Fórmula para cada célula da tabela (exemplo para B6)
Locale pt-BR (separador `;`):

```
=ESCOLHER(MOD(DATA(2026;MES($A6);COLUNA()-COLUNA($B$6))-DATA(2026;11;1);4)+1;"BRAVO";"CHARLIE";"DELTA";"ALFA")
```

Explicação:
- `COLUNA()-COLUNA($B$6)` = 0 para B, 1 para C, … (dia do mês dentro da linha)
- `DATA(2026;MES($A6);<dia>)` monta a data real daquela célula
- `MOD(data - âncora; 4) + 1` dá 1..4 (posição no ciclo)
- `ESCOLHER` mapeia 1=BRAVO, 2=CHARLIE, 3=DELTA, 4=ALFA

Aplicar essa fórmula em toda a matriz B6:AF17 (12 meses × 31 dias), ajustando o ano se necessário. O ano 2026 vem de `PARAMETROS` ou de uma célula fixa.

Validação já feita: com âncora 01/11/2026 = BRAVO, a fórmula reproduz exatamente os 30 dias de novembro atuais (linha 16 da tabela).

---

## Parte 2 — Linha auxiliar por SB (escala do dia)

Em cada aba SB, adicionar uma linha auxiliar (oculta) logo abaixo do cabeçalho de dias, que calcula a escala vigente para cada dia do mês a partir de `ESCALAS 2026`.

Exemplo para 1ªSB, assumindo que a linha de dias é a linha 2 (G2:AK2 com 1..30):

```
=ÍNDICE('ESCALAS 2026'!$B$6:$AF$17; CORRESP(MES(DATA(2026;11;1)); 'ESCALAS 2026'!$A$6:$A$17; 0); DIA(G$2))
```

- `CORRESP(mes; nomes dos meses; 0)` acha a linha do mês
- `DIA(dia)` acha a coluna do dia dentro da linha
- Resultado: a escala (ALFA/BRAVO/CHARLIE/DELTA) vigente naquele dia

Essa linha auxiliar fica em uma linha fixa (ex.: linha 2 ou uma linha oculta acima da grade).

---

## Parte 3 — Formatação condicional: marcação fora da escala

Regra de CF aplicada ao intervalo de marcação de cada militar (colunas G:AK do bloco pessoal):

```
=E($G3<>""; $G3<>"N"; $G3<>"PJ"; $G3<>"PJ12"; $G3<>"PJ12H"; $G3<>"SPJ"; $G3<>$escala_do_dia)
```

Onde `$escala_do_dia` refere-se à célula da linha auxiliar para aquele dia.

Sinalizar (ex.: fundo laranja) quando o militar tem marcação de serviço ordinário (`X`) ou PJES em dia que não é a escala dele.

**Conservador:** sinalizar só `X` no início. `PJ`/`PJ12`/`SPJ` podem cair fora da escala por natureza (PJES avulso), então não sinalizar esses inicialmente. Calibrar após um mês de uso.

---

## Arquivos afetados

| Item | Onde |
|------|------|
| Célula-âncora | `PARAMETROS!B1` (criar) |
| Tabela anual | `ESCALAS 2026`!B6:AF17 (substituir estáticos por fórmula) |
| Linha auxiliar | +1 linha oculta por SB (1ªSB a 4ªSB) |
| CF | +1 regra por SB, no intervalo de marcação |

---

**Status:** especificação pronta. Aplicar no Google Sheets (não no .xlsx).
