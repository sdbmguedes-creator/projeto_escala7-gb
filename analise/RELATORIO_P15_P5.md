# RELATÓRIO — P15 + P5

**Arquivo:** `BASE_Escala_de_SERVIÇO_v4.xlsx` (cumulativo: Bloco A + P1b + P9 + P15 + P5)

---

## P15 — ADM DA SB (191 células)

- `RESUMO!P1` = "ADM DA SB"; 18 textos `"Não"` residuais limpos; 13 ADM marcados.
- Coluna `ORDINÁRIO` passou a decidir o regime por `P`, não pela posição da linha.
- **Neutralidade provada:** 0 divergências entre regime antigo e novo.
- 3 células de ADM lançados como visitantes em outra SB passaram a usar a fórmula de saldo — valiam 0 e continuam 0.

### Erros meus que a conferência pegou
1. `RESUMO!F` é fórmula (`TEXTJOIN`), não texto → a primeira versão não casou nenhum nome.
2. A coluna `P` não estava vazia: 18 células com o texto `"Não"`. Sem limpar, 18 não-ADM teriam o saldo **dobrado**.

**Pendente:** quem é o ADM da 4ªSB (o ERALDIR ficou sem marcação, comportamento igual ao de hoje).

---

## P5 — faixa única de cadastro (456 células)

Todas as referências do tipo `RESUMO!F2:H189` (e variantes `H176`…`H229`) foram substituídas por **`RESUMO!$F$2:$H$400`**.

| Aba | Células |
|---|---|
| RESUMO | 111 |
| 1ªSB | 111 |
| 4ªSB | 72 |
| 2ªSB | 68 |
| ACO | 48 |
| 3ªSB | 46 |
| **Total** | **456** |

Também os literais de `INDIRECT` no RESUMO: `'!$D3:$D32` → `'!$D$3:$D$120`, idem `AL` e `AK` (444 ocorrências).

**Verificação:** 0 faixas antigas remanescentes; validações e formatações condicionais idênticas.

**Por que é seguro:** o cadastro tem 110 pessoas e nomes únicos; ampliar a faixa até a linha 400 não muda o primeiro resultado encontrado. O ganho é que o efetivo pode crescer até ~290 pessoas sem tocar em fórmula — e a cópia mensal deixa de gerar variantes.

---

## Achado novo na P5: busca binária no RESUMO

`RESUMO!I` (coluna ORD) usa `XLOOKUP(..., 0, 1)` — **modo 1 = busca binária**, que exige lista ordenada. A coluna `F` (nome) está em ordem de antiguidade, não alfabética.

Consequência: o resultado dessa coluna é **arbitrário** — pode achar, não achar ou achar a pessoa errada. Como ao ampliar a faixa o tamanho mudou, os valores podem ter se alterado.

Mitigação proposta (**P5b**): trocar `,0,1)` por `,0,0)` nas fórmulas do RESUMO!I — busca linear, resultado correto. São 110 células, mudança mecânica.

Impacto: nenhum cálculo consome `RESUMO!I` hoje, então é seguro.

---

## Aplicação no Sheets

### P15
1. `RESUMO!P2:P111` → `Delete`; em `P1` escrever `ADM DA SB`.
2. Marcar: linha 63 = `ADM 2ªSB`; linha 36 = `ADM 3ªSB`; linhas 2, 5, 10, 14, 15, 19, 22, 30, 55, 77, 91 = `ADM 7ºGB`.
3. Na coluna `ORDINÁRIO` (AL na 1ªSB, AK nas demais), na primeira linha de cada bloco:

```
=SE(XLOOKUP($D3; RESUMO!$F$2:$F$400; RESUMO!$P$2:$P$400; "")<>"";
    COUNTIF($G3:$AK3;"X")*2+COUNTIF($G3:$AK3;"SPJ")+COUNTIF($G3:$AK3;"S1");
    COUNTIF($G3:$AK3;"X"))
```

### P5
- Substituir, em toda a planilha, `RESUMO!F2:H189` (e as variantes) por `RESUMO!$F$2:$H$400`.
  No Sheets há **Localizar e substituir** com "também nas fórmulas" — dá para fazer em poucos passos, uma variante por vez.
- No RESUMO, trocar nos `INDIRECT`: `$D3:$D32` → `$D$3:$D$120`; `$AL3:$AL32` → `$AL$3:$AL$120`; `$AK3:$AK32` → `$AK$3:$AK$120`.

---

## Sequência restante

| Item | Situação |
|---|---|
| Bloco A · P1b · P9 · P15 · P5 | **executados** |
| P5b (busca linear no RESUMO) | proposta |
| P6 (blocos com reserva) | pendente |
| P7 (rotina de virada do mês) | pendente |
| P10–P12 (conferências: escala, descanso, unicidade) | pendente |
| P13–P14 (proteção, visual) | pendente |
