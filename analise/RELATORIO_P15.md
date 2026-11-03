# RELATÓRIO — P15 (coluna ADM DA SB + regime por seção)

**Arquivo:** `BASE_Escala_de_SERVIÇO_v3.xlsx` · **Base:** v2 (que já contém Bloco A + P1b + P9)

---

## 1. O que foi feito

| Item | Ação | Células |
|---|---|---|
| Limpeza | removidos 18 textos `"Não"` residuais da antiga coluna de voluntários | 18 |
| Cabeçalho | `RESUMO!P1` = **"ADM DA SB"** | 1 |
| Marcação | 13 ADM identificados na nova coluna | 13 |
| Fórmula | coluna `ORDINÁRIO` passou a decidir o regime pela coluna `P` | 159 |
| **Total** | | **191** |

### Quem foi marcado

| Valor em `P` | Quem | Linhas no RESUMO |
|---|---|---|
| `ADM 7ºGB` | 11 praças do administrativo da sede | 2, 5, 10, 14, 15, 19, 22, 30, 55, 77, 91 |
| `ADM 2ªSB` | CB JOÃO LIANDA | 63 |
| `ADM 3ªSB` | 3º SGT ANDRADE | 36 |
| *(em aberto)* | **4ªSB — ver §4** | — |

### A fórmula (exemplo, 1ªSB)

```
=IF(XLOOKUP($D3,RESUMO!$F$2:$F$400,RESUMO!$P$2:$P$400,"",0,0)<>"",
    COUNTIF($G3:$AK3,"X")*2+COUNTIF($G3:$AK3,"SPJ")+COUNTIF($G3:$AK3,"S1"),
    COUNTIF($G3:$AK3,"X"))
```

---

## 2. Dois erros que eu quase embuti (e a verificação pegou)

**2.1 O nome do militar é fórmula, não texto.**
`RESUMO!F` contém `=TEXTJOIN(...,B2," ",C2," ",D2)`. Na primeira versão eu comparava o conteúdo da célula com a lista de nomes, não casou nada, e os 13 ADM caíram todos no valor padrão. Corrigido montando o nome a partir de `B+C+D`.

**2.2 A coluna `P` não estava vazia.**
Apesar de o Bloco A ter esvaziado as fórmulas `#REF!`, restaram **18 células com o texto `"Não"`** (marca de "não é voluntário PJES"). Com a fórmula nova, esses 18 não-ADM seriam tratados como ADM e teriam o saldo **dobrado**. Corrigido limpando a coluna antes de marcar.

> Isso é argumento para manter a disciplina de conferência: mudança estrutural sem checagem de resíduo passa silenciosa.

---

## 3. Prova de neutralidade (nenhum número mudou)

Para cada uma das 159 células reescritas, comparei o regime antigo (presença de `×2` na fórmula) com o novo (presença de valor em `P`):

- células cuja fórmula antiga era de **saldo** → têm `P` preenchido ✔
- células cuja fórmula antiga era **contagem simples** → `P` vazio ✔

**Resultado: 0 divergências** — exceto 3 casos conhecidos e inofensivos:

| Célula | Situação | Valor antes | Valor depois |
|---|---|---|---|
| `1ªSB!AL58` | ANDRADE lançado como visitante na 1ªSB | 0 | 0 |
| `2ªSB!AK42` | ANDRADE visitante na 2ªSB | 0 | 0 |
| `4ªSB!AK40` | FLÁVIO visitante na 4ªSB | 0 | 0 |

Eles são ADM e passaram a usar a fórmula de saldo na linha de visitante — como não têm marcação, o valor continua 0.

Condicionais, validações e merges: **idênticos** à v2.

---

## 4. Pendência: quem é o ADM da 4ªSB?

Sua frase — *"a linha referente ao ADM daquela SB foi mudada e o ERALDIR ficou nela, isso é um problema"* — indica que o ERALDIR **não é** o ADM da 4ªSB; ele apenas ocupou a linha.

Por isso ele ficou **sem marcação** em `P` (comportamento igual ao de hoje). Consequências:
- se ele **for** o ADM da 4ªSB → marcar `ADM 4ªSB` e o saldo dele passa a contar `×2`;
- se ele **não for** → o ADM da 4ªSB está em outro lugar (ou não foi lançado), e é preciso localizá-lo.

**Preciso que você diga quem é o ADM da 4ªSB.**

---

## 5. Aplicação no Google Sheets

> Lembrete: no Sheets os nomes das funções são em inglês; o separador é `;` no seu locale.

**Passo 1 — limpar a coluna `P`**
- Selecione `RESUMO!P2:P111` → `Delete` (isso remove os 18 "Não").
- Em `P1`, digite `ADM DA SB`.

**Passo 2 — marcar os 13 ADM**

| Linha | Valor |
|---|---|
| 63 | `ADM 2ªSB` |
| 36 | `ADM 3ªSB` |
| 2, 5, 10, 14, 15, 19, 22, 30, 55, 77, 91 | `ADM 7ºGB` |

**Passo 3 — trocar a fórmula da coluna `ORDINÁRIO`**

Primeira célula de cada bloco, depois arrastar até o fim:

| Aba | Coluna | Blocos |
|---|---|---|
| 1ªSB | `AL` | 3–29 · 39–49 · 55–92 |
| 2ªSB | `AK` | 3–30 · 39–53 |
| 3ªSB | `AK` | 3–18 · 27–41 |
| 4ªSB | `AK` | 3–31 · 40–55 |

Fórmula (ajuste a faixa conforme o bloco):

```
=SE(XLOOKUP($D3; RESUMO!$F$2:$F$400; RESUMO!$P$2:$P$400; "")<>"";
    COUNTIF($G3:$AK3;"X")*2+COUNTIF($G3:$AK3;"SPJ")+COUNTIF($G3:$AK3;"S1");
    COUNTIF($G3:$AK3;"X"))
```

> ⚠️ Ao arrastar, as 11 células desalinhadas da 2ªSB são normalizadas junto — **corrige** a subcontagem, mas **muda número**. Continua pendente a sua decisão sobre isso.

---

## 6. Próximo

**P5 — faixa única de cadastro** (aproximadamente 700 células apontando para `RESUMO!$F$2:$H$400`). É o que impede a cópia mensal de voltar a degradar o modelo.
