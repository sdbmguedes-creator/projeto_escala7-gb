# P11 — Descanso (24h dias seguidos) + Dupla escala (militar em 2 SBs)

> Objetivo: sinalizar conflito operacional (descanso violado) e militar lançado em duas subunidades no mesmo dia. Tudo via formatação condicional nativa — sem bloqueio, apenas aviso visual.

---

## 1. Descanso — 24h em dias consecutivos

Regra: um militar com serviço `X` (24h) ou `PJ` (PJES 24h) em dia D não deve ter serviço `X` ou `PJ` no dia D+1.

### Fórmula de CF (aplicar no intervalo de marcação de cada militar)

Assumindo que as marcações de um militar estão na linha 3, colunas G:AK (dias 1 a 30):

```
=E($G3<>""; OU($G3="X"; $G3="PJ"); OU($H3="X"; $H3="PJ"))
```

Aplicar essa fórmula em cada célula da grade de marcação (G3:AK3, G4:AK4, etc.), destacando o dia D quando há 24h em D e em D+1.

### Variante para o último dia do mês

No último dia (coluna AK), comparar com o primeiro dia do mês seguinte — não tratável nativamente sem dados do mês seguinte. Deixar sem sinalizar no último dia (aceitável; o próximo mês corrige).

### Escopo
Aplicar a cada linha de militar nos blocos pessoal (1ªSB a 4ªSB) e bloco ADM da 1ªSB.

---

## 2. Dupla escala — militar em 2 SBs no mesmo dia

O Google Sheets **não permite CF entre abas diferentes** diretamente. Solução: célula auxiliar por dia por SB que conta em quantas SBs o militar aparece, ou um painel de conferência separado.

### Abordagem A — célula auxiliar por SB (recomendada para uso diário)

Em cada aba SB, criar uma área auxiliar (oculta) que, para cada dia, lista os militares que têm serviço (`X`, `PJ`, `PJ12`, `SPJ`) naquele dia. Isso já existe parcialmente na linha de totais.

Para detectar duplicidade entre SBs, criar no `DIÁRIO` (ou em uma aba auxiliar `CONFERENCIA`) uma matriz:

| Dia | Militar | SB 1 | SB 2 | SB 3 | SB 4 | Total SBs |
|-----|---------|------|------|------|------|-----------|

Fórmula para cada célula (ex.: militar TELES, dia 5, SB 1ªSB):
```
=SE(ES(1ªSB!G<linhaTELES>="X"; ... );1;0)
```
Complexo de manter por linha.

### Abordagem B — painel de conferência por nome (mais simples)

Criar em `CONF. INDIV.` ou nova aba `DUPLA_ESCALA` uma consulta: escolher um militar via dropdown, ver em quais SBs e dias ele tem serviço no mês.

```
=FILTRO(... )  — ou INDEX/CORRESP sobre cada aba SB
```

Como `FILTRO` pode não sobreviver à exportação .xlsx, preferir `ÍNDICE`+`CORRESP` com chave pré-calculada.

### Abordagem C — CF com referência a célula auxiliar na MESMA aba (compatível com limitação do Sheets)

Na aba `DIÁRIO`, adicionar para cada dia uma célula auxiliar que concatena os militares com serviço em cada SB. Então uma regra de CF em cada SB compara o nome do militar contra essa lista.

Devido à complexidade e ao fato de que a duplicidade militar entre SBs já é capturada pela função `conflito()` do app HTML (que roda em tempo real), recomendar:

- **Implementar agora:** apenas a regra de descanso (CF simples, intra-aba).
- **Deixar a dupla escala para o painel de conferências (P12)**, que cruza dados entre abas de forma estruturada.

---

## 3. Arquivos afetados

| Item | Onde |
|------|------|
| CF descanso | intervalo de marcação de cada militar (4 SBs + ADM) |
| Painel dupla escala | nova aba `DUPLA_ESCALA` ou dentro de `CONF. INDIV.` (P12) |

---

**Status:** especificação pronta. Descanso aplicável imediatamente via CF; dupla escala fica para o painel de conferências.
