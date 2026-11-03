# P13 — Proteção por área (5 usuários)

> Objetivo: impedir que um comandante apague ou altere fórmulas de total acidentemente. 5 pessoas editam a mesma planilha; precisa de cerca.

---

## Estratégia — 2 etapas

**Etapa 1 (agora):** proteger com "mostrar aviso ao editar". Não trava, apenas avisa. Testa com os 4 comandantes por um ciclo.
**Etapa 2 (depois de estabilizar):** bloquear de fato, liberando edição apenas nas células de marcação e cadastro.

---

## O que proteger (por aba)

| Aba | Intervalos a proteger | Motivo |
|-----|----------------------|--------|
| `RESUMO` | Colunas de fórmula (F, G, H, I, J…O, P) | Cadastro e fórmulas de saldo/PJES |
| `1ªSB` | Linhas 29-31 (totais BM/motoristas/PJES), coluna AL/AM (saldos), linha 49 (PJES ADM), linha 92-93 (totais externos) | Fórmulas de total |
| `2ªSB` | Linhas 30-32, coluna AK/AL, linha 53-54 | idem |
| `3ªSB` | Linhas 21-23, coluna AK/AL, linha 44-45 | idem |
| `4ªSB` | Linhas 32-34, coluna AK/AL/AM, linha 56-57 | idem |
| `ACO` | Linhas de totais e fórmulas de oficiais | Fórmulas de turno |
| `DIÁRIO` | Toda a aba (é derivada) | Não deve ser editada à mão |
| `PJES DISTRIBUIÇÃO` | Colunas de fórmula de rateio | Cotas PJES |
| `ESCALAS 2026` | Tabela anual (fórmula, P10) | Não alterar |

---

## O que deixar livre (edição permitida)

| Aba | Intervalos livres |
|-----|------------------|
| Todas as SBs | Colunas de marcação dos dias (G:AK, dias 1-30/31) — é onde se lança serviço |
| `RESUMO` | Coluna E (seção), H (escala), G (mot.) — cadastro mensal; P (ADM DA SB) |
| `ACO` | Marcações de oficiais/turnos |
| `PJES DISTRIBUIÇÃO` | Cotas distribuídas por SB (entrada manual da DOp) |

---

## Como aplicar no Google Sheets

1. Selecionar o intervalo a proteger.
2. Direito → Proteger intervalo.
3. Definir permissão: "Mostrar aviso ao editar" (etapa 1) ou restringir a usuários específicos (etapa 2).
4. Dar nome descritivo ao intervalo protegido (ex.: `Totais_1SB`).

---

## Treinamento

Os 4 comandantes + você precisam saber: células de marcação são livres; o resto avisa ou trava. Incluir na `COMO USAR` (P4).

---

## Arquivos afetados

| Item | Onde |
|------|------|
| Proteções por intervalo | Todas as SBs + RESUMO + ACO + DIÁRIO + PJES DISTRIBUIÇÃO + ESCALAS 2026 |

---

**Status:** especificação pronta. Aplicar no Google Sheets, começando pela etapa 1 (aviso ao editar).
