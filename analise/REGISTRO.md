# REGISTRO DE MELHORIAS

Atualizado em 01/11/2026, após confirmação de `RESUMO!S9 = 328`.
Fonte da verdade: **planilha no Google Sheets**.

---

## Implementado e verificado

| ID | Melhoria | Onde | Verificação |
|---|---|---|---|
| A1 | Limpeza das fórmulas da aba extinta `VOLUNTÁRIOS PJES` | `RESUMO!P`, `E` das 4 SBs, `ACO!E` | 356 `#REF!` → 0 |
| A2 | Remoção da duplicidade MONTEIRO | `1ªSB` | eliminada |
| A3 | Arquivamento de 3 abas órfãs | `FILTRO_DIAS`, `Cópia de RESUMO`, `Página30` | exportadas para `HISTORICO_2025.xlsx` |
| A4 | Manual desocultado e atualizado | `ORIENTAÇÕES` | códigos, afastamento, motoristas |
| P1b | Numeração dos visitantes + 4 órfãs | `1ªSB!A56:A92` | 42 `#N/A` → 0 |
| P9 | `PJ12` e `PJ12H` valem 0,5 na **meta diária** | BM e motoristas, 4 SBs | aplicado |
| P15 | Coluna `ADM DA SB` + regime por coluna | `RESUMO!P` e coluna `ORDINÁRIO` | neutralidade provada |
| P5 | Faixa única de cadastro | ~456 células | efetivo cresce sem tocar em fórmula |
| P5b | Busca exata no lugar de binária | `RESUMO!I` e `J:N` | sem `#VALUE` |
| — | Faixas deslocadas da 2ªSB | 11 células | 2 militares contados certo |
| ADM | **Vaga do ADM da SB na linha 3** | 4 abas | `D3` puxa do `RESUMO!P` |
| — | Deduplicação do ADERBAL na 4ªSB | linha 13 | cotas corrigidas |
| — | FERNANDA reclassificada | `ADM 7ºGB` | vaga da 1ªSB vazia |

**Estado final verificado: `RESUMO!S9 = 328`**, sem `#VALUE` e sem `#REF`.

---

## Corrigido depois de detectado em uso

| O que | Erro | Correção |
|---|---|---|
| `PJ12H` na coluna de cotas | contava cota na SB que recebe; a cota é da origem | removido da coluna de PJES |
| Modo de busca `0` | convenção do Excel; o Sheets não aceita (só `1`, `-1`, `2`, `-2`) | `ÍNDICE`+`CORRESP` |
| Coluna `P` com textos `"Não"` | teriam dobrado o saldo de 18 militares | limpeza antes da marcação |
| Marcações do ADERBAL na linha 3 | recebeu `SPJ` (rotina ADM) e passou a contar 6 cotas | restauradas: 7 `X` + 1 `PJ` = 2 |
| `1ªSB!B3/C3` e `4ªSB!B13/C13` | `#N/A` em linha de vaga vazia | `SEERRO` / limpeza |

---

## Pendente — decisão sua

| Item | O que é |
|---|---|
| `SPJ` do ADERBAL na 1ªSB (linha 73) | mesmo defeito da 4ªSB: `SPJ` em militar operacional. Vale 3 cotas hoje |
| Código `SB` em 30 células | bloco de visitantes da 1ªSB; **nenhuma fórmula conta `SB`** — marcação invisível |

## Pendente — execução

| ID | Melhoria | Situação |
|---|---|---|
| P6 | Folga de linhas por bloco | sugerido fazer no protótipo |
| P7 | Rotina de virada do mês | pendente |
| P10 | Aderência à escala do dia | pendente |
| P11 | Descanso e dupla escala | pendente |
| P12 | Unicidade e formato de cadastro | pendente |
| P13 | Proteção por área | pendente |
| P14 | Visual, CF e impressão | pendente |
| 2.0 | Protótipo paralelo | aguardando backups de ago/set/nov |

---

## Descartado / revertido

| Proposta | Situação |
|---|---|
| Afastamento parcial (colunas `AE`/`AF`, dia inicial e final) | **implementado e revertido em 01/11/2026** a pedido — a planilha voltou ao estado anterior, sem as colunas e sem a validação correspondente |
| Unificar `PJ12` e `PJ12H` (P9 original) | retirada: os dois códigos têm função distinta |
| Chave por ID numérico | adiado — `TRIM` + conferência de duplicidade resolvem por ora |

---

## Regras de contagem (consolidadas)

| Código | Meta diária | Cota PJES | Saldo (ADM) |
|---|---|---|---|
| `X` | 1 | — | 2 |
| `PJ` (24h) | 1 | 2 | — |
| `PJ12` (12h) | 0,5 | 1 | — |
| `PJ12H` (12h de militar de outra SB) | 0,5 | **0** | — |
| `S1` (12h de saldo — só ADM) | 0,5 | — | 1 |
| `SPJ` (saldo + 12h PJES — só ADM) | 1 | 1 | 1 |
| `N` | 0 | 0 | 0 |

**Regimes:** ADM cumpre **4 saldos/mês** e **não segue a escala 1x3** — pode ter serviços avulsos, sem ALFA/BRAVO/CHARLIE/DELTA fixa. Demais cumprem **7–8 serviços ordinários** em escala 1x3.

**Ordem permitida entre serviços:** `X` seguido de `PJ`/`PJ12` **é permitido**; `PJ` seguido de `X` **não é** (depois do PJES o militar precisa de 12h antes de assumir outra escala).

**`N`** = indisponível.

**Chave de busca:** sempre por nome exato (`CORRESP`), nunca por busca binária.
