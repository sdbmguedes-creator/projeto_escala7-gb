# ROTINA DE ABERTURA DO MÊS — Escala 7º GB / CBMPE

> Aplicar a cada virada de mês. Base: `ACO!A3` (primeiro dia do mês) é a **única célula de competência**; tudo deriva dela.
> Não copiar a planilha inteira — copiar **abas** e seguir os 7 passos abaixo.

---

## Passo 1 — Duplicar as abas

No Google Sheets, clique com o botão direito em cada aba e **Duplicar**:

| Aba a duplicar | Nome da cópia |
|---|---|
| `1ªSB` | `1ªSB_<mes>_<ano>` (ex.: `1ªSB_11_2026`) |
| `2ªSB` | idem |
| `3ªSB` | idem |
| `4ªSB` | idem |
| `ACO` | idem |
| `RESUMO` | idem |
| `DIÁRIO` | idem |
| `PJES DISTRIBUIÇÃO` | idem |
| `CONF. INDIV.` | idem |
| `Planilha de PJES` | idem |

As abas `PARAMETROS`, `PESSOAS`, `MES`, `FERIADOS`, `ESCALAS 2026`, `Oficiais`, `Base de Dados` são **comuns a todos os meses** — não duplicar.

---

## Passo 2 — Lançar a competência

Na cópia da aba `ACO`, preencher `A3` com o primeiro dia do mês (ex.: `01/12/2026`).

Essa célula é a âncora: `DIÁRIO` e as grades derivam a linha de datas dela.

---

## Passo 3 — Conferir a tabela `ESCALAS 2026`

Verificar que o ciclo ALFA→BRAVO→CHARLIE→DELTA bate com o novo mês. A tabela anual já existe; apenas conferir o primeiro dia do mês.

---

## Passo 4 — Limpar as marcações antigas

Selecionar o intervalo de marcação (dias do mês) de cada grade e apertar `Delete` (apenas conteúdo, não formatação):

| Aba | Intervalo de marcação |
|---|---|
| `1ªSB` | `G3:AK28` (pessoal), `G38:AK48` (ADM), `G54:AK91` (visitantes) |
| `2ªSB` | `G3:AK29` (pessoal), `G38:AK52` (visitantes) |
| `3ªSB` | `G3:AK20` (pessoal), `G29:AK43` (visitantes) |
| `4ªSB` | `G3:AK31` (pessoal), `G40:AK55` (visitantes) |
| `ACO` | marcarções dos oficiais/turnos |

Não apagar as linhas de totais (BM, motoristas, PJES).

---

## Passo 5 — Reordenar o efetivo por antiguidade

Na cópia do `RESUMO`, revisar a ordem das praças (coluna de antiguidade/ordem). Inserir novos militares na próxima linha vazia do bloco, nunca no meio. Usar `TRIM` para evitar espaços sobrando no nome.

---

## Passo 6 — Revisar as cotas PJES

Na `PJES DISTRIBUIÇÃO`, lançar o total recebido no SEI do mês e distribuir proporcionalmente entre as SBs (cada SB recebe no mínimo o maior entre Necessidade e Nec. mot.).

---

## Passo 7 — Rodar as conferências

Executar o painel de conferências e conferir:

- Diferença zero contra o mês anterior (cotas, BM/dia, motoristas/dia, saldo por militar).
- Nenhuma célula com `#REF!` ou `#N/A` inesperado.
- Nenhum militar marcado em duas SBs no mesmo dia.
- Nenhum `PJ` seguido de `X` (descanso de 12h).

---

## Regras de ouro

1. **Uma célula de competência** (`ACO!A3`); nada de data espalhada por célula.
2. **Nunca inserir linha no meio de um bloco** — usar a próxima linha vazia; os intervalos nomeados (`S1_PESSOAL`, etc.) se expandem sozinhos.
3. **`.xlsx` é veículo de importação** — gerar, importar uma vez no Drive, o trabalho passa a acontecer no Sheets.
4. **Sem Apps Script nesta fase** — tudo nativo (fórmula + formatação).
5. **Sempre trabalhar em cópia**, nunca na planilha de trabalho.

---

## Checklist de aceite (antes de publicar)

- [ ] `ACO!A3` = primeiro dia do novo mês
- [ ] Tabela `ESCALAS 2026` confere com o ciclo
- [ ] Nenhuma `#REF!` nas 4 SBs
- [ ] Diferença zero: cotas (326 ref.), BM/dia, motoristas/dia, saldo por militar vs mês anterior
- [ ] Painel de conferências sem pendências

---

**Autor:** continuação do agente anterior  
**Data:** 02/11/2026  
**Versão:** 1.0
