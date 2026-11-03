# P14 — Visual: legenda única, CF enxuta, área de impressão

> Objetivo: reduzir ~253 regras de formatação condicional (1ªSB tem 41) para um esquema único e padronizado, e definir área de impressão por SB para afixar a escala.

---

## 1. Legenda única de cores

Padronizar em todas as SBs:

| Código | Cor de fundo | Texto | Significado |
|--------|-------------|-------|-------------|
| `X` | verde | branco | Serviço ordinário 24h |
| `PJ` | azul | branco | PJES 24h |
| `PJ12` / `PJ12H` | ciano | branco | PJES 12h |
| `S1` | — | — | Saldo 12h |
| `SPJ` | verde escuro | branco | Saldo + PJES 24h |
| `N` | cinza | cinza escuro | Não disponível |
| Afastado (`X - ***`) | cinza escuro | — | Afastado no mês |
| FEM (`F`) | rosa claro | preto | Feminino (limite de gênero) |
| Conflito (dupla escala) | laranja | — | Sinalização P11 |
| Fora da escala | laranja | — | Sinalização P10 |

Criar uma célula-legenda na aba `COMO USAR` (ou no topo de cada SB) com essas cores.

---

## 2. Consolidar CF sobrepostas

Hoje: 253 regras, várias sobrepostas, 1ªSB tem 41.

Ação: revisar cada aba, identificar regras que cobrem a mesma faixa com critérios similares, e fundir em uma única regra por critério (ex.: uma regra para `X`, uma para `PJ`, uma para fora da escala, etc.). Meta: reduzir em ~40%.

Como: 
- Formatação → Regras de formatação condicional → revisar cada uma
- Manter apenas regras que tenham critério distinto
- Unir regras idênticas em diferentes faixas

---

## 3. Área de impressão por SB

Definir para cada aba de SB:

| Aba | Área de impressão |
|-----|------------------|
| `1ªSB` | `A1:AK93` |
| `2ªSB` | `A1:AN54` |
| `3ªSB` | `A1:AN45` |
| `4ªSB` | `A1:AN57` |
| `ACO` | área dos oficiais + turnos |
| `DIÁRIO` | toda a aba |

Configurar: Arquivo → Imprimir → Selecionar células → Definir como área de impressão.

---

## 4. Desocultar `ORIENTAÇÕES` (P4 — já executado no Bloco A)

A aba já foi desocultada e atualizada com o texto dos códigos, afastamento, motoristas, FEM, ordem de lançamento e rotina de abertura do mês (ver `RELATORIO_BLOCO_A.md`, §7).

---

## Arquivos afetados

| Item | Onde |
|------|------|
| Legenda única | aba `COMO USAR` / topo de cada SB |
| CF consolidada | todas as SBs (revisar ~253 regras) |
| Área de impressão | 1ªSB, 2ªSB, 3ªSB, 4ªSB, ACO, DIÁRIO |

---

**Status:** especificação pronta. Trabalho chato (revisar regras uma a uma) mas não complexo. Aplicar no Google Sheets.
