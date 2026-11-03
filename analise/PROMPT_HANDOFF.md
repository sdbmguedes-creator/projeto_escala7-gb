# PROMPT DE CONTINUIDADE — Escala 2.0 (7º GB / CBMPE)

> Copie o bloco abaixo inteiro para o novo agente e anexe os arquivos indicados no fim.

---

## 1. CONTEXTO

Você vai trabalhar na **Escala de Serviço do 7º Grupamento de Bombeiros (7º GB) do CBMPE**,
que cobre a sede administrativa e quatro subunidades (1ªSB Carpina, 2ªSB Goiana,
3ªSB Macaparana, 4ªSB Surubim).

Existe um modelo antigo em uso (Google Sheets), funcional mas degradado, e um **modelo novo
(2.0) em construção**, já com 12 abas validadas. Seu trabalho é **continuar o 2.0**, não
refazer o antigo.

## 2. OBJETIVO E LIMITES

- Construir/evoluir a planilha 2.0 em **Google Sheets**, entregue como `.xlsx` para importação.
- **Não usar Excel como ambiente.** O `.xlsx` é apenas veículo: gera-se, importa-se uma vez
  no Drive e o trabalho passa a acontecer no Sheets.
- **Não usar Apps Script nesta fase** (decisão do responsável). Script só depois do aceite.
- Manter a planilha **leve** (o modelo antigo tem 1,48 MB; o 2.0 tem ~65 KB) e evitar
  fórmula volátil repetida célula a célula.

## 3. ESTRUTURA (12 abas)

| Aba | Papel |
|---|---|
| `PARAMETROS` | competência (`B1` = data), metas por SB, códigos, cotas PJES do mês, ordem das escalas |
| `PESSOAS` | ID (matrícula), posto, nome de guerra, **ORDEM** (antiguidade), nome completo (fórmula) |
| `MES` | ID, posto, nome, seção, código, escala, ordem da escala, ADM, motorista, ordem, chave |
| `FERIADOS` | nacionais + Pernambuco |
| `GRADE 1ªSB` | operacionais (3–34) + **ADM (36–49)** + visitantes (52–63) + totais (65–67) |
| `GRADE 2ªSB/3ªSB/4ªSB` | mesma estrutura, sem bloco ADM preenchido |
| `DIARIO` | BM/dia, motoristas/dia, meta, déficit |
| `PJES` | cotas recebidas (SEI), utilizadas, saldo, situação |
| `ACO` | oficiais (linhas 3–4), militares (7–54), marcação `12`/`24` |
| `INTEGRACAO` | painel por militar: dias na SB, visitante, ACO, alerta |

## 4. REGRAS DE NEGÓCIO

**Códigos de marcação** (dropdown, recusa texto):

| Código | Meta diária | Cota PJES | Saldo ADM |
|---|---|---|---|
| `X` | 1 | — | 2 |
| `PJ` (PJES 24h) | 1 | 2 | — |
| `PJ12` (PJES 12h) | 0,5 | 1 | — |
| `PJ12H` (12h de militar de fora) | 0,5 | **0** | — |
| `S1` (12h de saldo) | 0,5 | — | 1 |
| `SPJ` (saldo + 12h PJES) | 1 | 1 | 1 |
| `SB` (PJES de fora, não conta) | 0 | 0 | 0 |
| `N` (indisponível) | 0 | 0 | 0 |

- **ADM**: cumpre **4 saldos/mês**, **não segue 1x3**, serviços avulsos.
- **Demais**: **7–8 serviços ordinários** em escala 1x3 (rodízio ALFA→BRAVO→CHARLIE→DELTA).
- **Ordem permitida**: `X` seguido de `PJ` pode; `PJ` seguido de `X` **não** (12h de descanso).
- **Hierarquia de lançamento**: SB de origem → SB host → ACO. Camada de cima prevalece.
- Serviço do visitante é lançado **na grade do host**.

## 5. DECISÕES FECHADAS (não reabrir)

- antiguidade = **ordem da lista**, editável; saída de militar = **apagar a linha**;
- escala e ADM são **mensais** (a DOp define); não há escala fixa;
- virada do mês = **copiar a planilha**; **um arquivo por mês**;
- cotas PJES vêm do **processo SEI**, transcritas pela DOp na `PARAMETROS`;
- `.xlsx` é veículo de importação, **não** ambiente de trabalho;
- afastamento parcial: no 2.0 é **dado** (não deduzir do `N`);
- script de virada: **adiado** até o aceite.

## 6. ARMADILHAS TÉCNICAS (custaram horas no projeto original)

1. **Locale pt-BR**: separador de argumentos `;`, decimal `,`. O Sheets **traduz nomes de
   função** (`SE`, `CORRESP`, `ÍNDICE`, `SEERRO`, `CONT.SE`, `SOMARPRODUTO`).
2. **`PROCX`/`XLOOKUP`**: o 6º argumento aceita `1` (binário, padrão), `-1`, `2`, `-2`.
   **`0` não existe no Sheets e gera `#VALUE`.** Nunca usar sem modo; prefira `INDEX`+`MATCH`.
3. **Busca binária em lista desordenada** devolve resultado errado silenciosamente. Sempre
   `MATCH(...,0)`.
4. **Chaves de matriz `{...}`** importadas do `.xlsx` podem falhar na conversão. Preferir
   `COUNTIF` somados a `{"x";"PJ"}`.
5. **`FILTER`/`SORTBY`** importados podem quebrar. Usar `MATCH` com chave pré-calculada.
6. **Formatação condicional não lê outra aba.** Para alerta entre abas, criar célula/linha
   auxiliar **na mesma aba** e a CF lê essa célula.
7. **Caractere `ª`** dentro de texto de fórmula pode virar entidade XML e não casar. Usar
   chaves ASCII (`SB4#1` em vez de `4ªSB#1`).
8. **Acentuação em PYTHON**: nunca editar `.py` com PowerShell `Set-Content` (corrompe UTF-8).

## 7. ESTADO ATUAL (validado pelo responsável)

Tudo abaixo foi testado e aprovado no Sheets:

- cadastro por matrícula, tabela do mês, dropdowns de seção e escala;
- grades derivadas nas 4 SBs, ordenadas por escala e, dentro dela, por antiguidade;
- cores por escala (ALFA verde, BRAVO amarelo, CHARLIE azul, DELTA cinza) + faixa por grupo;
- ADM dentro da 1ªSB, com contagem diária somando as duas categorias;
- visitantes com seleção por menu e dropdown de códigos que recusa texto;
- calendário por competência com fins de semana e feriados em vermelho;
- `DIARIO`, `PJES`, `ACO` e o painel `INTEGRACAO`.

## 8. O QUE FALTA (em ordem)

1. **Reciprocidade por dia** — mostrar, por visitante, a lista de dias já ocupados na origem
   (ex.: `3, 7, 12, 15`) e espelhar na origem os dias em que o militar serve fora.
   *Use uma célula com `TEXTJOIN` por visitante, não uma grade auxiliar de 1.488 fórmulas.*
2. **ACO no `DIARIO`**.
3. **Painel de conferências**: serviço fora da escala do dia, `PJ` seguido de `X`,
   militar em duas camadas no mesmo dia, cadastro duplicado.
4. **Aceite**: rodar novembro no 2.0 e comparar com o modelo atual — cotas (326 de referência),
   BM/dia, motoristas/dia e saldo por militar, **diferença zero**.

## 9. CRITÉRIO DE ACEITE

O 2.0 só substitui o modelo atual se **reproduzir os totais de um mês fechado com diferença
zero**. Se não reproduzir, o modelo está errado — não o dado.

## 10. ARQUIVOS PARA ANEXAR

| Arquivo | Para quê |
|---|---|
| `ESCALA_2.0_prototipo.xlsx` | base do 2.0 (12 abas) — **ponto de partida** |
| `BASE_Escala_de_SERVIÇO.xlsx` | modelo antigo de outubro (referência de regras) |
| `11 - Escala de Novembro 2026.xlsx` | mês em curso (referência de números) |
| `8` e `9 - Escala de ... 2026.xlsx` | agosto e setembro (variações reais) |
| `REGISTRO.md` / `DESENHO_2.0.md` | decisões e desenho |
| `PROMPT_HANDOFF.md` | este documento |

## 11. COMO ENTREGAR

- gerar `.xlsx` e pedir **importação como planilha nova** (nunca substituir a de trabalho);
- entregar sempre com **lista de verificação numerada**, no formato:
  `[1] estrutura · [2] dados · [3] fórmulas · [4] comportamento · [5] regressão`;
- gravar **checkpoint** (cópia + manifesto SHA-256) após cada bloco aprovado.
