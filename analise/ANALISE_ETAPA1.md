# ANÁLISE — BASE_Escala_de_SERVIÇO.xlsx (Etapa 1)

> Documento de trabalho. Nenhuma alteração foi feita na planilha.
> Arquivo analisado: `BASE_Escala_de_SERVIÇO.xlsx` (1.484.449 bytes), cópia local da planilha do Google Sheets.
> Data de referência interna: **OUTUBRO / 2026** (`DIÁRIO!C17` = "OUTUBRO", `DIÁRIO!D17` = "2026", `ACO!A3` = 01/10/2026).

---

## 0. Método e limites desta análise

**O que examinei:** as 17 abas, célula a célula — valores, fórmulas (inclusive matriz), referências entre abas, validações de dados, formatações condicionais, intervalos mesclados, estados de visibilidade, proteção e nomes definidos.

**Limites que você precisa conhecer:**

1. **Esta é uma exportação `.xlsx`, não a planilha viva.** Recursos exclusivos do Google não sobrevivem à exportação:
   - `QUERY(...)` e `FILTER(...)` aparecem como `=IFERROR(__xludf.DUMMYFUNCTION(...),"valor")` — o *resultado* foi congelado, a fórmula viva não veio;
   - resultados de fórmulas dinâmicas (spills) foram "achatados" em valores;
   - **Google Apps Script não viaja no arquivo** — nenhum script, gatilho ou menu customizado é visível aqui.
2. Por isso, tudo que eu afirmar sobre automação/script é **hipótese a confirmar**, não diagnóstico fechado.
3. Fórmulas com `QUERY`, `FILTER`, `XLOOKUP` e `INDIRECT` **não funcionam no Excel** — o modelo é Google Sheets de fato, não só de origem.

---

## 1. Identificação

| Item | Valor |
|---|---|
| Produto | Escala de serviço do **7º GB** (Grupamento de Bombeiros) — Pernambuco |
| Subunidades | ADM (7º GB) + 1ªSB Carpina, 2ªSB Goiana, 3ªSB Macaparana, 4ªSB Surubim + ACO |
| Abas | 17 (13 visíveis, 4 ocultas) |
| Células preenchidas | ~20.200 |
| Fórmulas | ~5.770 (≈1.200 em fórmula de matriz) |
| Tabelas / gráficos / imagens | nenhum |
| Nomes definidos | nenhum |
| Proteção de abas | nenhuma |
| Efetivo registrado | **110 praças**: ADM 13, 1ªSB 26, 2ªSB 27, 3ªSB 15, 4ªSB 29 |
| Distribuição por posto | 3º SGT 39, CB 28, SD 24, 2º SGT 13, ST 3, 1º SGT 3 |
| Cotas PJES do mês | **328** (294 "prontas" + 34 ACO) |
| Militares com PJES | 72 de 110 |

---

## 2. Mapa das abas

| # | Aba | Estado | Papel | Fórmulas |
|---|---|---|---|---|
| 01 | ORIENTAÇÕES | **oculta** | Manual de preenchimento (códigos, regras, responsabilidades) | 0 |
| 02 | RESUMO | visível | **Cadastro mestre** dos 110 militares + conferência de cotas + efetivo por posto | 523 |
| 03 | DIÁRIO | visível | Painel diário: BM e motoristas por dia, por SB, média do mês | 460 |
| 04 | ACO | visível | Escala da Aeronave/CO (oficiais, turnos 12h/24h) | 213 |
| 05 | 1ªSB | visível | Escala da 1ªSB (3 blocos) | 650 |
| 06 | 2ªSB | visível | Escala da 2ªSB (2 blocos) | 369 |
| 07 | 3ªSB | visível | Escala da 3ªSB (2 blocos) | 309 |
| 08 | 4ªSB | visível | Escala da 4ªSB (2 blocos) | 375 |
| 09 | PJES DISTRIBUIÇÃO | visível | Cálculo e rateio das 328 cotas PJES entre as SBs | 83 |
| 10 | CONF. INDIV. | visível | Consulta individual: todos os serviços de um militar | 389 |
| 11 | Planilha de PJES | visível | **Razão/ledger** de cada PJES lançado (data, OBM, cotas) | 750 |
| 12 | Oficiais | **oculta** | Espelho da escala dos oficiais no ACO | 151 |
| 13 | ESCALAS 2026 | visível | Tabela anual fixa: qual escala (ALFA/BRAVO/CHARLIE/DELTA) cai em cada dia | 0 |
| 14 | Base de Dados | visível | Listas derivadas por SB + contagem por posto/graduação | 450 |
| 15 | Página30 | visível | Estudo de capacidade ("fim da escala híbrida") | 18 |
| 16 | FILTRO_DIAS | **oculta** | Extração `QUERY` (congelada) — **órfã**, ninguém a referencia | 444 |
| 17 | Cópia de RESUMO | **oculta** | Backup antigo do RESUMO — **órfã**, 105 células em `#REF!` | 585 |

---

## 3. Fluxo de informação

```
        ESCALAS 2026 (tabela fixa: escala de cada dia do ano)
                      │  (referência humana — nada valida contra ela)
                      ▼
 RESUMO  ──────────────►  É a fonte da verdade:
 (cadastro: posto, mat., nome, seção, escala, PJES)
    │  ▲                        │
    │  │ VLOOKUP/XLOOKUP        │ INDIRECT por nome de SB
    ▼  │                        ▼
 1ªSB 2ªSB 3ªSB 4ªSB  ◄──── ACO (oficiais, turnos 12/24)
    │  (marcação diária X/PJ/PJ12/S1/SPJ/N)
    ├──────────────► DIÁRIO (BM e motoristas por dia + média)
    ├──────────────► RESUMO (colunas I,J..N: saldo e PJES por militar)
    ├──────────────► CONF. INDIV. (consulta individual)
    └──────────────► PJES DISTRIBUIÇÃO (cotas) ──► RESUMO (conferência)
                                    ▲
              Planilha de PJES ─────┘ (razão dos PJES lançados)
                                    ▲
                        Base de Dados (listas derivadas + efetivo)
```

**Ponto central:** todo o modelo gira em torno de **uma chave textual** — a string `"POSTO MATRÍCULA NOME"` (ex.: `2º SGT 707174-4 TELES`). Ela é o que liga RESUMO ↔ SBs ↔ ACO ↔ PJES ↔ CONF. INDIV. Não há ID numérico, não há tabela, não há nome definido.

---

## 4. Anatomia de uma aba de SB (o núcleo do modelo)

Tomando a 1ªSB como referência, cada aba tem até **6 blocos empilhados**:

| Bloco | Linhas (1ªSB) | Conteúdo |
|---|---|---|
| A. Escala da própria SB | 3–29 | militar + escala + mot. + dias 1–31 + saldo |
| B. Totais do dia | 30–36 | BM/dia vs meta, motoristas/dia vs meta, déficit, contagem por escala, limite masc./fem. |
| C. ADM | 37–51 | militares do administrativo + verificação OK/OBS + total de PJES do ADM |
| D. Outras SBs | 53–92 | militares de fora que tiram PJES na SB (só PJES) |
| E. Cabeçalho repetido | 38 | espelho da linha 2 |

**Colunas:** `A` Nº (fórmula de posição), `B` ESCALA (VLOOKUP no RESUMO), `C` MOT. (M/B/F), `D` MILITAR, `E` PJES (XLOOKUP → aba inexistente), `F` FEM, `G:AK` dias 1–31, `AL`/`AM` saldos, `AN` PJES total, `AO` observações.

**Marcações possíveis** (validação de dados): `SPJ, PJ, X, S1, PJ12/ PJ12H, N` — ver §7.

**Fórmula de BM/dia (linha 30)** — ponderação de carga:

```
SUMPRODUCT(COUNTIF(G$3:G$29,{"x";"PJ";"SPJ"}))          → vale 1
+ (COUNTIF(G$3:G$29,{"PJ12";"S1"}) * 0,5)               → vale 0,5
+ o mesmo para o bloco ADM (39:49)
+ PJES de fora (55:92) contando só PJ / PJ12*0,5
```

**Fórmula de motoristas/dia (linha 31):** idem, mas filtrando `C="M"`.

**Linha 36:** `=SUM(8-G30)` → déficit/superávit contra a meta.

---

## 5. Regras de negócio decifradas

### 5.1 Escalas e rodízio de 4 dias
Quatro escalas — **ALFA, BRAVO, CHARLIE, DELTA** — em rodízio fixo de 4 dias. A aba `ESCALAS 2026` é a tabela de referência do ano (ex.: em outubro, dia 1 = CHARLIE). Confirmei que a marcação real segue a tabela: TELES é ALFA e tem serviço nos dias 3, 7, 11… (dia 3 de outubro = ALFA ✅).

### 5.2 Códigos de marcação

| Código | Significado | Peso na carga |
|---|---|---|
| `X` | Serviço ordinário (24h) | 2 (ou 1 BM) |
| `PJ` | PJES 24h | 2 |
| `PJ12` / `PJ12H` | PJES 12h | 1 (0,5 BM) |
| `S1` | Saldo de 1º turno | 1 (0,5 BM) |
| `SPJ` | Saldo + PJES | 2 |
| `N` | Não disponível / não lançado | 0 |
| `12` / `24` | (só ACO) horas de turno do oficial | — |

### 5.3 Metas diárias (hardcoded)

| SB | Meta BM/dia | Motoristas/dia |
|---|---|---|
| 1ªSB | 8 | 2 |
| 2ªSB | 7 | — |
| 3ªSB | 4 | — |
| 4ªSB | 7 | — |

Cada meta aparece **repetida em ~30 células** (≈120 no total), como número dentro da fórmula (`=SUM(8-G30)`).

### 5.4 Limites de gênero
1ªSB + ACO: **masculino máx. 10**, **feminino máx. 4**. Existem apenas como **rótulo** — não há validação nem bloqueio; a fórmula calcula `G30-G34+ACO!F57`.

### 5.5 PJES
- Cota mensal: **328** (`PJES DISTRIBUIÇÃO!I3`), 34 para o ACO, 294 para as SBs.
- Rateio calculado por necessidade: `(meta − efetivo/4) × dias × 2 − ajustes`, depois **ajustado manualmente** (I9=127, I10=42, I11=23, I12=48).
- Conferência em duas camadas: por militar (RESUMO J:N) e por SB (RESUMO S/T → "OK / FALTA / EXCESSO").
- `Planilha de PJES` é o razão: cada lançamento com data, OBM, quantidade, e um `SUMIF` que reconcilia planilha × escala.

### 5.6 Afastamentos
Férias, Licença Especial, LTS/LM, readaptação e maternidade são lançados nominalmente em `PJES DISTRIBUIÇÃO` (linhas 21–25) e em "não voluntários para PJES" (linhas 33–37). Esse total (`R21:R25` = 16) abate o efetivo disponível e alimenta o rateio.

---

## 6. O que está funcionando bem

1. **Modelo mental coerente e completo** — cobre cadastro, escala diária, carga, PJES, afastamentos, conferência e razão. Não é uma planilha improvisada; tem lógica.
2. **Dobra de conferência em três níveis** — militar → SB → 7ºGB, com semáforo OK/FALTA/EXCESSO. É o ponto mais forte do projeto.
3. **Separação correta de responsabilidades** — cada SB preenche a própria aba; a DOp consolida.
4. **Tratamento da carga de trabalho** — a ponderação 1 / 0,5 (24h, 12h) é conceitualmente correta e aplicada de forma consistente na 1ªSB.
5. **Rastreabilidade dos PJES** — o razão permite auditar cada cota.
6. **Consulta individual** (CONF. INDIV.) — resolve a pergunta mais frequente do militar sem mexer na escala.
7. **Base de Dados derivada** — listas por SB prontas para validação, evitando digitação livre de nomes.

---

## 7. Achados técnicos (com evidência)

### 7.1 CRÍTICO — Aba `VOLUNTÁRIOS PJES` não existe mais
- **680 referências** a essa aba em fórmulas (`XLOOKUP(...,'VOLUNTÁRIOS PJES'!C:C,...)`).
- Resultado: **356 células em `#REF!`** — RESUMO (94), ACO (50), 4ªSB (29), 1ªSB e 2ªSB (30 cada), 3ªSB (18), Cópia de RESUMO (105).
- Afeta as colunas `E` das SBs (PJES do militar) e `P` do RESUMO (voluntário PJES). Ou seja: **a informação de quem é voluntário para PJES está morta no arquivo**, justamente a variável que alimenta o rateio.
- Consequência provável: o rateio em `PJES DISTRIBUIÇÃO` (K9=`E9-M34`) está usando colunas substitutas/manuais.

### 7.2 CRÍTICO — Chave = texto do nome, sem ID
- A string `"POSTO MATRÍCULA NOME"` é a chave de tudo. Um espaço a mais quebra a cadeia.
- Já existem **3 nomes com espaço à direita** no RESUMO: `"CB 718257-0 JEFERSON "`, `"SD 722011-1 AMANDA BRITO "`, `"SD 725141-6 MENEZES "`.
- Hoje funciona porque o erro está dos dois lados; qualquer um dos lados corrigido no futuro → `#N/A` silencioso em cascata.

### 7.3 ALTO — Faixas de VLOOKUP "congeladas" e divergentes
Cada célula guardou o intervalo do momento em que foi criada. Números reais:

| Aba | Variantes de fim de faixa | Intervalo |
|---|---|---|
| 1ªSB | **23 variantes** | H189 até H229 |
| 2ªSB | **27 variantes** | H189 até H217 |
| ACO | **33 variantes** | H176 até H218 |
| 3ªSB / 4ªSB | 1 (padrão H189) | — |

Isso significa: **ao inserir linhas no RESUMO, parte das células passa a enxergar o cadastro novo e parte não.** É a origem mais provável de `#N/A` e de escala errada no futuro.

### 7.4 ALTO — `COUNTIFS` do efetivo por posto também desalinhado
No bloco "Ef. Praças" (RESUMO R22:X28): `B2:B134`, `B2:B135`, `B2:B136`, `B2:B137`, `B2:B138` — e duas células usam **`B3:B137` / `B3:B138`** (começam na linha 3, ignorando o primeiro militar). A contagem por posto pode perder 1 pessoa.

### 7.5 ALTO — 2ªSB com 11 células de total deslocadas
`AK17 G17:AJ17`, `AK19 G19:AJ19`, `AL5 F5:AI5`, `AL14 F14:AH14`, `AK30 F30:AI30`, etc. — **11 células** contam 30 dos 31 dias (perdem dia 1 ou dia 31). O saldo daquele militar sai menor do que deveria, sem aviso.

### 7.6 MÉDIO — Vocabulário de códigos divergente entre abas

| Aba | Código para PJES 12h |
|---|---|
| 1ªSB | `PJ12` (33 células usadas) |
| 2ªSB, 3ªSB, 4ªSB | `PJ12H` (na lista de validação) |
| 3ªSB | 1 célula com `PJ12H` |
| ACO | `12` / `24` (numérico) |

E a **2ªSB se contradiz internamente**: a coluna de saldo (`AL`) conta `"PJ12H"`, mas a linha de BM/dia (31) conta `"PJ12"`. Hoje não há `PJ12H` lançado na 2ªSB, então o erro está latente — no primeiro lançamento ele aparece.

### 7.7 MÉDIO — Faixas de linha hardcoded nos totais
`$3:$29`, `$39:$49`, `$55:$92` (1ªSB); `$3:$30`, `$39:$53` (2ªSB); etc. **Inserir ou remover um militar não ajusta nada.** O total continua somando o intervalo antigo — o erro é silencioso.

### 7.8 MÉDIO — Metas hardcoded
`8`, `7`, `4`, `7` repetidos em ~120 fórmulas + rótulos ("Meta = 8"). Mudar a meta do mês = editar 120 células.

### 7.9 MÉDIO — Validações param antes do fim útil
- `RESUMO!E2:E112` e `H2:H112` — há 110 militares ⇒ **só 1 linha de folga**. O 112º militar fica sem validação (escala/seção digitada livre).
- `ACO!D6:D54` usa `RESUMO!$F$2:$F$176`; `CONF. INDIV.!A4` usa `$F$2:$F$136`; `1ªSB!D55:D92` usa `$F$2:$F$176`. Três alcances diferentes para a mesma lista.

### 7.10 BAIXO — Duplicidade de militar
`SD 722022-7 MONTEIRO` aparece **duas vezes na 1ªSB**: linha 21 (bloco da própria SB) e linha 82 (bloco "militares de outras SBs"). Precisa de confirmação se é a mesma pessoa.

### 7.11 BAIXO — Coluna `Nº` quebrada no bloco de visitantes
A fórmula de `A55:A92` procura o nome em `'Base de Dados'!$I$1:$V$40` — que só contém pessoal próprio. Resultado: **38 células em `#N/A`** no bloco "outras SBs".

### 7.12 BAIXO — `#N/A` por linhas preparadas
`1ªSB!B29/C29`: fórmulas ativas na linha sem militar. Total de `#N/A`: 42.

### 7.13 INFORMATIVO — Abas órfãs
- `FILTRO_DIAS`: 4.474 células congeladas de uma `QUERY`, **nenhuma aba a referencia**. Consome 360 KB do arquivo.
- `Cópia de RESUMO`: backup antigo, 105 `#REF!`, nenhuma referência.
- `ORIENTAÇÕES`: oculta — é o manual, mas invisível para quem opera.

### 7.14 INFORMATIVO — `Página30` e `ESCALAS 2026` desconectadas
`Página30` ("Fim da escala híbrida no 7ºGB") aponta **142 serviços/mês necessários contra 107 de efetivo ⇒ déficit de 35**. É um estudo de capacidade valioso, mas isolado: nenhuma fórmula o consome. Idem `ESCALAS 2026` — a tabela de rodízio não é usada por nenhuma validação.

### 7.15 INFORMATIVO — Nada protege nada
Nenhuma aba protegida, nenhum intervalo nomeado, nenhuma tabela. Qualquer usuário pode sobrescrever qualquer fórmula.

---

## 8. Riscos estruturais (o "porquê" dos achados acima)

1. **Acoplamento por texto livre** — a planilha inteira depende de uma string digitada em 5 lugares.
2. **Ausência de tabela/intervalo nomeado** — todos os intervalos são literais (`A1:B2`), logo envelhecem mal.
3. **Blocos fixos por linha** — a estrutura "3–29 / 39–49 / 55–92" não sobrevive a mudança de efetivo.
4. **Fórmulas duplicadas por célula** (não por coluna) — cada linha tem sua receita, o que explica o drift: 5.770 fórmulas, quase todas variações da mesma ideia.
5. **Duas linguagens de código** (PJ12 / PJ12H) e duas unidades de medida (peso 1/2 e horas 12/24).
6. **Snapshot exportado** — a planilha depende de funções Google-only; não há caminho de ida para Excel sem reescrita.

---

## 9. O que ainda preciso confirmar (antes de propor qualquer coisa)

1. A aba `VOLUNTÁRIOS PJES` foi **excluída de propósito** ou é perda? (Isso muda toda a estratégia de PJES.)
2. `MONTEIRO` nas linhas 21 e 82 da 1ªSB é a mesma pessoa?
3. O que significam **`M` / `B` / `F`** na coluna `MOT.`?
4. `Página30` ainda é uma decisão em aberto (fim da escala híbrida) ou já foi resolvida?
5. Existe **Apps Script** em uso (menus, envio de e-mail, geração automática de mês)? Não é visível no `.xlsx`.
6. A planilha é **virgem de novo mês** (reiniciada todo mês) ou o mês novo é criado copiando a aba?
7. Quem são os operadores (quantas pessoas preenchem, com que frequência)?
8. `Z22:Z26` (Ef. Praças) — qual a fonte daquelas quantidades "esperadas"?

---

## 10. Situação

- **Etapa 1 (esta):** concluída.
- **Etapa 2 (diagnóstico detalhado + propostas):** não iniciada — aguardando sua autorização.
- **Nenhuma célula foi alterada.**
