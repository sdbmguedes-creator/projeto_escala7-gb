# DIAGNÓSTICO E PROPOSTAS — Etapa 2

> **Nada foi alterado na planilha.** Este documento é só diagnóstico + proposta.
> Base: `BASE_Escala_de_SERVIÇO.xlsx` (outubro/2026) + respostas e observações do responsável.

---

## A. O que as suas respostas mudaram no diagnóstico

| # | Sua resposta | Consequência para a análise |
|---|---|---|
| 1 | `VOLUNTÁRIOS PJES` excluída **de propósito** | Os 356 `#REF!` são **resíduo**, não perda de dado. Não é caso de recuperação — é caso de limpeza cirúrgica. |
| 2 | `MONTEIRO` duplicado é a **mesma pessoa** | Confirmado defeito de cadastro (linhas 21 e 82 da 1ªSB). Não é caso legítimo de visitante. |
| 3 | `M` = motorista operacional, `B` = motorista administrativo, `F` = efetivo feminino | A conta de motoristas do dia (`C="M"`) **já está correta** — exclui o `B`. O modelo distingue bem; só não está documentado. |
| 4 | `Página30` **resolvida** | Deixa de ser estudo ativo → vira arquivo histórico. |
| 5 | **Sem Apps Script** | Todas as propostas abaixo são **100% nativas** (fórmula, formatação, proteção). Nada de script. |
| 6 | Mês novo = **copiar a planilha** | **Esta é a causa-raiz do drift de faixas.** Cada cópia herda e multiplica as divergências. Vira prioridade máxima. |
| 7 | Operam: **você + 4 comandantes de SB** | 5 editores → exige proteção por área e conferência automática (não dá para revisar tudo a olho). |
| 8 | `Z22:Z26` são **valores manuais** | Não é erro — é parâmetro manual sem rótulo. Só precisa ficar explícito. |

**Observações de contexto que passam a valer como regra do modelo:**
- `RESUMO` é a base única; lançamento manual, **por ordem de antiguidade**.
- `RESUMO` define: motorista (M/B), seção, escala (ALFA/BRAVO/CHARLIE/DELTA) e situação de afastamento (`X - ***`).
- O militar deve aparecer na aba da sua seção, **na ordem do RESUMO e agrupado pela sua escala**.
- `X - ***` = afastado no mês; **permanece** na aba da seção.
- No `ACO`, a tabela de oficiais no topo (linhas 2–3) **não deve ser alterada** — é requisito explícito.

---

## B. DIAGNÓSTICO

### B.1 O que está funcionando corretamente

1. **Arquitetura de fonte única.** O `RESUMO` centraliza cadastro e o resto deriva dele. Conceitualmente correto.
2. **Separação por seção.** Cada SB opera na própria aba; ninguém precisa entender a planilha toda.
3. **Ponderação de carga correta.** 24h = 1 BM, 12h = 0,5 — aplicada consistentemente na 1ªSB.
4. **Distinção M / B preservada** na conta de motoristas (`COUNTIFS(C,"M")` não conta o administrativo).
5. **Conferência em três níveis** (militar → SB → 7ºGB) com semáforo OK/FALTA/EXCESSO. É o melhor ativo do projeto.
6. **Razão de PJES** (`Planilha de PJES`) com reconciliação planilha × escala via `SUMIF`.
7. **Consulta individual** (`CONF. INDIV.`) resolvendo a pergunta mais frequente sem risco de edição.
8. **Datas em cascata.** `ACO!A3` → linha 2 das SBs → `DIÁRIO`. Uma data manda em tudo (o que já é meio caminho da rotina mensal).
9. **Listas de validação existentes** impedem digitação livre de nomes e códigos (mesmo com defeitos de alcance).
10. **Tabela de rodízio 2026** está correta e confere com a prática (validei ALFA nos dias 3, 7, 11 de outubro).

### B.2 Problemas identificados — agrupados por causa raiz

**Causa 1 — Cópia mensal sem âncora (a raiz de quase tudo)**
- Cada célula guardou o intervalo do mês em que foi criada. Resultado: **1ªSB com 23 faixas diferentes, 2ªSB com 27, ACO com 33** (H189…H229).
- Ao crescer o efetivo, parte das fórmulas enxerga a linha nova e parte não → `#N/A` e escala errada, silenciosamente.

**Causa 2 — Estrutura por posição, não por bloco**
- Totais somam faixas literais (`$3:$29`, `$39:$49`, `$55:$92`). Inserir um militar desalinha tudo sem aviso.
- Metas (`8`, `7`, `4`) repetidas em ~120 células. Mudar meta = caçar célula.

**Causa 3 — Vocabulário duplicado**
- PJES 12h: `PJ12` (1ªSB) × `PJ12H` (2ª/3ª/4ª) × `12` (ACO).
- A 2ªSB se contradiz: saldo conta `PJ12H`, BM/dia conta `PJ12`.
- Hoje o erro está mascarado porque não há `PJ12H` lançado — no primeiro lançamento, aparece.

**Causa 4 — Resíduos de decisões antigas**
- 356 `#REF!` (aba excluída + células órfãs), 105 na `Cópia de RESUMO`.
- 42 `#N/A`, sendo 38 na fórmula de `Nº` do bloco "outras SBs" (procura visitante na base de pessoal próprio).
- Duas abas órfãs (`FILTRO_DIAS` com 4.443 fórmulas congeladas / 360 KB; `Cópia de RESUMO`).

**Causa 5 — Nenhuma barreira de proteção**
- 0 abas protegidas, 0 intervalos nomeados, 0 tabelas. 5 pessoas editando o mesmo arquivo sem cerca.
- 253 regras de formatação condicional, várias sobrepostas — difícil saber se uma cor significa algo ou é resto.

### B.3 Pontos que podem ser simplificados

| Item | Situação hoje | Ganho |
|---|---|---|
| Colunas `E` das SBs e `P` do RESUMO | 77 fórmulas mortas (`#REF!`) | Remover ruído |
| `Cópia de RESUMO` | backup obsoleto, 1.111 fórmulas | -1 aba, -105 erros |
| `FILTRO_DIAS` | 4.443 fórmulas congeladas, ninguém usa | -360 KB |
| `Página30` | estudo encerrado | Arquivar |
| `Z22:Z26` | números manuais sem rótulo | Rotular como "parâmetro manual" |
| CF duplicadas (1ªSB tem 41) | várias regras cobrindo a mesma faixa | ~40% menos regras |
| `AN` das SBs | repete o total de PJES que já está no RESUMO | Referência única |

### B.4 Oportunidades de automação (100% nativas)

1. **Aderência à escala** — usar a tabela `ESCALAS 2026` (hoje decorativa) para sinalizar marcação em dia fora da escala do militar.
2. **Dupla escala / conflito** — sinalizar militar marcado em duas SBs no mesmo dia.
3. **Descanso** — sinalizar 24h em dias consecutivos.
4. **Duplicidade de cadastro** — `COUNTIF` no RESUMO destaca nome repetido (pegaria o caso MONTEIRO sozinho).
5. **Competência do mês** — avisar se a linha de datas não bate com o mês de `ACO!A3`.
6. **Rotação anual** — gerar `ESCALAS 2026` por fórmula, sem remontar tabela todo ano.
7. **Cotas** — conferência já existe; pode apenas ficar mais legível.

### B.5 Melhorias na organização visual
- Legenda única de cores (hoje cada aba tem esquema próprio).
- Consolidar CF sobrepostas.
- Área de impressão por SB (para afixar a escala).
- Desocultar `ORIENTAÇÕES` — o manual existe e está invisível.

### B.6 Melhorias na segurança e integridade
- Proteção por área (fórmulas × marcação × cadastro) para 5 usuários.
- Unicidade de nome e formato de matrícula.
- Folga nas listas de validação (hoje o RESUMO valida só até a linha 112 — 1 linha de folga para 110 militares).

### B.7 Melhorias na geração e manutenção das escalas
- Eliminar o "copiar e rezar": a cópia mensal deve nascer estruturalmente correta.
- Uma célula de competência; conferência de que tudo derivou dela.
- Checklist de abertura do mês (hoje é conhecimento tácito seu).

---

## C. PROPOSTAS

Formato: **Problema → Proposta → Como funciona na prática → Abas/colunas afetadas → Benefício → Risco.**

---

### BLOCO A — HIGIENE (baixo risco, ganho imediato)

#### P1. Limpar o resíduo da aba `VOLUNTÁRIOS PJES`

- **Problema:** 680 referências a uma aba extinta; **356 células em `#REF!`** nas colunas que só existem para buscá-la.
- **Proposta:** esvaziar essas células (mantendo a coluna e o cabeçalho, para não mudar o layout). Duas opções:
  - **A (mínima):** apagar o conteúdo das células com `#REF!`.
  - **B (reaproveitar):** transformar a coluna `E` das SBs em "situação PJES" puxando da lista de não-voluntários que hoje vive em `PJES DISTRIBUIÇÃO!C33:M37`.
- **Na prática:** selecionar `RESUMO!P2:P111`, `1ªSB!E3:E92`, `2ªSB!E3:E53`, `3ªSB!E3:E41`, `4ªSB!E3:E55`, `ACO!E6:E54` → apagar conteúdo. Nenhuma fórmula de cálculo depende delas.
- **Afeta:** `RESUMO!P`, `E` das 4 SBs, `ACO!E`.
- **Benefício:** zera 100% dos `#REF!`; a planilha volta a ser legível e as conferências, confiáveis.
- **Risco:** **baixo**. Único cuidado: confirmar que nenhum relatório seu usa essas colunas. Reversível pelo histórico do Google.

#### P2. Resolver a duplicidade do MONTEIRO

- **Problema:** mesma pessoa nas linhas 21 (bloco da SB) e 82 (bloco "outras SBs") da 1ªSB — duplica contagem de PJES.
- **Proposta:** excluir a linha **82** (bloco de visitantes), preservando o cadastro na própria SB. Antes, conferir se há marcação de PJES só na 82 e transferir para a 21.
- **Na prática:** comparar `G82:AK82` com `G21:AK21`; se houver PJ/PJ12 exclusivo da 82, mover; então apagar a linha.
- **Afeta:** `1ªSB!A82:AN82` (e o intervalo dos totais `$55:$92`).
- **Benefício:** corrige contagem; elimina ambiguidade de `VLOOKUP`/`XLOOKUP` (o primeiro resultado ganha).
- **Risco:** **baixo**, com o cuidado de não perder PJES lançado.

#### P3. Arquivar abas órfãs e o estudo encerrado

- **Problema:** `FILTRO_DIAS` (4.443 fórmulas congeladas, 360 KB, zero referências), `Cópia de RESUMO` (1.111 fórmulas, 105 `#REF!`), `Página30` (estudo resolvido).
- **Proposta:** exportar as três para um arquivo `HISTORICO_2025.xlsx` (ou PDF) e excluí-las da planilha. Se preferir não excluir: manter ocultas e renomear com prefixo `_HIST_`.
- **Na prática:** Arquivo → Fazer download → XLSX, escolhendo só essas abas; depois excluir da origem.
- **Afeta:** estrutura de abas (17 → 14).
- **Benefício:** arquivo ~30% menor; some a chance de alguém editar o backup por engano.
- **Risco:** **baixo** — é histórico, não operação. Confirme apenas que `FILTRO_DIAS` não alimenta nenhum painel externo (o Looker Studio, por exemplo).

#### P4. Desocultar e atualizar `ORIENTAÇÕES`

- **Problema:** o manual existe, está **oculto** e desatualizado — não cita `N`, fala só em `PJ12` (enquanto 3 abas usam `PJ12H`) e não documenta `X - ***` nem `M`/`B`/`F`.
- **Proposta:** tornar visível, renomear para `COMO USAR` e reescrever em 6 blocos: códigos, escalas, afastamento, motoristas, PJES, rotina do mês.
- **Na prática:** só texto na coluna C/D — nenhuma fórmula.
- **Afeta:** `ORIENTAÇÕES` (visibilidade + texto).
- **Benefício:** ataca a causa de metade dos erros de digitação; reduz perguntas a você.
- **Risco:** **nenhum**.

---

### BLOCO B — ESTRUTURA QUE NÃO QUEBRA NO MÊS SEGUINTE

#### P5. Faixa única de cadastro (fim do drift)

- **Problema:** 83 variantes de intervalo apontando para o RESUMO (1ªSB 23, 2ªSB 27, ACO 33). Toda cópia mensal gera novas.
- **Proposta:** padronizar **todas** as referências para uma faixa única e folgada: `RESUMO!$F$2:$H$400` (ou intervalo nomeado `CADASTRO`). Substituição mecânica das 83 cadeias de texto.
- **Como funciona na prática:** localizar/substituir por cadeia (ex.: `RESUMO!F$2:H189` → `RESUMO!$F$2:$H$400`), uma vez. A partir daí, o mês novo já nasce certo e o efetivo pode crescer até ~290 pessoas sem tocar em fórmula.
- **Afeta:** ~700 células — `RESUMO` (40), `ACO` (98), `1ªSB` (250), `2ªSB` (84), `3ªSB` (62), `4ªSB` (88), `CONF. INDIV.` (2), `Base de Dados` (49).
- **Benefício:** **elimina a maior fonte de erro futuro** e o principal sintoma da cópia mensal.
- **Risco:** **médio**. Precisa ser feito em cópia, com comparação célula a célula antes/depois (eu faço isso por script) e reversível pelo histórico.

#### P6. Intervalos nomeados nos blocos (inserir militar deixa de ser perigoso)

- **Problema:** totais somam `$3:$29`, `$39:$49`, `$55:$92`. Inserir linha desalinha sem aviso.
- **Proposta:** nomear cada bloco (`S1_PESSOAL`, `S1_ADM`, `S1_VISITANTES`, idem 2ª/3ª/4ª) e usar os nomes nas fórmulas de total. Intervalos nomeados **se expandem** quando se insere linha dentro deles.
- **Como funciona na prática:** criar os nomes (Dados → Intervalos nomeados) e trocar `G$3:G$29` por `G3:G29`… não: trocar por referência completa ao nome. Ex.: `COUNTIF(S1_PESSOAL, "X")` não serve (precisa da coluna do dia). Alternativa correta: `INDEX(S1_PESSOAL,0,COLUMN()-6)` ou simplesmente manter a faixa e aceitar a reserva.
  - **Versão recomendada e mais simples:** fixar a regra "sempre preencher linha vazia, nunca inserir linha", com cada bloco terminando em linha reserva, e os totais cobrindo a reserva (`$3:$40` em vez de `$3:$29`).
- **Afeta:** linhas 30–36 das 4 SBs.
- **Benefício:** acaba a classe de erro mais silenciosa do modelo.
- **Risco:** **médio** — mexer nos limites de bloco toca merges e o bloco ADM. Fazer em cópia.

#### P7. Rotina única de virada do mês

- **Problema:** "copio a planilha e atualizo o mês" é o que propaga drift, resíduo e formatação velha. Não há checklist e a competência está espalhada.
- **Proposta:** (a) manter **uma** célula de competência (`ACO!A3`) e derivar tudo dela; (b) substituir "copiar planilha" por **copiar abas** + 7 passos documentados; (c) conferência automática de competência: avisa se a data da linha 2 não pertence ao mês de `ACO!A3`.
- **Como funciona na prática:** checklist na aba `COMO USAR`: 1) duplicar abas; 2) lançar data em `ACO!A3`; 3) conferir `ESCALAS 2026`; 4) limpar marcações (selecionar intervalo → Delete); 5) reordenar efetivo por antiguidade; 6) rodar conferências; 7) revisar cotas PJES.
- **Afeta:** `ACO!A3`, linha 2 das SBs, `DIÁRIO`, e nova célula de conferência.
- **Benefício:** mês novo nasce correto; menos tempo seu; menos erro humano.
- **Risco:** **baixo** (processo). O ponto de atenção é treinar os 4 comandantes no novo checklist.

---

### BLOCO C — PARÂMETROS E CÓDIGOS

#### P8. Centralizar parâmetros (metas, limites, pesos, cotas)

- **Problema:** meta `8`/`7`/`4` em ~120 células; pesos 1/0,5 repetidos; limites 10/4 só como rótulo; cota 328 solta.
- **Proposta:** uma aba `PARÂMETROS` (visível e curta) com 1 linha por SB: meta BM/dia, meta motoristas, limite masc., limite fem., e a cota PJES do mês. Fórmulas passam a referenciar (via intervalo nomeado `META_1SB` etc.).
- **Como funciona na prática:** `=SUM(8-G30)` vira `=SUM(META_1SB-G30)`. Mudar a meta do mês = editar 1 célula.
- **Afeta:** linhas 30/31/36 das 4 SBs; `DIÁRIO` (médias/meta); `PJES DISTRIBUIÇÃO!G9:G12`.
- **Benefício:** consistência e manutenção trivial; os limites de gênero deixam de ser decoração e passam a valer.
- **Risco:** **baixo-médio** — muitas substituições, mas mecânicas e testáveis.

#### P9. Unificar o código de PJES 12h

- **Problema:** `PJ12` (1ªSB) × `PJ12H` (2ª/3ª/4ª) × `12` (ACO); e a 2ªSB usa **os dois** internamente.
- **Proposta:** adotar **um** código (`PJ12`) em todas as abas, tanto nas listas de validação quanto nas fórmulas; o `12`/`24` do ACO permanece (é hora de turno, não PJES).
- **Como funciona na prática:** atualizar as 4 listas de validação; localizar/substituir `PJ12H` → `PJ12` nas fórmulas e nos lançamentos do mês.
- **Afeta:** validações das 4 SBs; colunas de saldo (`AL`/`AM`/`AK`) e linhas de BM/dia.
- **Benefício:** acaba a subcontagem silenciosa; uma única regra para todos.
- **Risco:** **médio**. Antes, confirme comigo se `PJ12H` tem significado próprio (ex.: "12h em horário especial"). Se tiver, a unificação é outra.

---

### BLOCO D — CONFERÊNCIAS AUTOMÁTICAS

#### P10. Aderência à escala + `ESCALAS 2026` gerada por fórmula

- **Problema:** a tabela anual é decorativa — nada valida que o militar está marcado no dia da sua escala. E ela precisa ser remontada à mão todo ano.
- **Proposta:** (a) gerar a tabela por fórmula a partir de uma data-âncora do ciclo; (b) linha auxiliar (oculta) em cada SB calculando a escala do dia; (c) formatação condicional sinalizando marcação fora da escala.
- **Como funciona na prática:** célula auxiliar `G2` = `INDEX('ESCALAS 2026'!$B$6:$AF$17; MATCH(UPPER(TEXT(G$2,"mmmm")); 'ESCALAS 2026'!$A$6:$A$17;0); DAY(G$2))`; CF: `=E(GB3<>""; GB3<>"N"; $GB$2<>$B3)` → destaca. Para a tabela: `=CHOOSE(MOD(DATA(2026;1;1)-DATA(2026;1;1);4)+1;"BRAVO";"CHARLIE";"DELTA";"ALFA")` replicada.
- **Afeta:** `ESCALAS 2026` (372 células), +1 linha auxiliar por SB, +1 regra de CF por SB.
- **Benefício:** pega o erro mais comum (marcar no dia errado) **antes** de publicar; acaba a remontagem anual.
- **Risco:** **médio**. Só deve sinalizar `X` no início (PJ/PJ12 podem cair fora da escala por natureza). Começar conservador e calibrar.

#### P11. Conferência de descanso e de dupla escala

- **Problema:** nenhuma checagem de serviço em dias consecutivos (24h) nem de militar marcado em duas SBs no mesmo dia.
- **Proposta:** (a) CF sinalizando `X`/`PJ` em dias seguidos; (b) no bloco "outras SBs", avisar se o nome também consta na própria SB.
- **Como funciona na prática:** CF `=E(OU(GB3="X";GB3="PJ"); OU(GC3="X";GC3="PJ"))` → amarelo; conferência cruzada por `COUNTIF` do nome nas duas colunas D.
- **Afeta:** grade de marcação das 4 SBs + bloco de visitantes.
- **Benefício:** evita conflito operacional e questionamento jurídico de descanso.
- **Risco:** **baixo** (só sinaliza, não bloqueia).

#### P12. Integridade do cadastro (unicidade e formato)

- **Problema:** chave textual; 3 nomes com espaço sobrando; duplicidade passou despercebida.
- **Proposta:** CF de duplicidade no RESUMO (`=CONT.SE($F$2:$F$400;$F2)>1`), validação de formato de matrícula (`6-1` dígitos), e normalização de espaços.
- **Como funciona na prática:** a duplicidade acende sozinha — inclusive o caso MONTEIRO.
- **Afeta:** `RESUMO!C`, `RESUMO!F`, e os 3 nomes já identificados nas SBs.
- **Benefício:** erro pego na origem, não na conferência do mês.
- **Risco:** **baixo, com uma armadilha importante** — aplicar `TRIM` só no RESUMO **quebra** o `VLOOKUP` das SBs (que têm o nome com espaço). Tem de ser feito nos dois lados no mesmo passo. Recomendo corrigir os 3 nomes à mão primeiro e só depois ligar o `TRIM`.

---

### BLOCO E — OPERAÇÃO E VISUAL

#### P13. Proteção por área (5 usuários)

- **Problema:** 0 proteções. Um comandante pode apagar fórmula de total sem perceber.
- **Proposta:** proteger os blocos de fórmula (linhas 30–36 e colunas de saldo) com **"mostrar aviso ao editar"** num primeiro ciclo; depois, se estabilizar, bloquear de fato e liberar edição só às células de marcação e cadastro.
- **Como funciona na prática:** Dados → Proteger intervalos; por usuário, se as contas forem nominais.
- **Afeta:** as 4 SBs + `RESUMO` + `DIÁRIO`.
- **Benefício:** impede a quebra acidental mais grave.
- **Risco:** **médio** — proteção mal configurada trava o comandante no meio do lançamento. Implantar em 2 etapas (aviso → bloqueio) e com teste na cópia.

#### P14. Visual: legenda única, CF enxuta, área de impressão

- **Problema:** 253 regras de CF, esquemas de cor diferentes por aba, regras sobrepostas (1ªSB tem 41), sem área de impressão.
- **Proposta:** legenda única (X=verde, PJ=azul, PJ12=ciano, N=cinza, afastado=cinza escuro), consolidar CF sobrepostas, definir área de impressão `A1:AK92` por SB.
- **Como funciona na prática:** criar uma célula-legenda no topo (ou na aba `COMO USAR`) e padronizar as regras.
- **Afeta:** CF das 8 abas de escala.
- **Benefício:** leitura em 2 segundos; impressão correta para afixar.
- **Risco:** **baixo** — mas exige revisar 253 regras (trabalho chato, não complexo).

---

## D. MATRIZ DE PRIORIDADE

| Fase | Propostas | Esforço | Risco | Ganho |
|---|---|---|---|---|
| **1 — Higiene** (fazer já) | P1, P2, P3, P4 | 1–2 h | Baixo | Some o ruído: 356 `#REF!` + 42 `#N/A` + 3 abas mortas |
| **2 — Estabilização** | P5, P6, P7, P8 | 4–6 h | Médio | A cópia mensal para de degradar |
| **3 — Controle** | P9, P10, P11, P12 | 3–5 h | Médio | Erro de marcação pego antes de publicar |
| **4 — Operação** | P13, P14 | 2–3 h | Médio/baixo | 5 usuários sem risco de quebra |

**Sequência recomendada:** P1 → P4 → P2 → P5 → P8 → P9 → P10 → P12 → P6 → P7 → P11 → P13 → P14 → P3.
(O P3 por último: é o único irreversível na prática e não bloqueia nada.)

---

## E. O que eu **não** recomendo fazer

1. **Migrar para Excel/VBA** — você já decidiu; reforço: 1.400 fórmulas usam `INDIRECT`, `XLOOKUP` e `QUERY`. A migração seria reconstrução.
2. **Reconstruir do zero** — o modelo é bom; o problema é manutenção, não desenho.
3. **Introduzir Apps Script** — não há nenhum hoje, e todas as necessidades acima se resolvem com fórmula nativa. Script adiciona dependência e alguém terá de mantê-lo.
4. **Trocar a chave textual por ID numérico agora** — é a correção definitiva, mas toca ~700 fórmulas e 5 abas. Melhor conviver com `TRIM` + conferência de duplicidade. Fica como decisão de médio prazo.
5. **Mexer na tabela de oficiais do ACO (linhas 2–3)** — requisito explícito seu; nenhuma proposta acima toca nela.

---

## F. Como vamos testar (para a Etapa 3/4)

1. Trabalhar sempre em **cópia** da planilha, nunca na original.
2. Antes/depois de cada bloco: **comparação célula a célula** por script (valores calculados), para provar que nada fora do escopo mudou.
3. Cenários de teste: mês de 30 e de 31 dias; militar afastado (`X - ***`); visitante de outra SB; PJES 12h; militar novo inserido.
4. Só depois de o diff fechar limpo é que a alteração vai para a planilha real.

---

## G. REGISTRO DE MELHORIAS

| ID | Melhoria | Status | Fase |
|---|---|---|---|
| P1 | Limpar resíduo `VOLUNTÁRIOS PJES` | **Proposta — aguardando aprovação** | 1 |
| P2 | Duplicidade MONTEIRO | **Proposta** | 1 |
| P3 | Arquivar abas órfãs / Página30 | **Proposta** | 4 |
| P4 | `ORIENTAÇÕES` visível e atualizada | **Proposta** | 1 |
| P5 | Faixa única de cadastro | **Proposta** | 2 |
| P6 | Blocos com reserva | **Proposta** | 2 |
| P7 | Rotina de virada do mês | **Proposta** | 2 |
| P8 | Centralizar parâmetros | **Proposta** | 2 |
| P9 | Unificar `PJ12`/`PJ12H` | **Proposta** | 3 |
| P10 | Aderência à escala + ESCALAS por fórmula | **Proposta** | 3 |
| P11 | Descanso e dupla escala | **Proposta** | 3 |
| P12 | Unicidade/formato de cadastro | **Proposta** | 3 |
| P13 | Proteção por área | **Proposta** | 4 |
| P14 | Visual, CF e impressão | **Proposta** | 4 |
| — | Implementadas | **nenhuma** | — |
| — | Rejeitadas | **nenhuma** | — |
