# RELATÓRIO — BLOCO A executado (em cópia)

**Data:** 01/11/2026 · **Original:** `BASE_Escala_de_SERVIÇO.xlsx` (intacto) · **Resultado:** `BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx`

---

## 1. O que foi executado

| Item | Ação | Números |
|---|---|---|
| **P1** | Esvaziadas as fórmulas que apontavam para `VOLUNTÁRIOS PJES` (cabeçalhos preservados) | **240 células** |
| **P2** | Removida a duplicidade do `MONTEIRO` (linha 82 da 1ªSB, sem nenhuma marcação) | **38 células** |
| **P3** | Exportadas e removidas 3 abas | `FILTRO_DIAS`, `Cópia de RESUMO`, `Página30` → `HISTORICO_2025.xlsx` |
| **P4** | `ORIENTAÇÕES` desocultada + bloco novo de documentação | 27 células de texto |

**Resultado:** 17 → 14 abas. `#REF!` de **356 → 0** (por construção: as fórmulas que os produziam foram removidas).

Abas finais: ORIENTAÇÕES, RESUMO, DIÁRIO, ACO, 1ªSB, 2ªSB, 3ªSB, 4ªSB, PJES DISTRIBUIÇÃO, CONF. INDIV., Planilha de PJES, Oficiais, ESCALAS 2026, Base de Dados.

---

## 2. Prova de que nada mais mudou

Foram três checagens independentes:

1. **Diff célula a célula** contra o original: 305 células alteradas = 240 (P1) + 38 (P2) + 27 (texto de P4). **Zero** alterações fora desse conjunto.
2. **Análise estática de dependência:** nenhuma fórmula sobrevivente lê as células limpas. Os 273 "toques" detectados são `VLOOKUP` cujo *intervalo* inclui a coluna `E`, mas cujo índice de coluna (3, 4, 5…) nunca seleciona a coluna 2 do intervalo — ou seja, a coluna `E` nunca é lida. Também não há `SUM`/`COUNTIF` sobre ela.
3. **Integridade de recursos:** validações de dados, formatações condicionais e intervalos mesclados **idênticos** ao original em todas as abas (conferido aba por aba).

**O ponto sensível que foi preservado:** `1ªSB!E39:E49` contém **cotas manuais de PJES do ADM** (6, 0, 8, 6, 0…). Não são fórmula e **não foram tocadas** — e `PJES DISTRIBUIÇÃO!P5:P15` continua lendo esses valores.

---

## 3. O que **não** foi feito (e por quê)

| Item | Situação |
|---|---|
| **42 `#N/A`** (40 na 1ªSB, 2 no ACO) | **Permanecem.** São fórmulas de `Nº` no bloco de visitantes (`A55:A92`) que procuram o nome numa base que não os contém, e duas fórmulas órfãs em `B29`/`C29`. Isso é o **P1b** — mesma natureza do Bloco A, 2 minutos, mas não estava na sua autorização. |
| **P9 (PJ12H)** | **Revisado** — ver §5. |
| Demais propostas | Não iniciadas. |

---

## 4. Ressalva técnica importante

Ao salvar via openpyxl, o arquivo **perde os valores em cache**: as fórmulas voltam a "não calculadas". Tentativa de recalcular com LibreOffice falhou (a planilha usa `INDIRECT`/`XLOOKUP` demais).

**Consequência prática:**
- No **Google Sheets**, isso é irrelevante — ele recalcula ao importar.
- Mas **não recomendo importar este `.xlsx` por cima da planilha viva**: a `Base de Dados` tem 408 células com o artefato `__xludf.DUMMYFUNCTION` e há fórmulas `INDIRECT` que não sobrevivem bem a um ciclo exportar→importar.
- **Este arquivo é para conferência.** A aplicação na planilha real deve ser feita nativamente, pelo roteiro do §6.

---

## 5. P9 revisada — o `PJ12H` não é gambiarra a eliminar

Sua explicação muda a proposta: `PJ12H` é **código de controle**, não erro de digitação. Ele separa o PJES de 12h do **militar de outra SB** (visitante) do PJES do efetivo próprio, para não contaminar a contagem de serviço e de cotas da SB que recebe o serviço.

**Portanto: não unificar.** O que resta como problema real:

1. **A regra não está escrita em lugar nenhum** — quem abrir a planilha em 6 meses não vai saber por que existem dois códigos.
2. **A 2ªSB se contradiz internamente:** a coluna de saldo conta `PJ12H`, mas a linha de meta diária conta `PJ12`.
3. `PJ12H` só existe nas abas 2ª/3ª/4ª; a 1ªSB usa `PJ12` até para visitante.

**Nova proposta P9:** documentar a regra (feito parcialmente no P4) + alinhar os contadores de cada aba à mesma definição + estender o `PJ12H` à 1ªSB, se o critério for o mesmo.

**Decisão que preciso de você:** o PJES de 12h do visitante (`PJ12H`) deve **contar como BM na meta diária** da SB que o recebe? O serviço é prestado lá — mas você disse que ele não deve atrapalhar a contagem. Preciso saber qual das duas vale, porque é isso que define a correção da 2ªSB.

---

## 6. Como aplicar no Google Sheets (roteiro seguro)

> Faça na planilha viva, em 5 minutos. Sem importar arquivo.

**Antes de tudo:** Arquivo → Fazer download → Microsoft Excel (.xlsx). Isso é o seu backup do dia.

**P3 — arquivar e excluir 3 abas**
1. Clique com o botão direito em `FILTRO_DIAS` → Excluir. (A confirmação de exclusão é definitiva; o backup acima cobre.)
2. Idem para `Cópia de RESUMO` e `Página30`.

**P1 — limpar as colunas mortas** (selecione o intervalo e pressione `Delete` — só o conteúdo)
| Aba | Intervalo |
|---|---|
| RESUMO | `P2:P112` |
| ACO | `E6:E54` |
| 1ªSB | `E3:E29` **e** `E82` |
| 2ªSB | `E3:E30` |
| 3ªSB | `E3:E18` |
| 4ªSB | `E3:E31` |

> ⚠️ Na 1ªSB **não** apague `E39:E49` — são as cotas manuais do ADM.

**P2 — duplicidade MONTEIRO**
- Na 1ªSB, selecione `A82:AN82` e apague o conteúdo. (A linha 82 não tem nenhuma marcação; o MONTEIRO real está na linha 21, com 8 serviços e 4 PJES.)

**P4 — ORIENTAÇÕES**
- Botão direito na aba `ORIENTAÇÕES` → Mostrar aba.
- Clique na área mesclada `C18` (o texto "Algumas orientações"), apague e cole o texto do §7.

**Conferência final:** a planilha deve ficar sem nenhum `#REF!`. Os `#N/A` do bloco de visitantes continuam (é o P1b).

---

## 7. Texto novo para `ORIENTAÇÕES` (colar em C18)

```
CÓDIGOS DE PREENCHIMENTO

X    Serviço ordinário de 24h, na escala fixa do militar (rodízio ALFA/BRAVO/CHARLIE/DELTA). Vale 1 na meta diária.
PJ   PJES de 24h. Vale 1 na meta diária.
PJ12 PJES de 12h. Vale 0,5 na meta diária.
PJ12H Código de controle das abas 2ª/3ª/4ªSB para o PJES de 12h de militar de OUTRA SB (visitante),
     para não entrar na contagem de serviço e de cotas da SB que o recebe. Não é erro de digitação.
S1   Apenas saldo de 1º turno. Vale 0,5.
SPJ  Saldo de 1º turno + PJES de 24h. Vale 1.
N    Não disponível / sem lançamento no dia. Não conta na meta.
12 e 24 (só no ACO)  Turno de 12h ou 24h do oficial. É hora de turno, não é PJES.

AFASTAMENTO
X - A / X - B / X - C / X - D / X - ADM = militar afastado no mês (a letra é a escala dele).
Ele continua na aba da sua seção, com N nos dias. Não remova a linha.

MOT. (motorista)
M = motorista de viatura operacional (entra na conta de motoristas do dia).
B = motorista administrativo (não entra na conta).

FEM
F = efetivo feminino, usado na conferência de limite de feminino por dia.

ORDEM DE LANÇAMENTO
No RESUMO, lançar as praças por ANTIGUIDADE. Na aba da SB, manter a mesma ordem, agrupado pela
escala (ALFA, BRAVO, CHARLIE, DELTA). Não inserir linha: usar a próxima linha vazia do bloco.

ROTINA DE ABERTURA DO MÊS
1) duplicar as abas   2) lançar o 1º dia do mês em ACO!A3   3) conferir a tabela ESCALAS
4) limpar as marcações antigas   5) reordenar o efetivo por antiguidade
6) revisar as cotas PJES   7) rodar as conferências
```

---

## 8. Arquivos gerados

| Arquivo | Para quê |
|---|---|
| `BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx` | Versão limpa, **para conferência** (não importar) |
| `HISTORICO_2025.xlsx` | As 3 abas removidas, preservadas |
| `BASE_Escala_de_SERVIÇO.xlsx` | Original, **intacto** |
