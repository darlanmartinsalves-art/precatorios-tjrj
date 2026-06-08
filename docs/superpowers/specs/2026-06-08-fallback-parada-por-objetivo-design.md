# Design — Parada por objetivo no fallback de download

**Data:** 2026-06-08
**Status:** Aprovado (aguardando revisão do spec)

## Problema

No pipeline `baixar_requisitorios.py`, a busca das peças de um processo tem dois passos
(`_baixar_e_extrair_pecas_ofreq`):

1. **Passo 1 — por rótulo:** procura na árvore documentos cujo *rótulo* casa
   `PADROES_PECA_REQUISITORIO` ("OFÍCIO REQUISITÓRIO", "PREC DEFINITIVO",
   "PRECATÓRIO GERADO", "DEPJU"). Quando o cartório nomeia direito, acha os alvos e baixa
   só eles. Rápido.

2. **Passo 2 — fallback por conteúdo:** dispara quando o passo 1 não rende requisitório
   **ou** falta vínculo. Baixa até `LIMITE_FALLBACK_GENERICO = 30` documentos genéricos
   ("Petição"/"Documento"), lendo o conteúdo de cada PDF pra classificar.

**Sintoma observado:** em processos onde os requisitórios/vínculos são arquivados como
"Petição" genérica (comum em varas que não nomeiam as peças), o passo 2 baixa **30
documentos** por processo. O early-stop atual (`FALLBACK_EARLY_STOP_MISSES = 8`) só conta
misses **depois** de achar o primeiro documento relevante — então, quando o alvo demora a
aparecer (ou os primeiros candidatos falham), o teto inteiro de 30 é consumido.

**Causa raiz:** o fallback não usa o objetivo real — encontrar os N precatórios-alvo
daquele processo, que são conhecidos a partir da planilha (filtrados por saldo ≥ R$ 200k).
Ele só tem heurísticas de ordenação e de misses, não um sinal de "já achei o que precisava".

## Confirmação dos fatos (prints do usuário, processo 0110128-60.2015)

- Os alvos são **texto digital extraível** (não escaneados): Ofício Requisitório Definitivo
  (com "III - BENEFICIÁRIO") e Ofício DEPJU/vínculo ("Ofício {OFREQ}/OFREQ ... gerou o
  precatório {precatório}").
- Ambos são rotulados como **"Petição" genérica** na árvore (ex.: `684 - 202501640984 -
  Petição`), por isso o passo 1 não os acha.
- Ficam **geralmente no fim** do processo (eventos mais recentes).
- Os classificadores (`eh_documento_requisitorio`, `eh_documento_vinculo`, extração de
  OFREQ, beneficiário, advogado) **já reconhecem** esses formatos. O gap é 100% a
  **seleção/parada**, não a classificação.

## Solução (Abordagem A — parada por objetivo)

A cada documento baixado no fallback, checar se todos os precatórios-alvo do processo já
estão resolvidos e, em caso afirmativo, parar imediatamente.

### Componentes

**1. Função pura `_todos_resolvidos(precatorios_alvo, requisitorios, vinculos)`**

```python
def _todos_resolvidos(precatorios_alvo, requisitorios, vinculos):
    """True quando todo precatório-alvo tem vínculo (OFREQ->precatório) E o
    requisitório daquele OFREQ extraído. precatorios_alvo vazio/sem match -> False."""
    if not precatorios_alvo:
        return False
    ofreqs_req = {r["ofreq"] for r in requisitorios if r.get("ofreq")}
    resolvidos = {vinculos[of] for of in vinculos if of in ofreqs_req}
    return set(precatorios_alvo) <= resolvidos
```

Um precatório-alvo está "resolvido" quando existe um vínculo OFREQ→precatório **e** o
requisitório com aquele OFREQ foi extraído com sucesso (entrou na lista `requisitorios`).
Requisitórios com extração de beneficiário falha NÃO entram na lista (vão pra
`manual_revisar/`), logo não marcam o alvo como resolvido — comportamento correto.

**2. Threadar a lista de alvos no loop**

`_baixar_e_extrair_pecas_ofreq` passa a receber `precatorios_do_processo` (o caller já tem
essa lista; usa em `casar_requisitorios_com_vinculos`) e repassa a
`_baixar_e_classificar_indices` no passo 2 (fallback).

**3. Goal-stop dentro do loop** (`_baixar_e_classificar_indices`)

Depois de cada download/classificação, se `_todos_resolvidos(precatorios_alvo, ...)` →
`break` imediato. O early-stop por misses (`parar_apos_misses`) permanece como **backstop**
pro caso de documento faltante/escaneado.

**4. Reduzir `LIMITE_FALLBACK_GENERICO` 30 → 15**

Teto rígido do pior caso (alvo realmente ausente). Com o goal-stop, no caso comum o loop
para muito antes. Constante ajustável em `seletores.py`.

### Comportamento resultante

| Cenário | Antes | Depois |
|---|---|---|
| Caso bom (0110128, 4 precatórios digitais no fim) | até 30 downloads | para ao resolver os 4 (provável <10) |
| `--test-processar` (alvo `["TESTE"]`) | n/a | nunca casa → backstop teto/misses (sem regressão) |
| Alvo ausente/escaneado | 30 downloads | no máx. 15, depois desiste |

### O que NÃO muda

- Classificadores e extração (OFREQ, beneficiário, advogado, vínculo) — já funcionam.
- Ligação por OFREQ (`casar_requisitorios_com_vinculos`).
- Passo 1 (por rótulo) e a leitura do-fim-pra-frente nos dois passos.
- Gatilho do passo 2 (`if not requisitorios or _faltam_vinculos(...)`).

A mudança é cirúrgica na **seleção/parada** do passo 2.

## Testes (TDD)

**`_todos_resolvidos` (unitário):**
- lista de alvos vazia → False
- nenhum alvo resolvido → False
- parcial (alguns resolvidos) → False
- todos resolvidos → True
- modo "TESTE" (`["TESTE"]`) → False
- vínculo presente mas requisitório do OFREQ ausente → False

**Loop `_baixar_e_classificar_indices` (integração com fakes):**
- para (break) assim que os alvos ficam resolvidos
- não para antes do bloco quando ainda faltam alvos
- sem `precatorios_alvo` (None) → comportamento atual (só teto/misses), sem regressão

## Pendências/notas operacionais

- A correção muda o resultado de processos antes marcados `sem_requisitorio` por exaustão do
  fallback. Para reaplicar a processos já no checkpoint, usar `--apenas-erros` não basta
  (eles não são `erro_*`); só `--reset` reprocessa tudo. Avaliar no momento da re-execução
  do lote.
