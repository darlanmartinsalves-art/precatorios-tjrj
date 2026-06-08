# Parada por Objetivo no Fallback — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fazer o fallback por conteúdo parar de baixar documentos assim que todos os precatórios-alvo do processo estiverem resolvidos, em vez de moer até o teto de 30.

**Architecture:** Função pura `_todos_resolvidos` decide quando o objetivo foi atingido; o loop `_baixar_e_classificar_indices` ganha um goal-stop que a consulta a cada download. A lista de precatórios-alvo (já conhecida do caller) é threadada até o loop. O teto de fallback cai 30→15 como backstop.

**Tech Stack:** Python 3.12, pytest, Playwright (async). Testes com `monkeypatch` e fakes (sem browser).

---

### Task 1: Função pura `_todos_resolvidos`

**Files:**
- Modify: `baixar_requisitorios.py` (adicionar função perto de `_faltam_vinculos`, ~linha 1040)
- Modify: `tests/test_baixar_requisitorios.py` (import na linha ~614 + novos testes)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever os testes que falham**

Adicionar `_todos_resolvidos` ao bloco de import existente em `tests/test_baixar_requisitorios.py` (perto da linha 614, junto de `_faltam_vinculos`):

```python
from baixar_requisitorios import (
    REGEX_OFREQ, REGEX_VINCULO, _indices_fallback, _faltam_vinculos,
    _todos_resolvidos,
)
```

Adicionar os testes (depois dos testes de `_faltam_vinculos`):

```python
# ===== _todos_resolvidos =====
def test_todos_resolvidos_lista_vazia_false():
    assert _todos_resolvidos([], [{"ofreq": "2024.1"}], {"2024.1": "2025.10-0"}) is False

def test_todos_resolvidos_completo_true():
    reqs = [{"ofreq": "2024.1"}, {"ofreq": "2024.2"}]
    vinc = {"2024.1": "2025.10-0", "2024.2": "2025.20-1"}
    assert _todos_resolvidos(["2025.10-0", "2025.20-1"], reqs, vinc) is True

def test_todos_resolvidos_parcial_false():
    reqs = [{"ofreq": "2024.1"}]
    vinc = {"2024.1": "2025.10-0", "2024.2": "2025.20-1"}
    assert _todos_resolvidos(["2025.10-0", "2025.20-1"], reqs, vinc) is False

def test_todos_resolvidos_modo_teste_false():
    reqs = [{"ofreq": "2024.1"}]
    vinc = {"2024.1": "2025.10-0"}
    assert _todos_resolvidos(["TESTE"], reqs, vinc) is False

def test_todos_resolvidos_vinculo_sem_requisitorio_false():
    assert _todos_resolvidos(["2025.10-0"], [], {"2024.1": "2025.10-0"}) is False

def test_todos_resolvidos_requisitorio_sem_vinculo_false():
    assert _todos_resolvidos(["2025.10-0"], [{"ofreq": "2024.1"}], {}) is False
```

- [ ] **Step 2: Rodar os testes pra confirmar que falham**

Run: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -k todos_resolvidos -v`
Expected: FAIL com `ImportError: cannot import name '_todos_resolvidos'`

- [ ] **Step 3: Implementar a função**

Em `baixar_requisitorios.py`, logo após `_faltam_vinculos` (~linha 1055):

```python
def _todos_resolvidos(precatorios_alvo, requisitorios, vinculos):
    """True quando todo precatório-alvo tem vínculo (OFREQ->precatório) E o
    requisitório daquele OFREQ extraído (entrou em `requisitorios`).

    precatorios_alvo vazio (ou só com valores que nunca casam, ex: ["TESTE"]) -> False,
    pra não disparar goal-stop em modo de teste/sem alvos reais.
    """
    if not precatorios_alvo:
        return False
    ofreqs_req = {r["ofreq"] for r in requisitorios if r.get("ofreq")}
    resolvidos = {vinculos[of] for of in vinculos if of in ofreqs_req}
    return set(precatorios_alvo) <= resolvidos
```

- [ ] **Step 4: Rodar os testes pra confirmar que passam**

Run: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -k todos_resolvidos -v`
Expected: PASS (6 passed)

- [ ] **Step 5: Commit**

```bash
git add baixar_requisitorios.py tests/test_baixar_requisitorios.py
git commit -m "feat: _todos_resolvidos — detecta quando todos precatorios-alvo estao resolvidos"
```

---

### Task 2: Goal-stop no loop `_baixar_e_classificar_indices`

**Files:**
- Modify: `baixar_requisitorios.py:1221-1247` (`_baixar_e_classificar_indices`)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever o teste que falha**

Adicionar (perto dos testes de early-stop, ~linha 777). Os fakes `_FakeNth`/`_FakeVisu` já existem no arquivo:

```python
def test_baixar_indices_goal_stop_para_ao_resolver(monkeypatch):
    """Para assim que todos os precatórios-alvo ficam resolvidos (não lê o resto)."""
    processados = []

    async def fake_um(visualizador, el, idx, pasta_temp, requisitorios, vinculos, *a, **k):
        processados.append(idx)
        if idx == 10:
            vinculos["2024.1"] = "2025.10-0"
            requisitorios.append({"ofreq": "2024.1"})
            return True
        return False
    monkeypatch.setattr(_br_timeout, "_baixar_e_classificar_um", fake_um)
    asyncio.run(_br_timeout._baixar_e_classificar_indices(
        _FakeVisu(), [10, 11, 12], None, [], {}, set(),
        log=lambda m: None, numero_processo=None,
        precatorios_alvo=["2025.10-0"]))
    assert processados == [10]  # parou logo após resolver o único alvo


def test_baixar_indices_goal_stop_segue_se_falta_alvo(monkeypatch):
    """Com 2 alvos e só 1 resolvido, NÃO para — segue lendo."""
    processados = []

    async def fake_um(visualizador, el, idx, pasta_temp, requisitorios, vinculos, *a, **k):
        processados.append(idx)
        if idx == 10:
            vinculos["2024.1"] = "2025.10-0"
            requisitorios.append({"ofreq": "2024.1"})
            return True
        return False
    monkeypatch.setattr(_br_timeout, "_baixar_e_classificar_um", fake_um)
    asyncio.run(_br_timeout._baixar_e_classificar_indices(
        _FakeVisu(), [10, 11, 12], None, [], {}, set(),
        log=lambda m: None, numero_processo=None,
        precatorios_alvo=["2025.10-0", "2025.20-1"]))
    assert processados == [10, 11, 12]  # nunca resolveu o 2º alvo -> leu tudo
```

- [ ] **Step 2: Rodar pra confirmar que falha**

Run: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -k goal_stop -v`
Expected: FAIL com `TypeError: ... unexpected keyword argument 'precatorios_alvo'`

- [ ] **Step 3: Adicionar o parâmetro e o goal-stop**

Em `baixar_requisitorios.py`, alterar a assinatura e o loop de `_baixar_e_classificar_indices`:

```python
async def _baixar_e_classificar_indices(visualizador, indices, pasta_temp,
                                        requisitorios, vinculos, ofreqs_vistos,
                                        log, numero_processo, parar_apos_misses=None,
                                        precatorios_alvo=None):
    """Itera os índices baixando+classificando cada um. Muta os acumuladores.

    goal-stop: se `precatorios_alvo` for dado e TODOS ficarem resolvidos
    (vínculo + requisitório), para imediatamente.

    early-stop: se `parar_apos_misses` for dado, para depois de N downloads
    consecutivos IRRELEVANTES *após já ter achado* o bloco. Misses ANTES do primeiro
    achado não contam. (backstop pro caso de alvo faltante/escaneado)
    """
    todos_nos = visualizador.locator(seletores.ARVORE_ITEM_TREEITEM)
    achou_relevante = False
    misses_seguidos = 0
    for idx in indices:
        el = todos_nos.nth(idx)
        relevante = await _baixar_e_classificar_um(
            visualizador, el, idx, pasta_temp,
            requisitorios, vinculos, ofreqs_vistos, log, numero_processo)
        if precatorios_alvo and _todos_resolvidos(precatorios_alvo, requisitorios, vinculos):
            log("  goal-stop: todos os precatórios-alvo resolvidos — parando")
            break
        if relevante:
            achou_relevante = True
            misses_seguidos = 0
        elif achou_relevante:
            misses_seguidos += 1
            if parar_apos_misses is not None and misses_seguidos >= parar_apos_misses:
                log(f"  early-stop: {misses_seguidos} downloads irrelevantes após o bloco — parando")
                break
```

- [ ] **Step 4: Rodar pra confirmar que passa (e sem regressão no early-stop)**

Run: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -k "goal_stop or early_stop or baixar_indices" -v`
Expected: PASS (todos os de goal_stop + os 3 de early-stop/baixar_indices existentes)

- [ ] **Step 5: Commit**

```bash
git add baixar_requisitorios.py tests/test_baixar_requisitorios.py
git commit -m "feat: goal-stop no loop de fallback — para ao resolver todos os alvos"
```

---

### Task 3: Threadar `precatorios_do_processo` até o fallback

**Files:**
- Modify: `baixar_requisitorios.py:1249` (`_baixar_e_extrair_pecas_ofreq` assinatura + chamada do passo 2)
- Modify: `baixar_requisitorios.py:1406-1407` (call site em `processar_processo`)

> Nota: o passo 1 (por rótulo) permanece SEM goal-stop, conforme o spec ("Passo 1 inalterado"). Só o passo 2 (fallback) recebe `precatorios_alvo`.

- [ ] **Step 1: Adicionar o parâmetro na assinatura**

Em `baixar_requisitorios.py:1249`, alterar a assinatura de `_baixar_e_extrair_pecas_ofreq`:

```python
async def _baixar_e_extrair_pecas_ofreq(visualizador, pasta_temp, debug=False,
                                        numero_processo=None, precatorios_do_processo=None):
```

- [ ] **Step 2: Passar `precatorios_alvo` na chamada do passo 2**

No mesmo arquivo, na chamada de `_baixar_e_classificar_indices` DENTRO do bloco do passo 2 (atual ~linha 1301), adicionar o kwarg `precatorios_alvo`:

```python
        await _baixar_e_classificar_indices(
            visualizador, indices2, pasta_temp,
            requisitorios, vinculos, ofreqs_vistos, log, numero_processo,
            parar_apos_misses=seletores.FALLBACK_EARLY_STOP_MISSES,
            precatorios_alvo=precatorios_do_processo)
```

- [ ] **Step 3: Passar a lista no call site**

Em `baixar_requisitorios.py:1406-1407`, alterar a chamada em `processar_processo`:

```python
        requisitorios, vinculos = await _baixar_e_extrair_pecas_ofreq(
            visu, pasta_tmp, debug=debug, numero_processo=numero_processo,
            precatorios_do_processo=precatorios_do_processo)
```

- [ ] **Step 4: Rodar a suíte inteira (garantir zero regressão)**

Run: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -q`
Expected: PASS (todos os testes, incluindo os novos de Task 1 e 2)

- [ ] **Step 5: Commit**

```bash
git add baixar_requisitorios.py
git commit -m "feat: threadar precatorios_do_processo ate o fallback (goal-stop ativo)"
```

---

### Task 4: Reduzir o teto do fallback 30 → 15

**Files:**
- Modify: `seletores.py:137-141` (`LIMITE_FALLBACK_GENERICO`)

- [ ] **Step 1: Alterar a constante e o comentário**

Em `seletores.py`, substituir o bloco de `LIMITE_FALLBACK_GENERICO`:

```python
# Teto do fallback — cada candidato custa ~5-8s de clique+download. Com o goal-stop
# (para assim que todos os precatórios-alvo do processo são resolvidos), no caso comum
# o loop para muito antes; este teto só limita o PIOR caso (alvo realmente ausente ou
# escaneado). Como o fallback varre do FIM da árvore (onde ficam os requisitórios/DEPJU),
# 15 alcança os alvos com folga e corta o desperdício pela metade vs. o antigo 30.
LIMITE_FALLBACK_GENERICO = 15
```

- [ ] **Step 2: Rodar a suíte (a constante não deve quebrar nada)**

Run: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -q`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add seletores.py
git commit -m "tune: teto do fallback 30->15 (goal-stop torna o teto alto desnecessario)"
```

---

## Verificação final (após todas as tasks)

- [ ] Suíte completa verde: `& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -q`
- [ ] (Opcional, exige login) Verificação ao vivo num processo multi-precatório conhecido (ex.: 0110128-60.2015) com `--test-processar`: confirmar status `ok`, mesmos beneficiários por precatório, e que o log mostra `goal-stop` parando antes dos 15 downloads.

## Notas operacionais

- A correção muda o resultado de processos antes marcados `sem_requisitorio` por exaustão do fallback. Num run normal eles NÃO são reprocessados; para reaplicar, usar `--reset` no lote definitivo.
