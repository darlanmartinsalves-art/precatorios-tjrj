# Classificação por Prévia (Abordagem 2 + fallback 1) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Localizar requisitório + vínculo DEPJU no visualizador TJRJ classificando o texto da PRÉVIA (sem baixar), baixando só os requisitórios confirmados; fallback por download quando a prévia não é texto.

**Architecture:** Uma passada do FIM pro início sobre as folhas da árvore. Cada nó: clica → lê texto da prévia (DOM) → classifica com `eh_documento_*` → vínculo extrai OFREQ→precatório do texto (sem baixar); requisitório baixa o PDF e extrai beneficiário/advogado; prévia vazia/imagem cai no fallback por download. goal-stop encerra ao resolver todos os alvos.

**Tech Stack:** Python 3.12, Playwright (async), pypdf, pytest. Arquivos: `baixar_requisitorios.py`, `seletores.py`, `tests/test_baixar_requisitorios.py`.

---

## File Structure

- `baixar_requisitorios.py` — extratores por texto (novos), leitor de prévia (novo), refactor da classificação por nó.
- `seletores.py` — constante `SELETOR_PREVIA_TEXTO` (preenchida pelo spike) e ajuste de teto.
- `tests/test_baixar_requisitorios.py` — testes TDD.

Convenção do projeto: rodar testes com
`& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest`
(de dentro de `C:\Users\DARLANMARTINS\Documents\PROJETO 01`). Python full path; nunca o alias da Store.

⚠️ As tasks de SPIKE e de VERIFICAÇÃO AO VIVO abrem o Edge e exigem login — NÃO rodar
enquanto o lote `baixar_requisitorios.py` estiver em execução (sessão única; pausar o lote antes).

---

## Task 0: SPIKE — descobrir o seletor do texto da prévia (go/no-go)

**Files:** nenhum código de produção ainda; cria `seletores.py:SELETOR_PREVIA_TEXTO` no fim.

Manual/ao vivo (não-TDD). Pausar o lote antes.

- [ ] **Step 1: Rodar um processo de teste com debug e inspecionar o DOM da prévia**

Run:
```powershell
$env:PYTHONIOENCODING="utf-8"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -u baixar_requisitorios.py --test-processar "0090006-26.2015.8.19.0001" "2025.13367-1"
```
Logar quando o Edge abrir. Quando a árvore expandir e começar a clicar nós, observar o painel
direito (a prévia). Em paralelo, com o visualizador aberto, abrir DevTools (F12) no Edge e
inspecionar o contêiner que mostra o texto do documento.

- [ ] **Step 2: Confirmar se a prévia é TEXTO extraível**

Selecionar texto na prévia com o mouse: se seleciona/copia → é texto (GO). Se é imagem/canvas
(não seleciona) → NO-GO. Anotar o seletor CSS do contêiner de texto (ex: `div.documento-conteudo`,
`#pdf-text-layer`, `.textLayer`, etc. — o valor real vem da inspeção).

- [ ] **Step 3: Decisão go/no-go**

- **GO (prévia é texto):** gravar o seletor em `seletores.py`:
```python
# Contêiner do texto renderizado na prévia do visualizador (descoberto no spike 2026-06-09).
# Usado para classificar o documento SEM baixar o PDF.
SELETOR_PREVIA_TEXTO = "<SELETOR_REAL_DESCOBERTO_NO_SPIKE>"
```
Commit:
```bash
git add seletores.py
git commit -m "spike: seletor do texto da previa do visualizador (GO p/ abordagem 2)"
```
Seguir para a Task 1.

- **NO-GO (prévia é imagem):** registrar na memória do projeto que a Abordagem 2 não é viável
  (prévia não-textual) e PARAR — a Abordagem 1 (commit d0790f0, já em produção) permanece.
  Não executar Tasks 1-5.

---

## Task 1: Extratores por TEXTO (refactor TDD)

**Files:**
- Modify: `baixar_requisitorios.py:145-171` (funções `extrair_numero_ofreq`, `extrair_vinculo_ofreq_precatorio`)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever os testes que falham**

Adicionar ao fim de `tests/test_baixar_requisitorios.py`:
```python
# ===== Extratores por TEXTO (classificação pela prévia, sem baixar) =====
from baixar_requisitorios import (
    extrair_numero_ofreq_de_texto,
    extrair_vinculo_ofreq_precatorio_de_texto,
)

_TXT_REQ = "Definitivo OFÍCIO Nº: 2024.16127/OFREQ\nOFÍCIO REQUISITÓRIO DE PAGAMENTO"
_TXT_VINC = ("Ofício 2025.10413/OFREQ foi analisado pelo processo de análise "
             "00015568/2025 e gerou o precatório 2025.16841-6")


def test_extrair_numero_ofreq_de_texto_acha():
    assert extrair_numero_ofreq_de_texto(_TXT_REQ) == "2024.16127"


def test_extrair_numero_ofreq_de_texto_vazio_none():
    assert extrair_numero_ofreq_de_texto("") is None
    assert extrair_numero_ofreq_de_texto("nada aqui") is None


def test_extrair_vinculo_de_texto_acha_par():
    assert extrair_vinculo_ofreq_precatorio_de_texto(_TXT_VINC) == ("2025.10413", "2025.16841-6")


def test_extrair_vinculo_de_texto_vazio_none():
    assert extrair_vinculo_ofreq_precatorio_de_texto("") is None
    assert extrair_vinculo_ofreq_precatorio_de_texto("texto sem vinculo") is None


def test_extrair_numero_ofreq_bytes_ainda_funciona(pdf_modelo_bytes):
    # wrapper por bytes continua delegando ao por-texto (não quebrou)
    from baixar_requisitorios import extrair_numero_ofreq
    assert extrair_numero_ofreq(pdf_modelo_bytes) is None or isinstance(
        extrair_numero_ofreq(pdf_modelo_bytes), str)
```

- [ ] **Step 2: Rodar e ver falhar (ImportError)**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k "de_texto" -q`
Expected: FAIL — `cannot import name 'extrair_numero_ofreq_de_texto'`.

- [ ] **Step 3: Implementar (substituir as funções 145-171)**

```python
def extrair_numero_ofreq_de_texto(texto):
    """Extrai o número OFREQ (ex: '2025.14235') de um texto já extraído, ou None."""
    if not texto:
        return None
    m = REGEX_OFREQ.search(texto)
    return m.group(1) if m else None


def extrair_numero_ofreq(pdf_bytes):
    """Extrai o número OFREQ do texto do PDF, ou None. (delega ao por-texto)"""
    return extrair_numero_ofreq_de_texto(_extrair_texto_pdf(pdf_bytes))
```
e (após o `REGEX_VINCULO`):
```python
def extrair_vinculo_ofreq_precatorio_de_texto(texto):
    """De um texto já extraído do ofício DEPRE/DEPJU, retorna (ofreq, precatorio) ou None."""
    if not texto:
        return None
    m = REGEX_VINCULO.search(texto)
    return (m.group(1), m.group(2)) if m else None


def extrair_vinculo_ofreq_precatorio(pdf_bytes):
    """Do ofício DEPRE/DEPJU (bytes), retorna (ofreq, precatorio) ou None. (delega ao por-texto)"""
    return extrair_vinculo_ofreq_precatorio_de_texto(_extrair_texto_pdf(pdf_bytes))
```

- [ ] **Step 4: Rodar e ver passar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -q`
Expected: PASS (todos, incluindo os antigos).

- [ ] **Step 5: Commit**

```bash
git add baixar_requisitorios.py tests/test_baixar_requisitorios.py
git commit -m "refactor: extratores OFREQ/vinculo por texto (bytes delega ao texto)"
```

---

## Task 2: Leitor do texto da prévia

**Files:**
- Modify: `baixar_requisitorios.py` (nova função `_ler_texto_previa`, perto de `_baixar_e_classificar_um`)

Sem unit test (é I/O de DOM); coberta pela verificação ao vivo (Task 5). Mantê-la mínima.

- [ ] **Step 1: Implementar a função**

```python
async def _ler_texto_previa(visualizador, timeout_ms=4000):
    """Lê o texto renderizado na prévia do visualizador (DOM), SEM baixar o PDF.

    Retorna a string (pode ser "" se a prévia ainda não renderizou texto ou é imagem).
    Usado para classificar o documento de forma barata antes de decidir baixar.
    """
    try:
        loc = visualizador.locator(seletores.SELETOR_PREVIA_TEXTO).first
        await loc.wait_for(state="visible", timeout=timeout_ms)
        return (await loc.inner_text()).strip()
    except Exception:
        return ""
```

- [ ] **Step 2: Commit**

```bash
git add baixar_requisitorios.py
git commit -m "feat: _ler_texto_previa — le texto da previa do visualizador sem baixar"
```

---

## Task 3: Classificação por nó via prévia (com download só p/ requisitório e fallback)

**Files:**
- Modify: `baixar_requisitorios.py` (`_baixar_e_classificar_um`, ~1139-1232)
- Test: `tests/test_baixar_requisitorios.py`

O contrato de retorno e mutação de `_baixar_e_classificar_um` é mantido (retorna `relevante: bool`;
muta `requisitorios`/`vinculos`/`ofreqs_vistos`), então `_baixar_e_classificar_indices` não muda.

- [ ] **Step 1: Escrever os testes que falham**

Adicionar ao fim de `tests/test_baixar_requisitorios.py`:
```python
# ===== Classificação por nó via prévia =====
# vínculo: classifica pela prévia e extrai do TEXTO, sem baixar.
# requisitório: baixa o PDF e extrai. prévia vazia: cai no fallback de download.

def test_classificar_um_vinculo_pela_previa_nao_baixa(monkeypatch):
    """Prévia classifica como vínculo -> extrai OFREQ->precatório do texto, NÃO baixa."""
    baixou = {"n": 0}

    async def fake_previa(visu, **k):
        return ("Ofício 2025.10413/OFREQ foi analisado e gerou o precatório 2025.16841-6")

    async def fake_baixar(*a, **k):
        baixou["n"] += 1
        return b"PDF"
    monkeypatch.setattr(_br_timeout, "_ler_texto_previa", fake_previa)
    monkeypatch.setattr(_br_timeout, "_baixar_pdf_apos_clique", fake_baixar)

    class _El:
        async def evaluate(self, *a, **k): return None
        async def click(self, *a, **k): return None
        async def inner_text(self): return "903 - Petição"
    vinculos = {}
    rel = asyncio.run(_br_timeout._baixar_e_classificar_um(
        _FakeVisu(), _El(), 903, None, [], vinculos, set(),
        log=lambda m: None, numero_processo=None))
    assert vinculos == {"2025.10413": "2025.16841-6"}
    assert baixou["n"] == 0          # vínculo NÃO baixa
    assert rel is True


def test_classificar_um_requisitorio_pela_previa_baixa(monkeypatch, pdf_modelo_bytes):
    """Prévia classifica como requisitório -> BAIXA o PDF e extrai beneficiário."""
    async def fake_previa(visu, **k):
        return "OFÍCIO REQUISITÓRIO DE PAGAMENTO\nDefinitivo OFÍCIO Nº: 2024.16127/OFREQ"

    async def fake_baixar(*a, **k):
        return pdf_modelo_bytes
    monkeypatch.setattr(_br_timeout, "_ler_texto_previa", fake_previa)
    monkeypatch.setattr(_br_timeout, "_baixar_pdf_apos_clique", fake_baixar)

    class _El:
        async def evaluate(self, *a, **k): return None
        async def click(self, *a, **k): return None
        async def inner_text(self): return "654 - Fulano Def"
    reqs = []
    rel = asyncio.run(_br_timeout._baixar_e_classificar_um(
        _FakeVisu(), _El(), 654, _tmp_pasta(), reqs, {}, set(),
        log=lambda m: None, numero_processo="0090006-26.2015.8.19.0001"))
    assert rel is True
    assert len(reqs) == 1            # requisitório coletado (baixou)


def test_classificar_um_previa_vazia_cai_no_download(monkeypatch, pdf_modelo_bytes):
    """Prévia vazia/imagem -> fallback: baixa e classifica pelo PDF."""
    async def fake_previa(visu, **k):
        return ""   # imagem/sem texto

    async def fake_baixar(*a, **k):
        return pdf_modelo_bytes
    monkeypatch.setattr(_br_timeout, "_ler_texto_previa", fake_previa)
    monkeypatch.setattr(_br_timeout, "_baixar_pdf_apos_clique", fake_baixar)

    class _El:
        async def evaluate(self, *a, **k): return None
        async def click(self, *a, **k): return None
        async def inner_text(self): return "570 - Def Roberto"
    reqs = []
    asyncio.run(_br_timeout._baixar_e_classificar_um(
        _FakeVisu(), _El(), 570, _tmp_pasta(), reqs, {}, set(),
        log=lambda m: None, numero_processo="0090006-26.2015.8.19.0001"))
    # pdf_modelo é um requisitório -> coletado via fallback de download
    assert len(reqs) == 1
```
Adicionar este helper de pasta temporária no topo do bloco de testes (após os imports), se ainda não existir:
```python
import tempfile
def _tmp_pasta():
    return tempfile.mkdtemp()
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k "classificar_um" -q`
Expected: FAIL — `_baixar_pdf_apos_clique` não existe / lógica de prévia ausente.

- [ ] **Step 3: Refatorar `_baixar_e_classificar_um`**

Extrair o trecho de download (clique no botão Salvar + expect_download) para um helper, e
reescrever a classificação para usar a prévia primeiro. Substituir a função `_baixar_e_classificar_um`
(linhas ~1139-1232) por:

```python
async def _baixar_pdf_apos_clique(visualizador, idx, pasta_temp, log):
    """Após o nó já ter sido clicado e exibido, baixa o PDF e devolve os bytes (ou None)."""
    botao_salvar = visualizador.locator(seletores.BOTAO_SALVAR_COPIA).first
    try:
        await botao_salvar.wait_for(state="visible", timeout=5000)
    except Exception:
        log("  botao Salvar Copia nao apareceu — pulando")
        return None
    try:
        async with visualizador.expect_download(timeout=30000) as dl_info:
            await botao_salvar.evaluate("(el) => el.click()")
        download = await dl_info.value
        tmp_path = pasta_temp / f"_cand_{idx}_{download.suggested_filename}"
        await download.save_as(tmp_path)
        return tmp_path.read_bytes()
    except Exception as e:
        log(f"  erro download: {str(e)[:120]}")
        return None


def _coletar_requisitorio(pdf_bytes, pasta_temp, requisitorios, vinculos, ofreqs_vistos,
                          log, numero_processo, idx):
    """Classifica/extrai um PDF de requisitório já baixado. Muta acumuladores.
    Retorna True se o PDF é relevante (requisitório/vínculo), False caso contrário."""
    texto = _extrair_texto_pdf(pdf_bytes)
    if _parece_escaneado(texto):
        pasta_manual = pasta_temp.parent / "manual_revisar"
        pasta_manual.mkdir(parents=True, exist_ok=True)
        prefixo = f"{numero_processo}_" if numero_processo else ""
        (pasta_temp / f"_cand_{idx}.pdf").write_bytes(pdf_bytes)
        (pasta_temp / f"_cand_{idx}.pdf").replace(pasta_manual / f"{prefixo}_cand_{idx}.pdf")
        log("  PDF sem texto (provável scan) — manual_revisar")
        return False
    if eh_documento_vinculo(texto):
        v = extrair_vinculo_ofreq_precatorio_de_texto(texto)
        if v:
            vinculos[v[0]] = v[1]
            log(f"  vínculo (PDF): OFREQ {v[0]} -> precatório {v[1]}")
        return True
    if not eh_documento_requisitorio(texto):
        log("  PDF não é requisitório nem vínculo — descartando")
        return False
    ofreq = extrair_numero_ofreq_de_texto(texto)
    if ofreq and ofreq in ofreqs_vistos:
        log(f"  requisitório OFREQ {ofreq} repetido — descartando")
        return True
    benef = extrair_beneficiario_completo(pdf_bytes)
    if benef is None:
        pasta_manual = pasta_temp.parent / "manual_revisar"
        pasta_manual.mkdir(parents=True, exist_ok=True)
        (pasta_temp / f"_cand_{idx}.pdf").write_bytes(pdf_bytes)
        log("  beneficiário não extraído — manual_revisar")
        return True
    adv = extrair_advogado(pdf_bytes)
    if ofreq:
        ofreqs_vistos.add(ofreq)
    requisitorios.append({
        "ofreq": ofreq, "beneficiario": benef, "advogado": adv,
        "pdf_path": pasta_temp / f"_cand_{idx}.pdf", "pdf_bytes": pdf_bytes,
    })
    (pasta_temp / f"_cand_{idx}.pdf").write_bytes(pdf_bytes)
    log(f"  requisitório coletado: OFREQ {ofreq} benef={benef['nome']}")
    return True


async def _baixar_e_classificar_um(visualizador, el, idx, pasta_temp,
                                   requisitorios, vinculos, ofreqs_vistos,
                                   log, numero_processo):
    """Clica UM nó, classifica pela PRÉVIA (sem baixar):
      - vínculo  -> extrai OFREQ->precatório do texto da prévia, SEM baixar;
      - requisitório -> BAIXA o PDF e extrai beneficiário/advogado;
      - prévia vazia/imagem -> FALLBACK: baixa e classifica pelo PDF.
    Muta os acumuladores. Retorna True se o nó é relevante (requisitório/vínculo).
    """
    try:
        try:
            rotulo = (await el.inner_text())[:70].replace("\n", " ").strip()
        except Exception:
            rotulo = "?"
        log(f"candidato idx={idx} rotulo={rotulo!r}: clicando")
        await el.evaluate("(e) => e.scrollIntoView({block: 'center'})")
        await asyncio.sleep(0.5)
        await el.click(force=True, timeout=10000)
        await asyncio.sleep(2)
    except Exception as e:
        log(f"  click falhou: {str(e)[:120]}")
        return False

    texto_previa = await _ler_texto_previa(visualizador)

    # Prévia textual: classifica de graça.
    if texto_previa:
        if eh_documento_vinculo(texto_previa):
            v = extrair_vinculo_ofreq_precatorio_de_texto(texto_previa)
            if v:
                vinculos[v[0]] = v[1]
                log(f"  vínculo (prévia): OFREQ {v[0]} -> precatório {v[1]}")
            return True
        if eh_documento_requisitorio(texto_previa):
            pdf_bytes = await _baixar_pdf_apos_clique(visualizador, idx, pasta_temp, log)
            if not pdf_bytes:
                return False
            return _coletar_requisitorio(pdf_bytes, pasta_temp, requisitorios,
                                         vinculos, ofreqs_vistos, log, numero_processo, idx)
        return False  # prévia textual mas nem vínculo nem requisitório

    # Prévia vazia/imagem -> fallback Abordagem 1: baixa e classifica pelo PDF.
    pdf_bytes = await _baixar_pdf_apos_clique(visualizador, idx, pasta_temp, log)
    if not pdf_bytes:
        return False
    return _coletar_requisitorio(pdf_bytes, pasta_temp, requisitorios,
                                 vinculos, ofreqs_vistos, log, numero_processo, idx)
```

- [ ] **Step 4: Rodar e ver passar (incluindo a suíte toda)**

Run: `... -m pytest tests/test_baixar_requisitorios.py -q`
Expected: PASS. Conferir que os testes antigos de early-stop/goal-stop (que mockam
`_baixar_e_classificar_um`) seguem passando.

- [ ] **Step 5: Commit**

```bash
git add baixar_requisitorios.py tests/test_baixar_requisitorios.py
git commit -m "feat: classificar no por previa (vinculo sem download, requisitorio baixa, fallback)"
```

---

## Task 4: Ajustar o teto da varredura (prévia barata permite alcance maior)

**Files:**
- Modify: `seletores.py` (`LIMITE_FALLBACK_GENERICO`)

- [ ] **Step 1: Subir o teto**

Como cada nó agora custa ~1-2s (prévia) em vez de ~5-8s (download), o alcance do fim pode
ser maior sem estourar o `PROCESSO_TIMEOUT_SEG` (240s) — cobrindo os casos "no meio".
Editar em `seletores.py`:
```python
LIMITE_FALLBACK_GENERICO = 60
```
Atualizar o comentário acima da constante para explicar que o custo por nó caiu (classificação
por prévia) e o goal-stop encerra cedo no caso comum.

- [ ] **Step 2: Rodar a suíte**

Run: `... -m pytest -q`
Expected: PASS (123+).

- [ ] **Step 3: Commit**

```bash
git add seletores.py
git commit -m "tune: teto do fim 25->60 (previa barata permite alcance maior; goal-stop encerra cedo)"
```

---

## Task 5: Verificação AO VIVO (pausar o lote antes)

**Files:** nenhum (validação).

- [ ] **Step 1: Caso 1 alvo (Def Roberto)**

Run:
```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -u baixar_requisitorios.py --test-processar "0090006-26.2015.8.19.0001" "2025.13367-1"
```
Logar. Expected: `status=ok`, `2025.13367-1 - ROBERTO MORENO DE MELO`, e nos logs:
vínculo classificado pela PRÉVIA (sem download), só o requisitório baixado, goal-stop disparou.
Contar nos logs: nós abertos vs nós baixados (esperado: baixados ≈ 1).

- [ ] **Step 2: Caso multi-precatório (Carraro, 4 alvos)**

Run:
```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -u baixar_requisitorios.py --test-processar "0110128-60.2015.8.19.0001" "2025.09513-3,2025.09514-1,2025.09515-0,2025.09516-8"
```
Expected: `status=ok`, 4 beneficiários corretos por precatório, só ~4 requisitórios baixados,
goal-stop dispara. Comparar tempo/nós-baixados com a Abordagem 1.

- [ ] **Step 3: Registrar resultado na memória do projeto**

Atualizar `MEMORY.md`/`project-precatorios.md` com: prévia é textual (GO), seletor usado,
nós-abertos vs baixados medidos, e que a Abordagem 2 substituiu a varredura-por-download.

- [ ] **Step 4: Commit final / merge**

```bash
git add -A
git commit -m "docs: registra verificacao ao vivo da classificacao por previa"
```

---

## Self-Review (preenchido pelo autor do plano)

- **Cobertura da spec:** A (fluxo do fim + goal-stop) → Tasks 3,4,5. B (decisão por nó) → Task 3.
  C (refactors por texto) → Task 1. D (spike go/no-go) → Task 0. E (testes) → Tasks 1,3,5.
- **Placeholders:** o único valor a preencher é `SELETOR_PREVIA_TEXTO` (saída legítima do spike,
  Task 0) — não é placeholder de plano, é dado de runtime descoberto e gravado numa task.
- **Consistência de tipos:** `_baixar_e_classificar_um` mantém assinatura e contrato (retorna bool,
  muta acumuladores) → `_baixar_e_classificar_indices` intacto. Novos: `_ler_texto_previa`,
  `_baixar_pdf_apos_clique`, `_coletar_requisitorio`, `extrair_*_de_texto`. Nomes usados de forma
  idêntica entre tasks.
