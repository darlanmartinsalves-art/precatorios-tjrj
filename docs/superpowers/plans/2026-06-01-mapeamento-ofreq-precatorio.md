# Mapeamento precatório → ofício via OFREQ — Plano de Implementação

> **Para workers agênticos:** SUB-SKILL OBRIGATÓRIA: use superpowers:subagent-driven-development (recomendado) ou superpowers:executing-plans para implementar tarefa a tarefa. Os passos usam checkbox (`- [ ]`).

> **Sem git:** este projeto NÃO é um repositório git. Onde o template pediria "Commit", use **"Checkpoint: rodar a suíte completa"** e confirme verde antes de seguir. (Opcional: `git init` para versionar.)

**Goal:** Preencher as colunas N–R de cada precatório com o beneficiário/advogado CORRETO, casando o ofício requisitório ao precatório pelo número OFREQ (via documento de vínculo DEPRE/DEPJU).

**Architecture:** Em cada processo, baixar todos os ofícios definitivos, classificar cada PDF como *requisitório* (OFREQ→dados) ou *vínculo* (OFREQ→precatório), e fazer join por OFREQ. Função de join é pura e testável; a navegação no browser permanece em `_baixar_e_extrair_pecas_ofreq`/`processar_processo`. Precatório sem vínculo → coluna Status = `REVISAR`.

**Tech Stack:** Python 3.12, pytest, pypdf, openpyxl, playwright.

**Comando da suíte (use sempre este Python):**
`& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" -m pytest -q`

---

### Task 1: Fixture do documento de vínculo

**Files:**
- Create: `tests/fixtures/vinculo_depre_modelo.pdf` (cópia de `C:\Users\DARLANMARTINS\Downloads\00275429720148190001.pdf`)
- Modify: `tests/conftest.py`

- [ ] **Step 1: Copiar o PDF de vínculo para fixtures**

PowerShell:
```powershell
Copy-Item "C:\Users\DARLANMARTINS\Downloads\00275429720148190001.pdf" "C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\fixtures\vinculo_depre_modelo.pdf"
```

- [ ] **Step 2: Adicionar fixture no conftest**

Em `tests/conftest.py`, após o fixture `pdf_modelo_path` (linha ~37), adicionar:
```python
@pytest.fixture
def pdf_vinculo_bytes():
    """Conteúdo binário do ofício DEPRE/DEPJU que liga OFREQ -> precatório."""
    caminho = Path(__file__).parent / "fixtures" / "vinculo_depre_modelo.pdf"
    return caminho.read_bytes()
```

- [ ] **Step 3: Verificar que a fixture carrega**

Run: `& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" -c "from pathlib import Path; p=Path('tests/fixtures/vinculo_depre_modelo.pdf'); print('OK', p.stat().st_size)"`
Expected: `OK` seguido de um tamanho > 0 (≈112000).

- [ ] **Step 4: Checkpoint — rodar a suíte completa**

Run: `& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" -m pytest -q`
Expected: `72 passed` (nada quebrou; fixture nova ainda não usada).

---

### Task 2: `extrair_numero_ofreq`

**Files:**
- Modify: `baixar_requisitorios.py` (junto aos outros extratores, após `extrair_advogado`, ~linha 140)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever o teste que falha**

Adicionar em `tests/test_baixar_requisitorios.py`:
```python
from baixar_requisitorios import extrair_numero_ofreq


def test_extrair_numero_ofreq_do_modelo(pdf_modelo_bytes):
    assert extrair_numero_ofreq(pdf_modelo_bytes) == "2025.14235"


def test_extrair_numero_ofreq_pdf_invalido_retorna_none():
    assert extrair_numero_ofreq(b"nao eh pdf") is None
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" -m pytest tests/test_baixar_requisitorios.py -k extrair_numero_ofreq -v`
Expected: FAIL — `ImportError: cannot import name 'extrair_numero_ofreq'`.

- [ ] **Step 3: Implementar**

Em `baixar_requisitorios.py`, após `extrair_advogado`:
```python
REGEX_OFREQ = re.compile(r"(\d{4}\.\d+)\s*/\s*OFREQ", re.IGNORECASE)


def extrair_numero_ofreq(pdf_bytes):
    """Extrai o número OFREQ (ex: '2025.14235') do texto do PDF, ou None."""
    texto = _extrair_texto_pdf(pdf_bytes)
    if not texto:
        return None
    m = REGEX_OFREQ.search(texto)
    return m.group(1) if m else None
```

- [ ] **Step 4: Rodar e ver passar**

Run: `& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" -m pytest tests/test_baixar_requisitorios.py -k extrair_numero_ofreq -v`
Expected: 2 passed.

- [ ] **Step 5: Checkpoint — suíte completa**

Run: `... -m pytest -q` → Expected: `74 passed`.

---

### Task 3: `extrair_vinculo_ofreq_precatorio`

**Files:**
- Modify: `baixar_requisitorios.py` (após `extrair_numero_ofreq`)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever o teste que falha**

```python
from baixar_requisitorios import extrair_vinculo_ofreq_precatorio


def test_extrair_vinculo_do_modelo(pdf_vinculo_bytes):
    assert extrair_vinculo_ofreq_precatorio(pdf_vinculo_bytes) == ("2025.06478", "2025.06209-0")


def test_extrair_vinculo_pdf_requisitorio_retorna_none(pdf_modelo_bytes):
    # O requisitório NÃO é documento de vínculo (não tem "gerou o precatório")
    assert extrair_vinculo_ofreq_precatorio(pdf_modelo_bytes) is None


def test_extrair_vinculo_pdf_invalido_retorna_none():
    assert extrair_vinculo_ofreq_precatorio(b"nao eh pdf") is None
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k extrair_vinculo -v`
Expected: FAIL — ImportError.

- [ ] **Step 3: Implementar**

```python
REGEX_VINCULO = re.compile(
    r"Of[íi]cio\s+(\d{4}\.\d+)\s*/\s*OFREQ.*?"
    r"gerou\s+o\s+precat[óo]rio\s+(\d{4}\.\d+-\d+)",
    re.DOTALL | re.IGNORECASE,
)


def extrair_vinculo_ofreq_precatorio(pdf_bytes):
    """Do ofício DEPRE/DEPJU, retorna (ofreq, precatorio) ou None.

    Ex: 'Ofício 2025.06478/OFREQ ... gerou o precatório 2025.06209-0'
        -> ('2025.06478', '2025.06209-0')
    """
    texto = _extrair_texto_pdf(pdf_bytes)
    if not texto:
        return None
    m = REGEX_VINCULO.search(texto)
    return (m.group(1), m.group(2)) if m else None
```

- [ ] **Step 4: Rodar e ver passar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k extrair_vinculo -v`
Expected: 3 passed.

- [ ] **Step 5: Checkpoint — suíte completa**

Run: `... -m pytest -q` → Expected: `77 passed`.

---

### Task 4: Classificadores `eh_documento_requisitorio` / `eh_documento_vinculo`

**Files:**
- Modify: `baixar_requisitorios.py` (após `extrair_vinculo_ofreq_precatorio`)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever o teste que falha**

```python
from baixar_requisitorios import (
    eh_documento_requisitorio,
    eh_documento_vinculo,
    _extrair_texto_pdf,
)


def test_classifica_requisitorio(pdf_modelo_bytes):
    texto = _extrair_texto_pdf(pdf_modelo_bytes)
    assert eh_documento_requisitorio(texto) is True
    assert eh_documento_vinculo(texto) is False


def test_classifica_vinculo(pdf_vinculo_bytes):
    texto = _extrair_texto_pdf(pdf_vinculo_bytes)
    assert eh_documento_vinculo(texto) is True
    assert eh_documento_requisitorio(texto) is False
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k classifica -v`
Expected: FAIL — ImportError.

- [ ] **Step 3: Implementar**

```python
REGEX_EH_REQUISITORIO = re.compile(r"OF[ÍI]CIO\s+REQUISIT[ÓO]RIO", re.IGNORECASE)
REGEX_EH_VINCULO = re.compile(r"gerou\s+o\s+precat[óo]rio", re.IGNORECASE)


def eh_documento_requisitorio(texto):
    """True se o texto é de um ofício requisitório (tem dados de beneficiário)."""
    return bool(texto) and bool(REGEX_EH_REQUISITORIO.search(texto))


def eh_documento_vinculo(texto):
    """True se o texto é o ofício DEPRE/DEPJU que liga OFREQ ao precatório."""
    return bool(texto) and bool(REGEX_EH_VINCULO.search(texto))
```

- [ ] **Step 4: Rodar e ver passar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k classifica -v`
Expected: 2 passed.

- [ ] **Step 5: Checkpoint — suíte completa**

Run: `... -m pytest -q` → Expected: `79 passed`.

---

### Task 5: Função pura de join `casar_requisitorios_com_vinculos` (coração da correção)

**Files:**
- Modify: `baixar_requisitorios.py` (antes de `processar_processo`, após `_eh_falha_navegacao`, ~linha 1063)
- Test: `tests/test_baixar_requisitorios.py`

Contrato dos dados de entrada:
- `requisitorios`: `list[dict]`, cada um com `"ofreq": str`, `"beneficiario": {"nome","doc","tipo_doc"}`, `"advogado": {"nome","cpf","oab"} | None`.
- `vinculos`: `dict[str, str]` mapeando `ofreq -> precatorio`.
- `precatorios_do_processo`: `list[str]` (os precatórios filtrados daquele processo).

Saída: `dict[precatorio -> {beneficiario_nome, beneficiario_doc, advogado_nome, advogado_cpf, advogado_oab, status}]`. Precatórios da lista sem dados recebem `{"status": "REVISAR"}`.

- [ ] **Step 1: Escrever o teste que falha**

```python
from baixar_requisitorios import casar_requisitorios_com_vinculos


def _req(ofreq, nome, doc, tipo):
    return {
        "ofreq": ofreq,
        "beneficiario": {"nome": nome, "doc": doc, "tipo_doc": tipo},
        "advogado": {"nome": "ADV " + nome, "cpf": "33394784068", "oab": "RJ1"},
    }


def test_casar_um_requisitorio_com_vinculo():
    reqs = [_req("2025.06478", "CARRARO", "28123344000107", "CNPJ")]
    vinc = {"2025.06478": "2025.06209-0"}
    dados = casar_requisitorios_com_vinculos(reqs, vinc, ["2025.06209-0"])
    assert dados["2025.06209-0"]["beneficiario_nome"] == "CARRARO"
    assert dados["2025.06209-0"]["beneficiario_doc"] == "28.123.344/0001-07"
    assert dados["2025.06209-0"]["advogado_cpf"] == "333.947.840-68"
    assert dados["2025.06209-0"]["status"] == "OK"


def test_casar_multiplos_precatorios_no_processo():
    reqs = [
        _req("2025.001", "AUTOR", "11111111111", "CPF"),
        _req("2025.002", "HONORARIOS LTDA", "22222222000122", "CNPJ"),
    ]
    vinc = {"2025.001": "2025.10-0", "2025.002": "2025.20-1"}
    dados = casar_requisitorios_com_vinculos(reqs, vinc, ["2025.10-0", "2025.20-1"])
    assert dados["2025.10-0"]["beneficiario_nome"] == "AUTOR"
    assert dados["2025.20-1"]["beneficiario_nome"] == "HONORARIOS LTDA"


def test_casar_precatorio_sem_vinculo_vira_revisar():
    reqs = []
    dados = casar_requisitorios_com_vinculos(reqs, {}, ["2025.99-9"])
    assert dados["2025.99-9"] == {"status": "REVISAR"}


def test_casar_requisitorio_sem_advogado():
    reqs = [{"ofreq": "2025.003", "beneficiario": {"nome": "X", "doc": "11111111111", "tipo_doc": "CPF"}, "advogado": None}]
    vinc = {"2025.003": "2025.30-0"}
    dados = casar_requisitorios_com_vinculos(reqs, vinc, ["2025.30-0"])
    assert dados["2025.30-0"]["advogado_nome"] is None
    assert dados["2025.30-0"]["status"] == "OK"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k casar -v`
Expected: FAIL — ImportError.

- [ ] **Step 3: Implementar**

```python
def casar_requisitorios_com_vinculos(requisitorios, vinculos, precatorios_do_processo):
    """Casa cada requisitório ao seu precatório pelo OFREQ.

    Retorna {precatorio: {beneficiario_nome, beneficiario_doc, advogado_nome,
    advogado_cpf, advogado_oab, status}}. Precatórios da lista que não receberam
    nenhum requisitório vinculado ficam com {"status": "REVISAR"}.
    """
    dados = {}
    for req in requisitorios:
        precatorio = vinculos.get(req["ofreq"])
        if not precatorio:
            continue
        benef = req["beneficiario"]
        adv = req.get("advogado")
        dados[precatorio] = {
            "beneficiario_nome": benef["nome"],
            "beneficiario_doc": formatar_doc(benef["doc"], benef["tipo_doc"]),
            "advogado_nome": adv["nome"] if adv else None,
            "advogado_cpf": formatar_doc(adv["cpf"], "CPF") if adv else None,
            "advogado_oab": adv["oab"] if adv else None,
            "status": "OK",
        }
    for precatorio in precatorios_do_processo:
        if precatorio not in dados:
            dados[precatorio] = {"status": "REVISAR"}
    return dados
```

- [ ] **Step 4: Rodar e ver passar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k casar -v`
Expected: 4 passed.

- [ ] **Step 5: Checkpoint — suíte completa**

Run: `... -m pytest -q` → Expected: `83 passed`.

---

### Task 6: Coluna Status (S) em `COLUNAS_NOVAS` e `atualizar_planilha`

**Files:**
- Modify: `baixar_requisitorios.py:361-367` (dict `COLUNAS_NOVAS`) e `baixar_requisitorios.py:392-400` (loop de preenchimento em `atualizar_planilha`)
- Test: `tests/test_baixar_requisitorios.py`

- [ ] **Step 1: Escrever o teste que falha**

```python
def test_atualizar_planilha_grava_status_revisar(xlsx_para_atualizar, tmp_path):
    saida = tmp_path / "saida.xlsx"
    atualizar_planilha(xlsx_para_atualizar, saida, {
        "2025.00001-0": {"status": "REVISAR"},
    })
    from openpyxl import load_workbook
    wb = load_workbook(saida)
    ws = wb.active
    assert ws.cell(row=1, column=19).value == "Status"      # cabeçalho S
    assert ws.cell(row=2, column=19).value == "REVISAR"     # S
    assert ws.cell(row=2, column=14).value is None          # N vazio


def test_atualizar_planilha_grava_status_ok(xlsx_para_atualizar, tmp_path):
    saida = tmp_path / "saida.xlsx"
    atualizar_planilha(xlsx_para_atualizar, saida, {
        "2025.00001-0": {
            "beneficiario_nome": "TECHNE", "beneficiario_doc": "X",
            "advogado_nome": "A", "advogado_cpf": "B", "advogado_oab": "C",
            "status": "OK",
        },
    })
    from openpyxl import load_workbook
    wb = load_workbook(saida)
    ws = wb.active
    assert ws.cell(row=2, column=14).value == "TECHNE"
    assert ws.cell(row=2, column=19).value == "OK"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k "status_revisar or status_ok" -v`
Expected: FAIL — `assert None == "Status"` (coluna 19 ainda não existe).

- [ ] **Step 3: Implementar**

Em `baixar_requisitorios.py`, estender o dict (linhas 361-367):
```python
COLUNAS_NOVAS = {
    14: "Beneficiário Nome",   # N
    15: "Beneficiário Doc",    # O
    16: "Advogado Nome",       # P
    17: "Advogado CPF",        # Q
    18: "Advogado OAB",        # R
    19: "Status",              # S
}
```

No loop de preenchimento de `atualizar_planilha` (após a linha `ws.cell(row=row_idx, column=18).value = info.get("advogado_oab")`, ~linha 400), adicionar:
```python
        ws.cell(row=row_idx, column=19).value = info.get("status")
```

- [ ] **Step 4: Rodar e ver passar**

Run: `... -m pytest tests/test_baixar_requisitorios.py -k "status_revisar or status_ok" -v`
Expected: 2 passed.

- [ ] **Step 5: Checkpoint — suíte completa (regressão dos testes antigos de planilha)**

Run: `... -m pytest -q`
Expected: `85 passed` (os testes antigos de `atualizar_planilha` continuam verdes — eles não passam `status`, então S fica `None`).

---

### Task 7: `_baixar_e_extrair_pecas_ofreq` — coletar todos + classificar + retornar (requisitorios, vinculos)

**Files:**
- Modify: `baixar_requisitorios.py:923-1043` (corpo da função) e `seletores.py:121` (`LIMITE_CANDIDATOS`)

> **Sem teste unitário** (depende do browser). A correção é verificada ao vivo na Task 9. Mantém-se o comportamento de `manual_revisar` para PDFs escaneados.

- [ ] **Step 1: Elevar o teto de candidatos**

Em `seletores.py`, linha 121:
```python
LIMITE_CANDIDATOS = 30
```
(processos multi-precatório têm mais peças: requisitórios + vínculos).

- [ ] **Step 2: Substituir o corpo do loop de download e o retorno**

Na função `_baixar_e_extrair_pecas_ofreq`, trocar a docstring/retorno e o miolo do `for i, el in enumerate(candidatos_unicos):`. O novo corpo da função, da inicialização até o `return`, fica assim (substitui de `requisitorios = []` na ~linha 935 até o `return requisitorios` na ~linha 1043):

```python
    requisitorios = []
    vinculos = {}
    ofreqs_vistos = set()

    # Expandir árvore inteira primeiro (nodes colapsados não estão no DOM)
    log("expandindo arvore...")
    await _expandir_arvore_completa(visualizador)
    await asyncio.sleep(2)

    candidatos_info = await visualizador.evaluate("""
        (padroes) => {
            const nos = Array.from(document.querySelectorAll('mat-nested-tree-node'));
            const idxOf = new Map(nos.map((n, i) => [n, i]));
            const matches = [];
            nos.forEach((n, idx) => {
                const txt = (n.textContent || '').trim();
                const level = parseInt(n.getAttribute('aria-level') || '1');
                if (level < 2) return;
                for (const p of padroes) {
                    if (txt.includes(p)) {
                        const ids = [idx];
                        n.querySelectorAll('mat-nested-tree-node').forEach(c => {
                            const ci = idxOf.get(c);
                            if (ci !== undefined) ids.push(ci);
                        });
                        matches.push({idx, level, padrao: p, ids});
                        break;
                    }
                }
            });
            return matches;
        }
    """, seletores.PADROES_PECA_REQUISITORIO)
    log(f"matches encontrados via JS: {len(candidatos_info)}")

    ordem = _ordenar_candidatos(candidatos_info)
    todos_nos = visualizador.locator(seletores.ARVORE_ITEM_TREEITEM)
    candidatos_unicos = [todos_nos.nth(i) for i in ordem][: seletores.LIMITE_CANDIDATOS]
    log(f"alvos a tentar (nós + filhos): {len(candidatos_unicos)}")

    for i, el in enumerate(candidatos_unicos):
        try:
            log(f"candidato {i}: scrollando + clicando")
            await el.evaluate("(e) => e.scrollIntoView({block: 'center'})")
            await asyncio.sleep(0.5)
            await el.click(force=True, timeout=10000)
            await asyncio.sleep(4)
        except Exception as e:
            log(f"  click falhou: {str(e)[:120]}")
            continue

        botao_salvar = visualizador.locator(seletores.BOTAO_SALVAR_COPIA).first
        try:
            await botao_salvar.wait_for(state="visible", timeout=5000)
        except Exception:
            log(f"  botao Salvar Copia nao apareceu — pulando")
            continue

        try:
            async with visualizador.expect_download(timeout=30000) as dl_info:
                await botao_salvar.evaluate("(el) => el.click()")
            download = await dl_info.value
            tmp_path = pasta_temp / f"_cand_{i}_{download.suggested_filename}"
            await download.save_as(tmp_path)
            pdf_bytes = tmp_path.read_bytes()
            texto = _extrair_texto_pdf(pdf_bytes)
            log(f"  baixado: {tmp_path.stat().st_size} bytes")

            # PDF sem texto extraível = provável scan -> revisão manual
            if _parece_escaneado(texto):
                pasta_manual = pasta_temp.parent / "manual_revisar"
                pasta_manual.mkdir(parents=True, exist_ok=True)
                prefixo = f"{numero_processo}_" if numero_processo else ""
                tmp_path.replace(pasta_manual / f"{prefixo}{tmp_path.name}")
                log(f"  PDF sem texto (provável scan) — manual_revisar")
                continue

            # Documento de vínculo DEPRE/DEPJU: OFREQ -> precatório
            if eh_documento_vinculo(texto):
                v = extrair_vinculo_ofreq_precatorio(pdf_bytes)
                if v:
                    vinculos[v[0]] = v[1]
                    log(f"  vínculo: OFREQ {v[0]} -> precatório {v[1]}")
                tmp_path.unlink(missing_ok=True)  # não é entregável
                continue

            # Não é requisitório -> descartar
            if not eh_documento_requisitorio(texto):
                log(f"  PDF não é requisitório nem vínculo — descartando")
                tmp_path.unlink(missing_ok=True)
                continue

            # Requisitório: dedup por OFREQ
            ofreq = extrair_numero_ofreq(pdf_bytes)
            if ofreq and ofreq in ofreqs_vistos:
                log(f"  requisitório OFREQ {ofreq} repetido — descartando")
                tmp_path.unlink(missing_ok=True)
                continue

            benef = extrair_beneficiario_completo(pdf_bytes)
            if benef is None:
                pasta_manual = pasta_temp.parent / "manual_revisar"
                pasta_manual.mkdir(parents=True, exist_ok=True)
                tmp_path.replace(pasta_manual / tmp_path.name)
                log(f"  beneficiário não extraído — manual_revisar")
                continue
            adv = extrair_advogado(pdf_bytes)
            if ofreq:
                ofreqs_vistos.add(ofreq)
            requisitorios.append({
                "ofreq": ofreq,
                "beneficiario": benef,
                "advogado": adv,
                "pdf_path": tmp_path,
                "pdf_bytes": pdf_bytes,
            })
            log(f"  requisitório coletado: OFREQ {ofreq} benef={benef['nome']}")
            # NÃO dar break — continuar para coletar todos os requisitórios e vínculos
        except Exception as e:
            log(f"  erro download/parse: {e}")
            continue

    log(f"total: {len(requisitorios)} requisitórios, {len(vinculos)} vínculos")
    return requisitorios, vinculos
```

- [ ] **Step 3: Atualizar a docstring da função**

Trocar a docstring inicial (linhas ~924-929) por:
```python
    """No visualizador, baixa todas as peças candidatas e as classifica:
    - requisitório (OFÍCIO REQUISITÓRIO) -> dados de beneficiário/advogado;
    - vínculo DEPRE/DEPJU ('gerou o precatório') -> mapa OFREQ -> precatório.

    Retorna (requisitorios, vinculos):
      requisitorios: list[dict] com ofreq/beneficiario/advogado/pdf_path/pdf_bytes
      vinculos: dict {ofreq: precatorio}
    """
```

- [ ] **Step 4: Checkpoint — suíte completa**

Run: `... -m pytest -q`
Expected: `85 passed` (nenhum teste unitário cobre esta função async, mas a importação do módulo não pode quebrar). Se aparecer `SyntaxError`/`ImportError`, corrigir antes de seguir.

---

### Task 8: `processar_processo` — usar o join e nomear PDFs pelo precatório verdadeiro

**Files:**
- Modify: `baixar_requisitorios.py:1114-1146` (trecho da chamada a `_baixar_e_extrair_pecas_ofreq` até o `return {"status": "ok", ...}`)

> **Sem teste unitário** (browser). Verificado ao vivo na Task 9.

- [ ] **Step 1: Substituir o trecho do download/extração/retorno**

Trocar de `requisitorios = await _baixar_e_extrair_pecas_ofreq(` (linha ~1114) até o `return {"status": "ok", "arquivos": arquivos_finais, "dados": dados}` (linha ~1146) por:

```python
        requisitorios, vinculos = await _baixar_e_extrair_pecas_ofreq(
            visu, pasta_tmp, debug=False, numero_processo=numero_processo)
        if not requisitorios:
            return {"status": "sem_requisitorio", "arquivos": [], "dados": {}}

        pasta_saida = Path(pasta_saida)
        pasta_saida.mkdir(parents=True, exist_ok=True)

        # Nomear cada PDF pelo precatório VERDADEIRO (via vínculo OFREQ)
        arquivos_finais = []
        for req in requisitorios:
            precatorio = vinculos.get(req["ofreq"]) or "SEM_VINCULO"
            destino = gerar_nome_arquivo(
                precatorio, req["beneficiario"]["nome"], pasta_saida)
            req["pdf_path"].replace(destino)
            arquivos_finais.append(destino.name)

        # Join OFREQ -> precatório, montando os dados por precatório
        dados = casar_requisitorios_com_vinculos(
            requisitorios, vinculos, precatorios_do_processo)

        return {"status": "ok", "arquivos": arquivos_finais, "dados": dados}
```

- [ ] **Step 2: Checkpoint — suíte completa**

Run: `... -m pytest -q`
Expected: `85 passed`. Confirma que o módulo importa e nada quebrou.

- [ ] **Step 3: Sanity check de importação/sintaxe**

Run: `& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" -c "import baixar_requisitorios; print('import OK')"`
Expected: `import OK`.

---

### Task 9: Verificação AO VIVO (verification-before-completion)

> Esta tarefa NÃO pode ser pulada. Nenhuma alegação de "funciona" antes de ver a saída real num processo multi-precatório. Requer o Edge logado no Portal de Serviços (perfil Advogado), conforme `COMO_RODAR.md`.

**Files:** nenhum (execução).

- [ ] **Step 1: Subir o Edge logado** (passos 1–2 do `COMO_RODAR.md`).

- [ ] **Step 2: Rodar o pipeline num processo multi-precatório conhecido**

PowerShell (2ª janela):
```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:/Users/DARLANMARTINS/AppData/Local/Programs/Python/Python312/python.exe" baixar_requisitorios.py --test-processar "0027542-97.2014.8.19.0001"
```

- [ ] **Step 3: LER a saída e confirmar evidência**

Critérios de sucesso (todos):
- `status` = `ok`.
- `dados` contém **mais de um precatório** com beneficiários **diferentes** (não o mesmo carimbado).
- Os PDFs em `Downloads\Precatórios_Requisitórios\` foram nomeados com o **precatório verdadeiro** (ex.: `2025.06209-0 - Escritório Carraro...`), não mais `2025.09453-6`.
- Conferir 1 par contra o documento de vínculo: o beneficiário do OFREQ X bate com o precatório que o vínculo aponta.

- [ ] **Step 4: Rodar um lote pequeno end-to-end e inspecionar a planilha**

```powershell
& "C:/.../python.exe" baixar_requisitorios.py --reset --limit 5
```
Depois, inspecionar a planilha de saída:
```powershell
& "C:/.../python.exe" -c "from openpyxl import load_workbook; from pathlib import Path; wb=load_workbook(Path.home()/'Downloads'/'Precatórios 2027 - Atualizado - Etapa2.xlsx'); ws=wb.active; [print(ws.cell(row=r,column=14).value, '|', ws.cell(row=r,column=19).value) for r in range(2,40) if ws.cell(row=r,column=14).value or ws.cell(row=r,column=19).value=='REVISAR']"
```
Expected: linhas com beneficiários distintos por precatório e `Status` = `OK`/`REVISAR` coerente.

- [ ] **Step 5: Checkpoint final — suíte completa**

Run: `... -m pytest -q` → Expected: `85 passed`.

---

## Resumo de cobertura (self-review)

- Extrair OFREQ do requisitório → Task 2 ✅
- Extrair vínculo OFREQ→precatório → Task 3 ✅
- Classificar requisitório vs vínculo → Task 4 ✅
- Join determinístico por OFREQ + REVISAR → Task 5 ✅
- Coluna Status (S) + escrita → Task 6 ✅
- Coletar TODOS os definitivos (sem `break`) + dedup → Task 7 ✅
- Nomear PDF pelo precatório verdadeiro + alimentar join → Task 8 ✅
- Verificação ao vivo (multi-precatório real) → Task 9 ✅

**Contagem de testes esperada ao final:** 85 passed (72 atuais + 2 ofreq + 3 vínculo + 2 classifica + 4 casar + 2 status).
