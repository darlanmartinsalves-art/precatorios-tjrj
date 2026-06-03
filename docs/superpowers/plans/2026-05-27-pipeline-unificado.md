# Pipeline Unificado TJRJ - Plano de Implementação

> **Para agentes:** USAR SUB-SKILL: superpowers:subagent-driven-development. Tasks 1-6 são TDD (subagent-friendly). Task 7 (exploração interativa) **requer usuário no computador** — NÃO delegar. Tasks 8-12 dependem dos selectors da Task 7.

**Goal:** Estender `baixar_requisitorios.py` para um pipeline unificado que conecta no Chrome aberto via CDP, navega TJRJ, baixa PDFs requisitórios, extrai beneficiário+advogado, e atualiza colunas N–R da planilha em uma única passagem por processo (~1.952 processos com saldo ≥ R$ 200k).

**Architecture:** Reutiliza Tasks 1-7 da Etapa 2 anterior (split_cnj, sanitizar, extrair_beneficiario, gerar_nome, filtrar, checkpoint, infra Playwright). Adiciona conexão CDP, extração de advogado, formatação de docs, e escrita na planilha. Mantém async + 3 workers paralelos.

**Tech Stack:** Python 3.12, Playwright (CDP mode), pypdf, openpyxl, pytest.

**Spec:** [docs/superpowers/specs/2026-05-27-pipeline-unificado-design.md](../specs/2026-05-27-pipeline-unificado-design.md)

---

## File Structure

```
C:\Users\DARLANMARTINS\Documents\PROJETO 01\
  baixar_requisitorios.py        # estender (atualmente 36 testes passando)
  seletores.py                   # criar na Task 7 (exploração)
  iniciar_chrome_debug.ps1       # criar na Task 1 (helper)
  tests/test_baixar_requisitorios.py
  tests/conftest.py              # já tem fixtures necessárias
```

**Estado atual (já implementado nas Tasks 1-7 da Etapa 2 anterior):**
- `split_cnj(numero) -> tuple[str, str]`
- `sanitizar_nome(nome) -> str`
- `extrair_beneficiario(pdf_bytes) -> str | None`  ← retorna só o nome
- `gerar_nome_arquivo(precatorio, beneficiario, pasta) -> Path`
- `filtrar_precatorios(caminho_xlsx, saldo_minimo) -> list`
- `carregar_checkpoint_dl(caminho) / salvar_checkpoint_dl(caminho, dados)`
- `abrir_context(playwright, headless)` — persistent context (preservado para retrocompat)
- `verificar_sessao(context) -> bool`
- `login_interativo(context) -> bool`
- 36 testes passando

**Novas funções a adicionar:**
- `extrair_beneficiario_completo(pdf_bytes) -> dict | None`
- `extrair_advogado(pdf_bytes) -> dict | None`
- `formatar_doc(doc: str, tipo: str) -> str`
- `atualizar_planilha(caminho_entrada, caminho_saida, dados)`
- `conectar_chrome_cdp(porta) -> BrowserContext`
- `consultar_processo`, `localizar_pecas_ofreq`, `baixar_peca`, `processar_processo`
- `main()` async + CLI

---

### E2U Task 1: Helper iniciar_chrome_debug.ps1

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\iniciar_chrome_debug.ps1`

Esta task cria um script PowerShell que o usuário roda **antes** do Python para abrir Chrome com debug port. Tarefa simples (sem testes).

- [ ] **Step 1: Criar iniciar_chrome_debug.ps1**

```powershell
<#
.SYNOPSIS
Abre o Chrome com debug port para que o script Python possa se conectar via CDP.

.DESCRIPTION
Lança o Chrome instalado no sistema com --remote-debugging-port=9222 e
um user-data-dir dedicado em C:\temp\chrome_debug. Após abrir, vc loga
manualmente no TJRJ e no confirmeonline. Daí roda o script Python.
#>

$ChromePaths = @(
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:LocalAppData}\Google\Chrome\Application\chrome.exe"
)

$ChromeExe = $ChromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $ChromeExe) {
    Write-Host "ERRO: Chrome não encontrado nos caminhos padrão." -ForegroundColor Red
    Write-Host "Caminhos verificados:" -ForegroundColor Yellow
    $ChromePaths | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$UserDataDir = "C:\temp\chrome_debug"
$DebugPort = 9222

if (-not (Test-Path $UserDataDir)) {
    New-Item -ItemType Directory -Path $UserDataDir -Force | Out-Null
    Write-Host "Pasta de perfil criada: $UserDataDir" -ForegroundColor Green
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Iniciando Chrome com debug port $DebugPort" -ForegroundColor Cyan
Write-Host " Chrome: $ChromeExe" -ForegroundColor Cyan
Write-Host " Perfil: $UserDataDir" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Espere o Chrome abrir" -ForegroundColor Yellow
Write-Host "  2. Faça login no TJRJ (resolva o código verificador)" -ForegroundColor Yellow
Write-Host "  3. Faça login no confirmeonline (se ainda nao logou)" -ForegroundColor Yellow
Write-Host "  4. NÃO feche essa janela do Chrome" -ForegroundColor Yellow
Write-Host "  5. Em outro PowerShell, rode o script Python" -ForegroundColor Yellow
Write-Host ""

& $ChromeExe --remote-debugging-port=$DebugPort --user-data-dir=$UserDataDir
```

- [ ] **Step 2: Smoke test — rodar o script**

```powershell
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& .\iniciar_chrome_debug.ps1
```

Expected:
- Chrome abre (não o bundled chromium, mas o Chrome real instalado)
- Console mostra as instruções
- Chrome fica aberto

Verificar manualmente que o debug port está respondendo:

```powershell
Invoke-WebRequest -Uri "http://localhost:9222/json/version" -UseBasicParsing | Select-Object -ExpandProperty Content
```
Expected: JSON com `"Browser": "Chrome/..."`.

Após esse teste, pode fechar o Chrome (será aberto de novo quando rodar o pipeline).

---

### E2U Task 2: extrair_beneficiario_completo (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

Estende a extração do PDF para também retornar o documento (CPF ou CNPJ). A função antiga `extrair_beneficiario` é mantida (delegando à nova) para preservar os 3 testes existentes.

- [ ] **Step 1: Adicionar testes**

APPEND ao `tests/test_baixar_requisitorios.py`:

```python


from baixar_requisitorios import extrair_beneficiario_completo


def test_extrair_beneficiario_completo_do_modelo(pdf_modelo_bytes):
    resultado = extrair_beneficiario_completo(pdf_modelo_bytes)
    assert resultado == {
        "nome": "TECHNE ENGENHARIA E SISTEMAS LTDA",
        "doc": "50737766000121",
        "tipo_doc": "CNPJ",
    }


def test_extrair_beneficiario_completo_pdf_invalido_retorna_none():
    assert extrair_beneficiario_completo(b"nao eh pdf") is None


def test_extrair_beneficiario_completo_pdf_vazio_retorna_none():
    assert extrair_beneficiario_completo(b"%PDF-1.4\n%%EOF") is None


def test_extrair_beneficiario_legado_ainda_retorna_nome(pdf_modelo_bytes):
    # Função legada continua funcionando
    assert extrair_beneficiario(pdf_modelo_bytes) == "TECHNE ENGENHARIA E SISTEMAS LTDA"
```

- [ ] **Step 2: Confirmar falha**

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k beneficiario_completo
```
Expected: FAIL ImportError.

- [ ] **Step 3: Refatorar extrair_beneficiario + adicionar extrair_beneficiario_completo**

Em `baixar_requisitorios.py`, **substituir** a função existente `extrair_beneficiario` e a constante `REGEX_BENEFICIARIO` por:

```python
REGEX_BENEFICIARIO_NOME = re.compile(
    r"III\s*[-–]\s*BENEFICI[ÁA]RIO.*?Nome:\s*(.+?)(?:\n|CNPJ|CPF)",
    re.DOTALL | re.IGNORECASE,
)
REGEX_BENEFICIARIO_DOC = re.compile(
    r"III\s*[-–]\s*BENEFICI[ÁA]RIO.*?(CPF|CNPJ):\s*([\d./\-]+)",
    re.DOTALL | re.IGNORECASE,
)


def _extrair_texto_pdf(pdf_bytes):
    """Helper: extrai todo o texto de um PDF bytes. Retorna string vazia se falhar."""
    if not pdf_bytes:
        return ""
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
    except Exception:
        return ""
    texto = ""
    for pagina in reader.pages:
        try:
            texto += pagina.extract_text() + "\n"
        except Exception:
            continue
    return texto


def extrair_beneficiario_completo(pdf_bytes):
    """Extrai nome, doc e tipo_doc do beneficiário (seção III) do PDF.

    Retorna dict {"nome": str, "doc": str (só dígitos), "tipo_doc": "CPF"|"CNPJ"}
    ou None se não conseguir.
    """
    texto = _extrair_texto_pdf(pdf_bytes)
    if not texto:
        return None
    m_nome = REGEX_BENEFICIARIO_NOME.search(texto)
    if not m_nome:
        return None
    m_doc = REGEX_BENEFICIARIO_DOC.search(texto)
    if not m_doc:
        return None
    tipo = m_doc.group(1).upper()
    doc_raw = re.sub(r"\D", "", m_doc.group(2))
    return {"nome": m_nome.group(1).strip(), "doc": doc_raw, "tipo_doc": tipo}


def extrair_beneficiario(pdf_bytes):
    """Função legada: retorna apenas o nome do beneficiário. Use extrair_beneficiario_completo
    para obter também CPF/CNPJ.
    """
    completo = extrair_beneficiario_completo(pdf_bytes)
    return completo["nome"] if completo else None
```

- [ ] **Step 4: Rodar todos os testes**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 40 passed (36 existentes + 4 novos).

---

### E2U Task 3: extrair_advogado (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

- [ ] **Step 1: Adicionar testes**

APPEND ao `tests/test_baixar_requisitorios.py`:

```python


from baixar_requisitorios import extrair_advogado


def test_extrair_advogado_do_modelo(pdf_modelo_bytes):
    resultado = extrair_advogado(pdf_modelo_bytes)
    assert resultado == {
        "nome": "ARTUR GARRASTAZU GOMES FERREIRA",
        "cpf": "33394784068",
        "oab": "RJ185918",
    }


def test_extrair_advogado_pdf_invalido_retorna_none():
    assert extrair_advogado(b"nao eh pdf") is None


def test_extrair_advogado_pdf_vazio_retorna_none():
    assert extrair_advogado(b"%PDF-1.4\n%%EOF") is None
```

- [ ] **Step 2: Confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k extrair_advogado
```
Expected: FAIL ImportError.

- [ ] **Step 3: Implementar extrair_advogado**

APPEND ao `baixar_requisitorios.py`:

```python


REGEX_ADVOGADO = re.compile(
    r"VII\s*[-–]\s*ADVOGADO[S]?\s*DO\s*BENEFICI[ÁA]RIO.*?"
    r"([A-Z]{2}\d+)\s*[-–]\s*(.+?)\s*[-–]\s*(\d{11})",
    re.DOTALL | re.IGNORECASE,
)


def extrair_advogado(pdf_bytes):
    """Extrai o primeiro advogado (seção VII) do PDF.

    Formato esperado: "OAB - NOME - CPF" onde OAB é tipo "RJ185918".

    Retorna dict {"nome": str, "cpf": str (só dígitos), "oab": str}
    ou None se não conseguir.
    """
    texto = _extrair_texto_pdf(pdf_bytes)
    if not texto:
        return None
    m = REGEX_ADVOGADO.search(texto)
    if not m:
        return None
    return {
        "oab": m.group(1).strip(),
        "nome": m.group(2).strip(),
        "cpf": m.group(3).strip(),
    }
```

- [ ] **Step 4: Rodar todos os testes**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 43 passed (40 + 3 novos).

---

### E2U Task 4: formatar_doc (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

- [ ] **Step 1: Adicionar testes**

```python


from baixar_requisitorios import formatar_doc


def test_formatar_doc_cpf():
    assert formatar_doc("33394784068", "CPF") == "333.947.840-68"


def test_formatar_doc_cnpj():
    assert formatar_doc("50737766000121", "CNPJ") == "50.737.766/0001-21"


def test_formatar_doc_cpf_ja_formatado_retorna_formatado():
    assert formatar_doc("333.947.840-68", "CPF") == "333.947.840-68"


def test_formatar_doc_tipo_invalido_retorna_doc_cru():
    assert formatar_doc("12345", "OUTRO") == "12345"


def test_formatar_doc_doc_com_tamanho_errado_retorna_doc_cru():
    # CPF com 10 dígitos não casa
    assert formatar_doc("1234567890", "CPF") == "1234567890"
```

- [ ] **Step 2: Confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k formatar_doc
```
Expected: FAIL ImportError.

- [ ] **Step 3: Implementar formatar_doc**

APPEND ao `baixar_requisitorios.py`:

```python


def formatar_doc(doc, tipo):
    """Formata um CPF ou CNPJ.

    CPF: '33394784068' -> '333.947.840-68'
    CNPJ: '50737766000121' -> '50.737.766/0001-21'

    Se já estiver formatado ou tamanho não bater, retorna o doc original.
    """
    if not doc:
        return doc
    digitos = re.sub(r"\D", "", doc)
    if tipo == "CPF" and len(digitos) == 11:
        return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    if tipo == "CNPJ" and len(digitos) == 14:
        return f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"
    return doc
```

- [ ] **Step 4: Rodar todos os testes**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 48 passed (43 + 5 novos).

---

### E2U Task 5: atualizar_planilha (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\conftest.py`

- [ ] **Step 1: APPEND nova fixture xlsx_com_2350**

APPEND ao `tests/conftest.py`:

```python


@pytest.fixture
def xlsx_para_atualizar(tmp_path):
    """xlsx pequeno com 3 precatórios para testar atualizar_planilha."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["Tribunal", "Precatório", "Natureza", "Devedor", "Orçamento",
               "Processo Judicial", "Valor Histórico", "Saldo atualizado",
               "Cedente", "CPF", "Celular", "E-mail", "OBSERVAÇÃO"])
    ws.append(("TJ", "2025.00001-0", "Comum", "ESTADO", 2027,
               "0000001-00.2020.8.19.0001", 100, 250000.00,
               "ALGUEM", "111.111.111-11", "(21) 11111-1111", "a@a.com", ""))
    ws.append(("TJ", "2025.00002-1", "Comum", "ESTADO", 2027,
               "0000002-00.2020.8.19.0001", 100, 500000.00,
               "OUTRO", "222.222.222-22", "(21) 22222-2222", "b@b.com", ""))
    ws.append(("TJ", "2025.00003-2", "Comum", "ESTADO", 2027,
               "0000003-00.2020.8.19.0001", 100, 100000.00,
               "MAIS UM", "333.333.333-33", "", "c@c.com", ""))
    caminho = tmp_path / "planilha_atualizavel.xlsx"
    wb.save(caminho)
    return caminho
```

- [ ] **Step 2: APPEND testes**

```python


from baixar_requisitorios import atualizar_planilha


def test_atualizar_planilha_grava_colunas_N_a_R(xlsx_para_atualizar, tmp_path):
    saida = tmp_path / "saida.xlsx"
    dados = {
        "2025.00001-0": {
            "beneficiario_nome": "TECHNE ENGENHARIA",
            "beneficiario_doc": "50.737.766/0001-21",
            "advogado_nome": "ARTUR GARRASTAZU",
            "advogado_cpf": "333.947.840-68",
            "advogado_oab": "RJ185918",
        }
    }
    atualizar_planilha(xlsx_para_atualizar, saida, dados)

    from openpyxl import load_workbook
    wb = load_workbook(saida)
    ws = wb.active
    assert ws.cell(row=2, column=14).value == "TECHNE ENGENHARIA"    # N
    assert ws.cell(row=2, column=15).value == "50.737.766/0001-21"   # O
    assert ws.cell(row=2, column=16).value == "ARTUR GARRASTAZU"     # P
    assert ws.cell(row=2, column=17).value == "333.947.840-68"       # Q
    assert ws.cell(row=2, column=18).value == "RJ185918"             # R


def test_atualizar_planilha_preserva_entrada(xlsx_para_atualizar, tmp_path):
    import hashlib
    saida = tmp_path / "saida.xlsx"
    hash_antes = hashlib.md5(xlsx_para_atualizar.read_bytes()).hexdigest()
    atualizar_planilha(xlsx_para_atualizar, saida, {"2025.00001-0": {
        "beneficiario_nome": "X", "beneficiario_doc": "Y",
        "advogado_nome": "Z", "advogado_cpf": "W", "advogado_oab": "V",
    }})
    hash_depois = hashlib.md5(xlsx_para_atualizar.read_bytes()).hexdigest()
    assert hash_antes == hash_depois


def test_atualizar_planilha_preserva_colunas_existentes(xlsx_para_atualizar, tmp_path):
    saida = tmp_path / "saida.xlsx"
    atualizar_planilha(xlsx_para_atualizar, saida, {"2025.00001-0": {
        "beneficiario_nome": "X", "beneficiario_doc": "Y",
        "advogado_nome": "Z", "advogado_cpf": "W", "advogado_oab": "V",
    }})
    from openpyxl import load_workbook
    wb = load_workbook(saida)
    ws = wb.active
    # Coluna I=Cedente, J=CPF, K=Celular existentes intactos
    assert ws.cell(row=2, column=9).value == "ALGUEM"
    assert ws.cell(row=2, column=10).value == "111.111.111-11"
    assert ws.cell(row=2, column=11).value == "(21) 11111-1111"


def test_atualizar_planilha_grava_cabecalhos_novos(xlsx_para_atualizar, tmp_path):
    saida = tmp_path / "saida.xlsx"
    atualizar_planilha(xlsx_para_atualizar, saida, {})
    from openpyxl import load_workbook
    wb = load_workbook(saida)
    ws = wb.active
    assert ws.cell(row=1, column=14).value == "Beneficiário Nome"
    assert ws.cell(row=1, column=15).value == "Beneficiário Doc"
    assert ws.cell(row=1, column=16).value == "Advogado Nome"
    assert ws.cell(row=1, column=17).value == "Advogado CPF"
    assert ws.cell(row=1, column=18).value == "Advogado OAB"


def test_atualizar_planilha_ignora_precatorio_inexistente(xlsx_para_atualizar, tmp_path):
    saida = tmp_path / "saida.xlsx"
    # Esse precatório não está na planilha
    atualizar_planilha(xlsx_para_atualizar, saida, {"2099.99999-9": {
        "beneficiario_nome": "X", "beneficiario_doc": "Y",
        "advogado_nome": "Z", "advogado_cpf": "W", "advogado_oab": "V",
    }})
    # Não deve dar erro; só não preenche nada
    from openpyxl import load_workbook
    wb = load_workbook(saida)
    ws = wb.active
    for row in range(2, 5):
        assert ws.cell(row=row, column=14).value is None
```

- [ ] **Step 3: Confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k atualizar_planilha
```
Expected: FAIL ImportError.

- [ ] **Step 4: Implementar atualizar_planilha**

APPEND ao `baixar_requisitorios.py`:

```python


COLUNAS_NOVAS = {
    14: "Beneficiário Nome",   # N
    15: "Beneficiário Doc",    # O
    16: "Advogado Nome",       # P
    17: "Advogado CPF",        # Q
    18: "Advogado OAB",        # R
}


def atualizar_planilha(caminho_entrada, caminho_saida, dados):
    """Carrega planilha de entrada, atualiza colunas N-R, salva em saída.

    `dados` é dict {precatorio: {"beneficiario_nome", "beneficiario_doc",
                                  "advogado_nome", "advogado_cpf", "advogado_oab"}}

    Entrada permanece inalterada.
    """
    wb = load_workbook(caminho_entrada)
    ws = wb.active

    # Cabeçalhos novos (linha 1)
    for col, header in COLUNAS_NOVAS.items():
        ws.cell(row=1, column=col).value = header

    # Mapear precatório -> linha
    mapa = {}
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row and len(row) > 1 and isinstance(row[1], str):
            mapa[row[1]] = row_idx

    # Preencher
    for precatorio, info in dados.items():
        row_idx = mapa.get(precatorio)
        if row_idx is None:
            continue
        ws.cell(row=row_idx, column=14).value = info.get("beneficiario_nome")
        ws.cell(row=row_idx, column=15).value = info.get("beneficiario_doc")
        ws.cell(row=row_idx, column=16).value = info.get("advogado_nome")
        ws.cell(row=row_idx, column=17).value = info.get("advogado_cpf")
        ws.cell(row=row_idx, column=18).value = info.get("advogado_oab")

    wb.save(caminho_saida)
    wb.close()
```

- [ ] **Step 5: Rodar todos os testes**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 53 passed (48 + 5 novos).

---

### E2U Task 6: conectar_chrome_cdp

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

Adiciona a função de conexão via CDP. Não tem unit test convencional (precisa de Chrome rodando) — validação é manual.

- [ ] **Step 1: APPEND conectar_chrome_cdp**

APPEND ao `baixar_requisitorios.py`:

```python


CDP_PORTA_PADRAO = 9222


async def conectar_chrome_cdp(porta=CDP_PORTA_PADRAO):
    """Conecta no Chrome aberto via CDP. Retorna (browser, context).

    Pré-requisito: usuário deve ter aberto Chrome com:
      chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\\temp\\chrome_debug

    Levanta RuntimeError se não conseguir conectar.
    """
    pw = await async_playwright().start()
    try:
        browser = await pw.chromium.connect_over_cdp(f"http://localhost:{porta}")
    except Exception as e:
        await pw.stop()
        raise RuntimeError(
            f"Não foi possível conectar ao Chrome em localhost:{porta}.\n"
            f"Erro: {e}\n"
            f"Verifique se você abriu o Chrome com:\n"
            f"  .\\iniciar_chrome_debug.ps1"
        )
    # Usa o primeiro context disponível (Chrome real geralmente já tem um)
    if not browser.contexts:
        await browser.close()
        await pw.stop()
        raise RuntimeError("Chrome não tem nenhum context. Reabra o Chrome.")
    context = browser.contexts[0]
    return pw, browser, context


async def _test_cdp_manual():
    """Função de teste manual: conecta no Chrome, lista URLs abertas."""
    pw = None
    browser = None
    try:
        pw, browser, context = await conectar_chrome_cdp()
        print(f"Conectado! Pages abertas:")
        for i, page in enumerate(context.pages):
            print(f"  [{i}] {page.url}")
        print(f"Total: {len(context.pages)} page(s)")
        await asyncio.sleep(1)
    finally:
        if browser:
            await browser.close()
        if pw:
            await pw.stop()
```

Atualizar o bloco `if __name__ == "__main__":` para incluir o modo `--test-cdp`:

```python
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 1 and sys.argv[1] == "--test-cdp":
        asyncio.run(_test_cdp_manual())
    else:
        print("Modos disponíveis:")
        print("  --test-login  (com chromium bundled, persistent context)")
        print("  --test-cdp    (conecta no Chrome aberto via debug port)")
        print("(main completo será implementado em task posterior)")
```

- [ ] **Step 2: Confirmar que testes existentes passam**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 53 passed (sem testes novos nesta task).

- [ ] **Step 3: Validação manual — testar conexão CDP**

Pré-requisito: abrir Chrome com debug port em outro terminal:

```powershell
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& .\iniciar_chrome_debug.ps1
```

(Pode deixar o Chrome aberto com qualquer aba.)

Em outro PowerShell, rodar:

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --test-cdp
```

Expected:
- Console imprime "Conectado! Pages abertas:"
- Lista as URLs das abas atualmente abertas no Chrome
- Sai sem erro

---

### E2U Task 7: Exploração interativa da UI (COM O USUÁRIO)

**⚠️ Esta task NÃO deve ser delegada para subagent.** Requer presença do usuário no computador. O Claude controlador conduz.

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\seletores.py`

**Objetivo:** Com o Chrome real aberto (via CDP) e logado no TJRJ, mapear todos os seletores DOM e salvar em `seletores.py`.

- [ ] **Step 1: Garantir Chrome aberto e logado**

Usuário deve ter rodado `.\iniciar_chrome_debug.ps1` e logado no TJRJ. Pode confirmar via `--test-cdp` (Task 6).

- [ ] **Step 2: Criar script temporário _explorar_cdp.py**

```python
"""Script temporário pra explorar o TJRJ via Playwright Inspector.

Conecta no Chrome aberto via CDP e abre o Inspector pra picar elementos.
"""
import asyncio
from playwright.async_api import async_playwright
from baixar_requisitorios import conectar_chrome_cdp, URL_CONSULTA


async def explorar():
    pw, browser, context = await conectar_chrome_cdp()
    try:
        # Abre nova aba pra exploração (não interfere nas existentes)
        page = await context.new_page()
        await page.goto(URL_CONSULTA)
        print(">>> Playwright Inspector aberto. Use 'Pick locator' pra clicar nos elementos.")
        print(">>> Quando terminar de mapear todos, feche o Inspector pra continuar.")
        await page.pause()
        await page.close()
    finally:
        await browser.close()
        await pw.stop()


asyncio.run(explorar())
```

Salvar em `C:\Users\DARLANMARTINS\Documents\PROJETO 01\_explorar_cdp.py`.

- [ ] **Step 3: Rodar exploração + picar elementos**

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" _explorar_cdp.py
```

Com o Inspector aberto, o usuário (guiado pelo Claude controlador) deve identificar e ANOTAR cada um dos seguintes elementos. Para cada um:
1. Clicar "Pick locator" no Inspector
2. Clicar no elemento na página
3. Copiar o locator que aparece no Inspector

Lista de elementos a mapear, usando o processo `0156129-30.2020.8.19.0001` como teste:

  - [ ] Radio "Única" (tipo de numeração)
  - [ ] Campo CNJ prefixo (1º textbox antes do ".8.19.")
  - [ ] Campo CNJ sufixo (2º textbox depois do ".8.19.")
  - [ ] Botão "Pesquisar"
  - [ ] Mensagem "nenhum processo encontrado" (digitar um CNJ inválido para vê-la)
  - [ ] Link/linha do resultado na tabela após pesquisar
  - [ ] Botão "Processo Eletrônico - Visualizador" (na página do processo)
  - [ ] No visualizador: estrutura dos documentos (árvore? lista? filtros?)
  - [ ] Como identificar peças com "OFREQ" ou "REQUISITÓRIO"
  - [ ] Como triggerar download (click direto? botão? URL de PDF?)
  - [ ] Indicador de sessão expirada (campo de senha, ou URL contendo "login")

- [ ] **Step 4: Criar seletores.py com os locators reais**

Substituir os valores **AJUSTAR** abaixo pelos locators reais descobertos:

```python
"""Seletores DOM do portal TJRJ — gerados pela exploração interativa.

Fluxo descoberto:
1. GET URL_CONSULTA
2. Click RADIO_UNICA
3. Fill CAMPO_CNJ_PREFIX
4. Fill CAMPO_CNJ_SUFIX
5. Click BOTAO_PESQUISAR
6. Aguardar LINK_RESULTADO OU MSG_SEM_RESULTADOS
7. Click LINK_RESULTADO -> abre [mesma aba | nova aba: VERIFICAR]
8. Na pagina do processo: click BOTAO_VISUALIZADOR -> abre [VERIFICAR]
9. No visualizador: peças OFREQ identificadas via PECA_OFREQ_LOCATOR
10. Click na peça -> [VERIFICAR como ocorre o download]
"""

URL_CONSULTA = "https://www3.tjrj.jus.br/portalservicos/#/consproc/consultaportal"

# AJUSTAR todos os valores abaixo com locators reais da exploração:
RADIO_UNICA = 'input[type="radio"][value="unica"]'
CAMPO_CNJ_PREFIX = '#numProcUnico'
CAMPO_CNJ_SUFIX = '#numProcOrigem'
BOTAO_PESQUISAR = 'button:has-text("Pesquisar")'
MSG_SEM_RESULTADOS = 'text=/nenhum.*encontrado/i'
LINK_RESULTADO = 'table tbody tr a'

BOTAO_VISUALIZADOR = 'text="Processo Eletrônico - Visualizador"'
PECA_OFREQ_LOCATOR = 'text=/OFREQ|REQUISIT[ÓO]RIO/i'

INDICADOR_LOGIN_EXPIRADO = 'input[type="password"]'

# Comportamento das novas abas
ABRE_NOVA_ABA_APOS_PESQUISAR = True   # AJUSTAR conforme observado
ABRE_NOVA_ABA_VISUALIZADOR = True     # AJUSTAR conforme observado
```

- [ ] **Step 5: Validar import do módulo**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -c "import seletores; print(dir(seletores))"
```
Expected: lista das constantes.

- [ ] **Step 6: Limpar arquivo temporário**

```powershell
Remove-Item "C:\Users\DARLANMARTINS\Documents\PROJETO 01\_explorar_cdp.py" -ErrorAction SilentlyContinue
```

**Saída esperada da Task 7:** `seletores.py` com TODOS os valores reais (não chutes) e o comentário do fluxo descoberto preenchido.

---

### E2U Task 8: consultar_processo (navegação)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

- [ ] **Step 1: Adicionar import e exceção**

Adicionar no topo de `baixar_requisitorios.py`:

```python
import seletores
```

E APPEND uma classe de exceção bem no início do "bloco de funções" (após constantes):

```python


class LoginExpiradoError(Exception):
    """Disparado quando a sessão expirou durante a navegação."""
    pass
```

- [ ] **Step 2: APPEND consultar_processo**

```python


async def consultar_processo(context, numero_processo):
    """Navega de Consulta Processual até o Visualizador do processo.

    Retorna a Page do visualizador, ou None se o processo não foi encontrado.

    Levanta LoginExpiradoError se detectar redirect para login.
    """
    page = await context.new_page()
    try:
        await page.goto(seletores.URL_CONSULTA, wait_until="networkidle", timeout=30000)

        if await page.locator(seletores.INDICADOR_LOGIN_EXPIRADO).count() > 0:
            raise LoginExpiradoError(f"Sessão expirada ao consultar {numero_processo}")

        prefix, sufix = split_cnj(numero_processo)
        await page.locator(seletores.RADIO_UNICA).check()
        await page.locator(seletores.CAMPO_CNJ_PREFIX).fill(prefix)
        await page.locator(seletores.CAMPO_CNJ_SUFIX).fill(sufix)
        await page.locator(seletores.BOTAO_PESQUISAR).click()

        try:
            await page.wait_for_selector(
                f"{seletores.LINK_RESULTADO}, {seletores.MSG_SEM_RESULTADOS}",
                timeout=15000,
            )
        except Exception:
            await page.close()
            return None

        if await page.locator(seletores.MSG_SEM_RESULTADOS).count() > 0:
            await page.close()
            return None

        # Click no resultado abre nova aba (ou mesma — depende do que foi observado)
        if seletores.ABRE_NOVA_ABA_APOS_PESQUISAR:
            async with context.expect_page(timeout=15000) as page_info:
                await page.locator(seletores.LINK_RESULTADO).first.click()
            processo_page = await page_info.value
        else:
            await page.locator(seletores.LINK_RESULTADO).first.click()
            processo_page = page
        await processo_page.wait_for_load_state("networkidle", timeout=30000)

        # Click visualizador abre outra aba
        if seletores.ABRE_NOVA_ABA_VISUALIZADOR:
            async with context.expect_page(timeout=30000) as visu_info:
                await processo_page.locator(seletores.BOTAO_VISUALIZADOR).click()
            visualizador_page = await visu_info.value
        else:
            await processo_page.locator(seletores.BOTAO_VISUALIZADOR).click()
            visualizador_page = processo_page
        await visualizador_page.wait_for_load_state("networkidle", timeout=60000)

        # Fechar abas intermediárias (só se não forem a visualizador)
        if page != visualizador_page:
            await page.close()
        if processo_page != visualizador_page and processo_page != page:
            await processo_page.close()

        return visualizador_page
    except LoginExpiradoError:
        raise
    except Exception:
        if not page.is_closed():
            await page.close()
        raise
```

- [ ] **Step 3: Adicionar modo `--test-consulta` no `__main__`**

Substituir o bloco `if __name__ == "__main__":` por:

```python
async def _test_consulta(numero):
    pw = None
    browser = None
    try:
        pw, browser, context = await conectar_chrome_cdp()
        print(f"Consultando {numero}...")
        visu = await consultar_processo(context, numero)
        if visu is None:
            print("Processo NÃO encontrado.")
        else:
            print(f"Visualizador aberto: {visu.url}")
            input("Inspecione e pressione ENTER para fechar a aba...")
            await visu.close()
    finally:
        if browser:
            await browser.close()
        if pw:
            await pw.stop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 1 and sys.argv[1] == "--test-cdp":
        asyncio.run(_test_cdp_manual())
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-consulta":
        asyncio.run(_test_consulta(sys.argv[2]))
    else:
        print("Modos disponíveis:")
        print("  --test-login")
        print("  --test-cdp")
        print("  --test-consulta NNN")
```

- [ ] **Step 4: Validação manual**

Pré-requisito: Chrome aberto via `.\iniciar_chrome_debug.ps1`, logado no TJRJ.

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --test-consulta "0156129-30.2020.8.19.0001"
```

Expected:
- Nova aba abre no Chrome
- Navega consulta → resultado → visualizador
- Console: "Visualizador aberto: <url do visualizador>"
- Usuário inspeciona, dá ENTER, aba fecha

Se algo não funcionar (timeout, etc), ajustar `seletores.py` e tentar de novo.

---

### E2U Task 9: baixar_peca + processar_processo unificado

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

- [ ] **Step 1: APPEND localizar_pecas_ofreq + baixar_peca + processar_processo**

```python


async def localizar_pecas_ofreq(visualizador_page):
    """Retorna lista de Locators correspondentes às peças OFREQ/REQUISITÓRIO."""
    locator = visualizador_page.locator(seletores.PECA_OFREQ_LOCATOR)
    count = await locator.count()
    return [locator.nth(i) for i in range(count)]


async def baixar_peca(visualizador_page, peca_locator, pasta_temp):
    """Clica numa peça e baixa o PDF. Retorna Path do arquivo temp."""
    pasta_temp.mkdir(parents=True, exist_ok=True)
    async with visualizador_page.expect_download(timeout=60000) as dl_info:
        await peca_locator.click()
    download = await dl_info.value
    destino_tmp = pasta_temp / f"_temp_{download.suggested_filename}"
    await download.save_as(destino_tmp)
    return destino_tmp


def _eh_definitivo(pdf_bytes):
    """Heurística simples: True se o PDF tem 'Definitivo OFÍCIO' no texto."""
    texto = _extrair_texto_pdf(pdf_bytes)
    return "definitivo" in texto.lower() and "ofício" in texto.lower()


async def processar_processo(context, precatorios_do_processo, numero_processo, pasta_saida, pasta_tmp):
    """Pipeline completo para um processo:
       consulta -> baixa peças -> extrai dados -> renomeia.

    `precatorios_do_processo`: lista de números de precatório que apontam para esse processo
    (alguns processos têm múltiplos precatórios; todos recebem os mesmos dados extraídos).

    Retorna dict:
      - status: "ok" | "sem_requisitorio" | "processo_nao_encontrado" | "erro_*"
      - arquivos: list[str]
      - dados: dict {precatorio: {beneficiario_nome, beneficiario_doc, advogado_nome, advogado_cpf, advogado_oab}}
        (preenchido para todos os precatórios do processo, com os mesmos dados extraídos)
    """
    visu = None
    try:
        visu = await consultar_processo(context, numero_processo)
        if visu is None:
            return {"status": "processo_nao_encontrado", "arquivos": [], "dados": {}}

        pecas = await localizar_pecas_ofreq(visu)
        if not pecas:
            return {"status": "sem_requisitorio", "arquivos": [], "dados": {}}

        # Baixar todas as peças
        arquivos_baixados = []  # list de (Path, pdf_bytes)
        for peca in pecas:
            tmp_pdf = await baixar_peca(visu, peca, pasta_tmp)
            pdf_bytes = tmp_pdf.read_bytes()
            arquivos_baixados.append((tmp_pdf, pdf_bytes))

        # Escolher PDF "Definitivo" pra extração; senão, primeiro
        definitivos = [(p, b) for (p, b) in arquivos_baixados if _eh_definitivo(b)]
        escolhido = definitivos[0] if definitivos else arquivos_baixados[0]
        _, pdf_bytes_ref = escolhido

        # Extrair dados do PDF escolhido
        benef = extrair_beneficiario_completo(pdf_bytes_ref)
        adv = extrair_advogado(pdf_bytes_ref)

        if benef is None:
            # Salva PDFs em manual_revisar, sem dados pra planilha
            pasta_manual = pasta_saida / "manual_revisar"
            pasta_manual.mkdir(parents=True, exist_ok=True)
            arquivos_finais = []
            for (tmp_pdf, _) in arquivos_baixados:
                primeiro_prec = precatorios_do_processo[0]
                final = gerar_nome_arquivo(primeiro_prec, "SEM_NOME", pasta_manual)
                tmp_pdf.replace(final)
                arquivos_finais.append(str(final.relative_to(pasta_saida)))
            return {"status": "erro_parsing", "arquivos": arquivos_finais, "dados": {}}

        # Renomear todos os PDFs baixados usando o beneficiário identificado
        nome_benef = benef["nome"]
        arquivos_finais = []
        primeiro_prec = precatorios_do_processo[0]
        for (tmp_pdf, _) in arquivos_baixados:
            final = gerar_nome_arquivo(primeiro_prec, nome_benef, pasta_saida)
            tmp_pdf.replace(final)
            arquivos_finais.append(final.name)

        # Preparar dados de extração — mesmo dict pra todos os precatórios do processo
        info = {
            "beneficiario_nome": nome_benef,
            "beneficiario_doc": formatar_doc(benef["doc"], benef["tipo_doc"]),
            "advogado_nome": adv["nome"] if adv else None,
            "advogado_cpf": formatar_doc(adv["cpf"], "CPF") if adv else None,
            "advogado_oab": adv["oab"] if adv else None,
        }
        dados = {p: info for p in precatorios_do_processo}

        return {"status": "ok", "arquivos": arquivos_finais, "dados": dados}

    except LoginExpiradoError:
        raise
    except Exception as e:
        return {
            "status": "erro_navegacao",
            "arquivos": [],
            "dados": {},
            "motivo": f"{type(e).__name__}: {e}",
        }
    finally:
        if visu is not None and not visu.is_closed():
            await visu.close()
```

- [ ] **Step 2: Adicionar modo `--test-processar`**

Atualizar `__main__`:

```python
async def _test_processar(numero):
    pw = None
    browser = None
    try:
        pw, browser, context = await conectar_chrome_cdp()
        pasta_saida = Path.home() / "Downloads" / "Precatórios_Requisitórios"
        pasta_saida.mkdir(parents=True, exist_ok=True)
        pasta_tmp = PROJETO_DIR / "_tmp_downloads"
        print(f"Processando {numero}...")
        resultado = await processar_processo(
            context,
            precatorios_do_processo=["TESTE"],
            numero_processo=numero,
            pasta_saida=pasta_saida,
            pasta_tmp=pasta_tmp,
        )
        print(f"Resultado: {resultado}")
    finally:
        if browser:
            await browser.close()
        if pw:
            await pw.stop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 1 and sys.argv[1] == "--test-cdp":
        asyncio.run(_test_cdp_manual())
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-consulta":
        asyncio.run(_test_consulta(sys.argv[2]))
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-processar":
        asyncio.run(_test_processar(sys.argv[2]))
    else:
        print("Modos: --test-login | --test-cdp | --test-consulta NNN | --test-processar NNN")
```

- [ ] **Step 3: Validação manual com Garrastazu**

Chrome aberto + logado:

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --test-processar "0156129-30.2020.8.19.0001"
```

Expected:
- Console: `Resultado: {'status': 'ok', 'arquivos': ['TESTE - TECHNE ENGENHARIA E SISTEMAS LTDA.pdf'], 'dados': {'TESTE': {'beneficiario_nome': 'TECHNE ENGENHARIA E SISTEMAS LTDA', 'beneficiario_doc': '50.737.766/0001-21', 'advogado_nome': 'ARTUR GARRASTAZU GOMES FERREIRA', 'advogado_cpf': '333.947.840-68', 'advogado_oab': 'RJ185918'}}}`
- PDF salvo em `Downloads/Precatórios_Requisitórios/TESTE - TECHNE ENGENHARIA E SISTEMAS LTDA.pdf`

---

### E2U Task 10: Orquestração async + CLI

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

- [ ] **Step 1: APPEND argparse + main + worker pool + relatório**

Adicionar imports (se ainda não tem):

```python
import argparse
import csv
from datetime import datetime
```

APPEND ao final:

```python


PASTA_DOWNLOADS = Path.home() / "Downloads"
SAIDA_PDFS_PADRAO = PASTA_DOWNLOADS / "Precatórios_Requisitórios"
ENTRADA_PADRAO = PASTA_DOWNLOADS / "Precatórios 2027 - Atualizado.xlsx"


def caminho_saida_padrao(entrada):
    return entrada.with_name(entrada.stem + " - Etapa2.xlsx")


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Pipeline TJRJ: baixa requisitórios + extrai dados + atualiza planilha.")
    p.add_argument("--entrada", type=Path, default=ENTRADA_PADRAO)
    p.add_argument("--saida", type=Path, default=None)
    p.add_argument("--pasta-pdfs", type=Path, default=SAIDA_PDFS_PADRAO)
    p.add_argument("--saldo-minimo", type=float, default=200000.0)
    p.add_argument("--workers", type=int, default=3)
    p.add_argument("--cdp-port", type=int, default=CDP_PORTA_PADRAO)
    p.add_argument("--limit", type=int, default=None)
    p.add_argument("--reset", action="store_true")
    p.add_argument("--apenas-erros", action="store_true")
    p.add_argument("--processo", type=str, default=None)
    return p.parse_args(argv)


def precisa_processar(processo, resultados, apenas_erros):
    if processo not in resultados:
        return True
    status = resultados[processo].get("status", "")
    if apenas_erros and status != "ok":
        return True
    return status.startswith("erro_")


async def worker_pipeline(nome, fila, resultados, dados_acumulados, contador, lock, context, args, saida_xlsx_ref):
    while True:
        try:
            item = await fila.get()
        except asyncio.CancelledError:
            return
        if item is None:
            fila.task_done()
            return
        numero_processo, precatorios = item
        try:
            pasta_tmp = PROJETO_DIR / "_tmp_downloads" / nome
            resultado = await processar_processo(
                context, precatorios, numero_processo, args.pasta_pdfs, pasta_tmp,
            )
            async with lock:
                # Persistir TUDO no checkpoint (incluindo dados extraídos) para
                # que re-execuções recuperem o estado completo.
                resultados[numero_processo] = {
                    "status": resultado["status"],
                    "arquivos": resultado.get("arquivos", []),
                    "motivo": resultado.get("motivo"),
                    "dados": resultado.get("dados", {}),
                }
                if resultado.get("dados"):
                    dados_acumulados.update(resultado["dados"])
                contador["n"] += 1
                if contador["n"] % 10 == 0:
                    salvar_checkpoint_dl(CHECKPOINT_PATH, resultados)
                    atualizar_planilha(args.entrada, saida_xlsx_ref, dados_acumulados)
                    _print_progresso(contador["n"], contador["total"])
        except LoginExpiradoError:
            print(f"\n[{nome}] LOGIN EXPIRADO. Saindo.")
            raise
        finally:
            fila.task_done()


def _print_progresso(concluidos, total):
    pct = concluidos / total * 100 if total else 0
    barra_size = 30
    preenchido = int(barra_size * concluidos / total) if total else 0
    barra = "█" * preenchido + "░" * (barra_size - preenchido)
    print(f"\r  [{barra}] {concluidos}/{total} ({pct:.1f}%)", end="", flush=True)


async def main():
    args = parse_args()
    if not args.entrada.exists():
        print(f"ERRO: arquivo não encontrado: {args.entrada}", file=sys.stderr)
        return 1

    args.pasta_pdfs.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    saida_xlsx = args.saida or caminho_saida_padrao(args.entrada)

    if args.reset and CHECKPOINT_PATH.exists():
        CHECKPOINT_PATH.unlink()
        print(f"Checkpoint apagado: {CHECKPOINT_PATH}")

    resultados = carregar_checkpoint_dl(CHECKPOINT_PATH)
    print(f"Checkpoint: {len(resultados)} processos")

    print(f"Lendo {args.entrada}...")
    pares = filtrar_precatorios(args.entrada, args.saldo_minimo)
    print(f"  {len(pares)} precatórios com saldo >= R$ {args.saldo_minimo:,.2f}")

    # Agrupar precatórios por processo
    mapa_proc_to_precs = {}
    for prec, proc in pares:
        mapa_proc_to_precs.setdefault(proc, []).append(prec)
    print(f"  {len(mapa_proc_to_precs)} processos únicos")

    if args.processo:
        if args.processo in mapa_proc_to_precs:
            mapa_proc_to_precs = {args.processo: mapa_proc_to_precs[args.processo]}
        else:
            mapa_proc_to_precs = {args.processo: ["DEBUG"]}

    pendentes = [
        (proc, precs) for proc, precs in mapa_proc_to_precs.items()
        if precisa_processar(proc, resultados, args.apenas_erros)
    ]
    if args.limit:
        pendentes = pendentes[: args.limit]
    print(f"  {len(pendentes)} pendentes para processar")

    if not pendentes:
        print("Nada a fazer.")
        return 0

    # Reconstruir dados_acumulados a partir do checkpoint
    # (worker_pipeline persiste "dados" no checkpoint, então recuperamos aqui)
    dados_acumulados = {}
    for proc, r in resultados.items():
        if r.get("status") == "ok" and proc in mapa_proc_to_precs:
            dados_proc = r.get("dados", {})
            dados_acumulados.update(dados_proc)
    print(f"  {len(dados_acumulados)} precatórios já com dados extraídos do checkpoint")

    fila = asyncio.Queue()
    for item in pendentes:
        fila.put_nowait(item)
    for _ in range(args.workers):
        fila.put_nowait(None)

    contador = {"n": 0, "total": len(pendentes)}
    lock = asyncio.Lock()
    inicio = datetime.now()

    pw = None
    browser = None
    try:
        pw, browser, context = await conectar_chrome_cdp(args.cdp_port)
        print(f"Iniciando {args.workers} workers...")
        tasks = [
            asyncio.create_task(worker_pipeline(
                f"w{i}", fila, resultados, dados_acumulados,
                contador, lock, context, args, saida_xlsx,
            ))
            for i in range(args.workers)
        ]
        try:
            await asyncio.gather(*tasks)
        except LoginExpiradoError:
            print("\nSessão expirou. Re-logue no Chrome e re-execute o script.")
            return 2
        finally:
            salvar_checkpoint_dl(CHECKPOINT_PATH, resultados)
            atualizar_planilha(args.entrada, saida_xlsx, dados_acumulados)
    finally:
        if browser:
            await browser.close()
        if pw:
            await pw.stop()

    print()
    _print_relatorio(mapa_proc_to_precs, resultados, saida_xlsx, args.pasta_pdfs, inicio)
    return 0


def _print_relatorio(mapa, resultados, saida_xlsx, pasta_pdfs, inicio):
    ok = sem_req = nao_enc = erro_rede = erro_nav = erro_pars = 0
    erros_detalhe = []
    total_pdfs = 0
    for proc in mapa:
        r = resultados.get(proc, {})
        status = r.get("status", "ausente")
        if status == "ok":
            ok += 1
            total_pdfs += len(r.get("arquivos", []))
        elif status == "sem_requisitorio":
            sem_req += 1
        elif status == "processo_nao_encontrado":
            nao_enc += 1
        elif status == "erro_rede":
            erro_rede += 1
            erros_detalhe.append((proc, status, r.get("motivo", "")))
        elif status == "erro_navegacao":
            erro_nav += 1
            erros_detalhe.append((proc, status, r.get("motivo", "")))
        elif status == "erro_parsing":
            erro_pars += 1
            erros_detalhe.append((proc, status, r.get("motivo", "")))

    decorrido = (datetime.now() - inicio).total_seconds()
    print()
    print("=" * 60)
    print(f" CONCLUÍDO  ({int(decorrido//60)}min{int(decorrido%60)}s)")
    print(f" Processos OK              : {ok}  ({total_pdfs} PDFs)")
    print(f" Sem requisitório          : {sem_req}")
    print(f" Processo não encontrado   : {nao_enc}")
    print(f" Erros de rede             : {erro_rede}")
    print(f" Erros de navegação        : {erro_nav}")
    print(f" Erros de parsing          : {erro_pars}")
    print(f" Planilha de saída         : {saida_xlsx}")
    print(f" Pasta de PDFs             : {pasta_pdfs}")
    print("=" * 60)
    if erros_detalhe:
        with open(ERROS_CSV, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Processo", "Status", "Motivo"])
            w.writerows(erros_detalhe)
        print(f" Log de erros: {ERROS_CSV}")
```

- [ ] **Step 2: Atualizar `__main__` para chamar `main()` quando sem args de teste**

```python
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 1 and sys.argv[1] == "--test-cdp":
        asyncio.run(_test_cdp_manual())
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-consulta":
        asyncio.run(_test_consulta(sys.argv[2]))
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-processar":
        asyncio.run(_test_processar(sys.argv[2]))
    else:
        sys.exit(asyncio.run(main()))
```

- [ ] **Step 3: Confirmar testes ainda passam**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 53 passed.

- [ ] **Step 4: Smoke test --help**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --help
```
Expected: lista das flags.

---

### E2U Task 11: Validação --limit 5

**Files:** nenhum

- [ ] **Step 1: Garantir Chrome aberto + logado**

Pré-requisito: Chrome aberto via `iniciar_chrome_debug.ps1`, logado no TJRJ.

- [ ] **Step 2: Rodar com limite 5**

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --limit 5 --reset
```

Expected:
- Console: progresso de 0/5 → 5/5
- 5 PDFs salvos em `Downloads/Precatórios_Requisitórios/`
- Planilha gerada: `Downloads/Precatórios 2027 - Atualizado - Etapa2.xlsx`
- Tempo: ~2-5 minutos
- Relatório no final

- [ ] **Step 3: Verificar planilha de saída**

Abrir a planilha gerada e conferir:
- Colunas A-M preservadas (Cedente, CPF, Celular etc. intactas)
- Cabeçalhos N-R adicionados: Beneficiário Nome, Beneficiário Doc, Advogado Nome, Advogado CPF, Advogado OAB
- 5 linhas preenchidas nas novas colunas (das 5 processadas)

- [ ] **Step 4: Verificar 1 PDF baixado**

Abrir 1 dos PDFs em `Downloads/Precatórios_Requisitórios/` e conferir:
- Cabeçalho "OFÍCIO REQUISITÓRIO"
- Beneficiário do PDF == valor da coluna N da planilha pra esse precatório
- CPF/CNPJ formatado igual à coluna O

- [ ] **Step 5: Teste de retomada**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --limit 5
```
Expected: "Checkpoint: 5 processos" e "0 pendentes" → sai.

---

### E2U Task 12: Execução completa

**Files:** nenhum

- [ ] **Step 1: Iniciar execução**

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py
```

Expected:
- Conecta no Chrome
- Filtra 2.350 precatórios → 1.952 processos únicos
- 1.952 - 5 (do teste) = ~1.947 pendentes
- 3 workers ativos
- Tempo: ~4-7h
- Salva checkpoint + planilha a cada 10 processos

- [ ] **Step 2: Monitorar sem interromper**

Verificar periodicamente:
- Console progresso
- `Downloads/Precatórios_Requisitórios/` crescendo
- `Downloads/Precatórios 2027 - Atualizado - Etapa2.xlsx` sendo atualizado

- [ ] **Step 3: Relatório final + auditoria**

```powershell
$pasta = "C:\Users\DARLANMARTINS\Downloads\Precatórios_Requisitórios"
$pdfs = Get-ChildItem $pasta -Filter *.pdf -Recurse
Write-Host "Total PDFs: $($pdfs.Count)"
```

Auditar a planilha:

```powershell
$env:PYTHONIOENCODING = "utf-8"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -c "
from openpyxl import load_workbook
wb = load_workbook(r'C:\Users\DARLANMARTINS\Downloads\Precatórios 2027 - Atualizado - Etapa2.xlsx', data_only=True)
ws = wb.active
preenchidas = 0
total = 0
for r in ws.iter_rows(min_row=2, min_col=2, max_col=14, values_only=True):
    if r[0] and isinstance(r[0], str):
        total += 1
        if r[12] is not None:  # coluna N
            preenchidas += 1
print(f'Linhas com beneficiário: {preenchidas} / {total}')
"
```

- [ ] **Step 4: Re-tentar erros se houver**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --apenas-erros
```

---

## Critérios de aceite (do spec)

- [x] Conecta no Chrome aberto via CDP — Task 6
- [x] Filtra >= R$ 200k → ~1.952 processos — Task 5 (já feito em Etapa 2 anterior, reusado)
- [x] PDF baixado + dados extraídos para cada processo — Task 9
- [x] Planilha saída tem colunas N–R preenchidas — Task 5 (atualizar_planilha) + Task 10 (orquestração)
- [x] Planilha entrada inalterada — Task 5 (load + write em arquivo separado)
- [x] CPF/CNPJ formatados — Task 4 (formatar_doc)
- [x] Múltiplos requisitórios numerados — Task 9 usa gerar_nome_arquivo que já lida com isso
- [x] Ctrl+C salva checkpoint + planilha — Task 10 (try/finally em main)
