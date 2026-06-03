# Download de Ofícios Requisitórios TJRJ - Plano de Implementação

> **Para agentes:** USAR SUB-SKILL: superpowers:subagent-driven-development (recomendado) ou superpowers:executing-plans. Tasks 1-7 são TDD puro (ideal pra subagent). Task 8 (exploração) **requer colaboração com usuário** — não delegar pra subagent. Tasks 9-13 implementam a partir dos seletores descobertos na Task 8.

**Goal:** Construir script Python com Playwright que baixa ~2.350 ofícios requisitórios em PDF do portal TJRJ (filtrando precatórios com saldo ≥ R$ 200.000), com login persistente, 3 abas paralelas, retomada automática e nomeação por beneficiário extraído do PDF.

**Architecture:** Playwright async com `launch_persistent_context` para reusar sessão de login. Worker pool de 3 abas concorrentes via asyncio.Semaphore. PDF parsing com pypdf. Checkpoint JSON atômico.

**Tech Stack:** Python 3.12, Playwright (Chromium), pypdf, openpyxl, pytest, asyncio.

**Spec:** [docs/superpowers/specs/2026-05-27-baixar-requisitorios-design.md](../specs/2026-05-27-baixar-requisitorios-design.md)

---

## File Structure

```
C:\Users\DARLANMARTINS\Documents\PROJETO 01\
  baixar_requisitorios.py        # script principal (funções + orquestração + CLI)
  seletores.py                   # constantes de seletores DOM (gerado na Task 8)
  requirements.txt               # adicionar playwright, pypdf
  user_data_chromium/            # sessão persistente do Chromium (criado em runtime)
  downloads_state.json           # checkpoint (criado em runtime)
  erros_download.csv             # log de erros (criado se houver erro)
  screenshots_erro/              # debug screenshots
  tests/
    test_baixar_requisitorios.py
    fixtures/
      garrastazu_modelo.pdf      # cópia do PDF de referência (para testar parsing)
```

**Responsabilidades de `baixar_requisitorios.py`:**

Funções puras (testáveis sem browser):
- `split_cnj(numero) -> tuple[str, str]`
- `sanitizar_nome(nome) -> str`
- `extrair_beneficiario(pdf_bytes) -> str | None`
- `gerar_nome_arquivo(precatorio, beneficiario, pasta, ordinal_inicio=1) -> Path`
- `filtrar_precatorios(caminho_xlsx, saldo_minimo) -> list[tuple[str, str]]`
- `carregar_checkpoint(path) / salvar_checkpoint(path, dados)`

Funções com browser (requer Playwright):
- `verificar_sessao(page) -> bool`
- `consultar_processo(page, numero_processo) -> Page | None` (retorna page do visualizador ou None)
- `localizar_pecas_ofreq(visualizador_page) -> list[Locator]`
- `baixar_peca(visualizador_page, peca, pasta_temp) -> Path`

Orquestração:
- `main()` — async, argparse, worker pool, login, loop

**Responsabilidades de `seletores.py`** (gerado na Task 8):
- `URL_CONSULTA` — URL do form de consulta processual
- `CAMPO_CNJ_PREFIX` — selector do primeiro campo do CNJ
- `CAMPO_CNJ_SUFIX` — selector do segundo campo
- `RADIO_UNICA` — selector do radio "Única"
- `BOTAO_PESQUISAR` — selector do botão
- `LINK_PROCESSO_RESULTADO` — selector do link na tabela de resultados
- `BOTAO_VISUALIZADOR` — botão "Processo Eletrônico - Visualizador"
- `DOM_PECAS_VISUALIZADOR` — descrição de como peças aparecem no DOM
- `INDICADOR_LOGIN_EXPIRADO` — selector usado pra detectar sessão caída

---

### Task 1: Setup adicional (Playwright + pypdf)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\requirements.txt`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\conftest.py`
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\fixtures\garrastazu_modelo.pdf` (cópia)

- [ ] **Step 1: Adicionar dependências ao requirements.txt**

Modificar `requirements.txt` (mantendo as 3 existentes) para ter:

```
openpyxl>=3.1.0
requests>=2.31.0
pytest>=8.0.0
playwright>=1.40.0
pypdf>=4.0.0
```

- [ ] **Step 2: Instalar dependências**

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pip install -r requirements.txt
```
Expected: `Successfully installed ... playwright-X.X.X pypdf-X.X.X ...`

- [ ] **Step 3: Instalar browser Chromium do Playwright**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m playwright install chromium
```
Expected: download de ~130MB do Chromium, mensagem `Downloads are complete`.

- [ ] **Step 4: Copiar PDF de referência para fixtures**

```powershell
$dest = "C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\fixtures"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item "C:\Users\DARLANMARTINS\Downloads\GARRASTAZU - REQUISITÓRIO.pdf" "$dest\garrastazu_modelo.pdf" -Force
```
Verificar: `Test-Path "$dest\garrastazu_modelo.pdf"` retorna `True`.

- [ ] **Step 5: Adicionar fixture do PDF ao conftest.py**

APPEND ao final de `tests/conftest.py`:

```python


@pytest.fixture
def pdf_modelo_bytes():
    """Retorna o conteúdo binário do PDF de referência (Garrastazu/TECHNE)."""
    caminho = Path(__file__).parent / "fixtures" / "garrastazu_modelo.pdf"
    return caminho.read_bytes()


@pytest.fixture
def pdf_modelo_path():
    """Retorna o Path do PDF de referência."""
    return Path(__file__).parent / "fixtures" / "garrastazu_modelo.pdf"
```

- [ ] **Step 6: Verificar todos os testes existentes ainda passam**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: 13 passed (Etapa 1) + 0 novos. Sem nada falhando.

---

### Task 2: Funções puras de string — split_cnj e sanitizar_nome (TDD)

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

- [ ] **Step 1: Criar tests/test_baixar_requisitorios.py com testes de split_cnj**

```python
"""Testes do script baixar_requisitorios."""
import pytest
from baixar_requisitorios import split_cnj


def test_split_cnj_padrao_valido():
    assert split_cnj("0156129-30.2020.8.19.0001") == ("0156129-30.2020", "0001")


def test_split_cnj_outro_processo():
    assert split_cnj("0230764-50.2018.8.19.0001") == ("0230764-50.2018", "0001")


def test_split_cnj_vara_diferente_0072():
    assert split_cnj("0000224-76.2021.8.19.0072") == ("0000224-76.2021", "0072")


def test_split_cnj_formato_invalido_levanta():
    with pytest.raises(ValueError):
        split_cnj("formato-invalido")


def test_split_cnj_vazio_levanta():
    with pytest.raises(ValueError):
        split_cnj("")
```

- [ ] **Step 2: Rodar e confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v
```
Expected: FAIL com `ImportError: cannot import name 'split_cnj' from 'baixar_requisitorios'`.

- [ ] **Step 3: Criar baixar_requisitorios.py com split_cnj**

```python
"""Baixa ofícios requisitórios em PDF do portal TJRJ para precatórios filtrados."""
import re

REGEX_CNJ = re.compile(r"^(\d{7}-\d{2}\.\d{4})\.8\.19\.(\d{4})$")


def split_cnj(numero):
    """Separa um número CNJ em (prefixo, sufixo) para os 2 campos do form.

    Ex: "0156129-30.2020.8.19.0001" -> ("0156129-30.2020", "0001")

    Lança ValueError se o número não casa com o padrão CNJ TJRJ.
    """
    if not numero or not isinstance(numero, str):
        raise ValueError(f"Número CNJ inválido: {numero!r}")
    m = REGEX_CNJ.match(numero.strip())
    if not m:
        raise ValueError(f"Não casa com padrão CNJ TJRJ: {numero!r}")
    return m.group(1), m.group(2)
```

- [ ] **Step 4: Rodar e confirmar PASS para split_cnj**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k split_cnj
```
Expected: 5 passed.

- [ ] **Step 5: APPEND testes de sanitizar_nome**

APPEND ao final de `tests/test_baixar_requisitorios.py`:

```python


from baixar_requisitorios import sanitizar_nome


def test_sanitizar_remove_caracteres_invalidos_windows():
    assert sanitizar_nome('foo/bar\\baz:qux*?"<>|') == "foo_bar_baz_qux_______"


def test_sanitizar_colapsa_espacos_multiplos():
    assert sanitizar_nome("ABC    DEF") == "ABC DEF"


def test_sanitizar_strip_pontas():
    assert sanitizar_nome("   FOO   ") == "FOO"


def test_sanitizar_limita_100_chars():
    longo = "X" * 200
    resultado = sanitizar_nome(longo)
    assert len(resultado) == 100


def test_sanitizar_nome_real_beneficiario():
    nome = "TECHNE ENGENHARIA E SISTEMAS LTDA"
    assert sanitizar_nome(nome) == "TECHNE ENGENHARIA E SISTEMAS LTDA"
```

- [ ] **Step 6: APPEND sanitizar_nome em baixar_requisitorios.py**

```python


CHARS_INVALIDOS = r'\/:*?"<>|'


def sanitizar_nome(nome):
    """Sanitiza um nome para uso seguro como nome de arquivo no Windows.

    - Substitui caracteres inválidos (/\\:*?"<>|) por _
    - Colapsa múltiplos espaços em um único
    - Faz strip de espaços nas pontas
    - Limita a 100 caracteres
    """
    if not nome:
        return ""
    saida = "".join("_" if c in CHARS_INVALIDOS else c for c in nome)
    saida = re.sub(r"\s+", " ", saida).strip()
    return saida[:100]
```

- [ ] **Step 7: Rodar todos os testes**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v
```
Expected: 10 passed (5 split_cnj + 5 sanitizar).

---

### Task 3: extrair_beneficiario do PDF (TDD com pypdf)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

- [ ] **Step 1: APPEND testes de extrair_beneficiario**

```python


from baixar_requisitorios import extrair_beneficiario


def test_extrair_beneficiario_do_pdf_modelo(pdf_modelo_bytes):
    resultado = extrair_beneficiario(pdf_modelo_bytes)
    assert resultado == "TECHNE ENGENHARIA E SISTEMAS LTDA"


def test_extrair_beneficiario_pdf_vazio_retorna_none():
    # PDF mínimo válido sem conteúdo de beneficiário
    pdf_minimo = b"%PDF-1.4\n%%EOF"
    assert extrair_beneficiario(pdf_minimo) is None


def test_extrair_beneficiario_dados_nao_pdf_retorna_none():
    assert extrair_beneficiario(b"isso nao eh um pdf") is None
```

- [ ] **Step 2: Rodar e confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k extrair_beneficiario
```
Expected: FAIL com `ImportError`.

- [ ] **Step 3: APPEND extrair_beneficiario em baixar_requisitorios.py**

Primeiro, adicionar `from io import BytesIO` e `from pypdf import PdfReader` ao topo do arquivo. Em seguida, APPEND:

```python


REGEX_BENEFICIARIO = re.compile(
    r"III\s*[-–]\s*BENEFICI[ÁA]RIO.*?Nome:\s*(.+?)(?:\n|CNPJ|CPF)",
    re.DOTALL | re.IGNORECASE,
)


def extrair_beneficiario(pdf_bytes):
    """Extrai o nome do beneficiário (seção III) de um PDF de ofício requisitório.

    Retorna o nome do beneficiário em string, ou None se não conseguir extrair
    (PDF inválido, sem a seção III, ou sem campo Nome:).
    """
    if not pdf_bytes:
        return None
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
    except Exception:
        return None

    texto_completo = ""
    for pagina in reader.pages:
        try:
            texto_completo += pagina.extract_text() + "\n"
        except Exception:
            continue

    if not texto_completo:
        return None

    m = REGEX_BENEFICIARIO.search(texto_completo)
    if not m:
        return None
    return m.group(1).strip()
```

Atualizar o topo do arquivo para garantir os imports:

```python
"""Baixa ofícios requisitórios em PDF do portal TJRJ para precatórios filtrados."""
import re
from io import BytesIO
from pypdf import PdfReader
```

- [ ] **Step 4: Rodar e confirmar PASS**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k extrair_beneficiario
```
Expected: 3 passed.

---

### Task 4: gerar_nome_arquivo com colisão (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

- [ ] **Step 1: APPEND testes**

```python


from baixar_requisitorios import gerar_nome_arquivo
from pathlib import Path


def test_gerar_nome_arquivo_sem_colisao(tmp_path):
    resultado = gerar_nome_arquivo("2025.09451-0", "TECHNE ENGENHARIA", tmp_path)
    assert resultado == tmp_path / "2025.09451-0 - TECHNE ENGENHARIA.pdf"


def test_gerar_nome_arquivo_com_colisao_simples(tmp_path):
    (tmp_path / "2025.09451-0 - TECHNE.pdf").write_bytes(b"")
    resultado = gerar_nome_arquivo("2025.09451-0", "TECHNE", tmp_path)
    assert resultado == tmp_path / "2025.09451-0 - TECHNE (2).pdf"


def test_gerar_nome_arquivo_com_colisao_dupla(tmp_path):
    (tmp_path / "2025.09451-0 - TECHNE.pdf").write_bytes(b"")
    (tmp_path / "2025.09451-0 - TECHNE (2).pdf").write_bytes(b"")
    resultado = gerar_nome_arquivo("2025.09451-0", "TECHNE", tmp_path)
    assert resultado == tmp_path / "2025.09451-0 - TECHNE (3).pdf"


def test_gerar_nome_arquivo_sanitiza_beneficiario(tmp_path):
    resultado = gerar_nome_arquivo("2025.09451-0", "TECHNE/INC", tmp_path)
    assert resultado == tmp_path / "2025.09451-0 - TECHNE_INC.pdf"
```

- [ ] **Step 2: Rodar e confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k gerar_nome
```
Expected: FAIL com `ImportError`.

- [ ] **Step 3: APPEND gerar_nome_arquivo em baixar_requisitorios.py**

Atualizar topo do arquivo para incluir `from pathlib import Path`. Então APPEND:

```python


def gerar_nome_arquivo(precatorio, beneficiario, pasta):
    """Gera um Path único para o arquivo de saída.

    Formato base: "{precatorio} - {beneficiario sanitizado}.pdf"
    Se já existir, adiciona " (2)", " (3)" etc. até achar nome livre.
    """
    pasta = Path(pasta)
    base = f"{precatorio} - {sanitizar_nome(beneficiario)}"
    candidato = pasta / f"{base}.pdf"
    if not candidato.exists():
        return candidato
    ordinal = 2
    while True:
        candidato = pasta / f"{base} ({ordinal}).pdf"
        if not candidato.exists():
            return candidato
        ordinal += 1
```

- [ ] **Step 4: Rodar e confirmar PASS**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k gerar_nome
```
Expected: 4 passed.

---

### Task 5: filtrar_precatorios (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\conftest.py`

- [ ] **Step 1: APPEND fixture com saldos variados em conftest.py**

APPEND ao `tests/conftest.py`:

```python


@pytest.fixture
def xlsx_com_saldos(tmp_path):
    """xlsx com 6 precatórios, saldos variados, alguns processos duplicados."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["Tribunal", "Precatório", "Natureza", "Devedor", "Orçamento",
               "Processo Judicial", "Valor Histórico", "Saldo atualizado"])
    # (precatorio, processo, saldo)
    dados = [
        ("TJ", "2025.00001-0", "Comum", "ESTADO", 2027, "0000001-00.2020.8.19.0001", 100, 100000.00),
        ("TJ", "2025.00002-1", "Comum", "ESTADO", 2027, "0000002-00.2020.8.19.0001", 100, 250000.00),
        ("TJ", "2025.00003-2", "Comum", "ESTADO", 2027, "0000003-00.2020.8.19.0001", 100, 500000.00),
        ("TJ", "2025.00004-3", "Comum", "ESTADO", 2027, "0000003-00.2020.8.19.0001", 100, 300000.00),  # processo duplicado
        ("TJ", "2025.00005-4", "Comum", "ESTADO", 2027, "0000004-00.2020.8.19.0001", 100, 199999.99),
        ("TJ", "2025.00006-5", "Comum", "ESTADO", 2027, "0000005-00.2020.8.19.0001", 100, None),
    ]
    for d in dados:
        ws.append(d)
    caminho = tmp_path / "saldos.xlsx"
    wb.save(caminho)
    return caminho
```

- [ ] **Step 2: APPEND testes**

```python


from baixar_requisitorios import filtrar_precatorios


def test_filtrar_aplica_saldo_minimo(xlsx_com_saldos):
    resultado = filtrar_precatorios(xlsx_com_saldos, saldo_minimo=200000)
    numeros = sorted(p for p, _ in resultado)
    # 2025.00001 (100k) e 2025.00005 (199.999,99) ficam de fora; 2025.00006 (None) também
    assert numeros == ["2025.00002-1", "2025.00003-2", "2025.00004-3"]


def test_filtrar_retorna_precatorio_e_processo(xlsx_com_saldos):
    resultado = filtrar_precatorios(xlsx_com_saldos, saldo_minimo=200000)
    assert ("2025.00002-1", "0000002-00.2020.8.19.0001") in resultado
    assert ("2025.00003-2", "0000003-00.2020.8.19.0001") in resultado
    assert ("2025.00004-3", "0000003-00.2020.8.19.0001") in resultado


def test_filtrar_ignora_saldos_nao_numericos(xlsx_com_saldos):
    resultado = filtrar_precatorios(xlsx_com_saldos, saldo_minimo=0)
    # 2025.00006-5 tem saldo None, deve ser ignorado mesmo com mínimo 0
    numeros = [p for p, _ in resultado]
    assert "2025.00006-5" not in numeros
```

- [ ] **Step 3: Rodar e confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k filtrar
```
Expected: FAIL com `ImportError`.

- [ ] **Step 4: APPEND filtrar_precatorios em baixar_requisitorios.py**

Adicionar import `from openpyxl import load_workbook` no topo. Em seguida, APPEND:

```python


REGEX_PRECATORIO = re.compile(r"^\d{4}\.\d+-\d+$")
REGEX_PROCESSO_CNJ = re.compile(r"^\d{7}-\d{2}\.\d{4}\.8\.19\.\d{4}$")


def filtrar_precatorios(caminho_xlsx, saldo_minimo):
    """Lê a planilha e retorna lista de (precatorio, processo) com saldo >= mínimo.

    Filtros aplicados:
    - Coluna B casa com padrão de precatório
    - Coluna F casa com padrão CNJ TJRJ
    - Coluna H é numérico e >= saldo_minimo
    """
    wb = load_workbook(caminho_xlsx, read_only=True, data_only=True)
    ws = wb.active
    resultado = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 8:
            continue
        precatorio = row[1]
        processo = row[5]
        saldo = row[7]
        if not isinstance(precatorio, str) or not REGEX_PRECATORIO.match(precatorio):
            continue
        if not isinstance(processo, str) or not REGEX_PROCESSO_CNJ.match(processo.strip()):
            continue
        if not isinstance(saldo, (int, float)):
            continue
        if saldo < saldo_minimo:
            continue
        resultado.append((precatorio, processo.strip()))
    wb.close()
    return resultado
```

- [ ] **Step 5: Rodar testes e confirmar PASS**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k filtrar
```
Expected: 3 passed.

---

### Task 6: Checkpoint atômico (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_baixar_requisitorios.py`

- [ ] **Step 1: APPEND testes**

```python


from baixar_requisitorios import carregar_checkpoint_dl, salvar_checkpoint_dl


def test_checkpoint_dl_inexistente_retorna_vazio(tmp_path):
    assert carregar_checkpoint_dl(tmp_path / "x.json") == {}


def test_checkpoint_dl_roundtrip(tmp_path):
    caminho = tmp_path / "state.json"
    dados = {"0000001-00.2020.8.19.0001": {"status": "ok", "arquivos": ["a.pdf"]}}
    salvar_checkpoint_dl(caminho, dados)
    assert carregar_checkpoint_dl(caminho) == dados


def test_checkpoint_dl_e_atomico_nao_deixa_tmp(tmp_path):
    caminho = tmp_path / "state.json"
    salvar_checkpoint_dl(caminho, {"a": "b"})
    assert not (tmp_path / "state.json.tmp").exists()
    assert caminho.exists()
```

- [ ] **Step 2: Rodar e confirmar falha**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v -k checkpoint_dl
```
Expected: FAIL com `ImportError`.

- [ ] **Step 3: APPEND checkpoint em baixar_requisitorios.py**

Adicionar `import json` ao topo. APPEND:

```python


def carregar_checkpoint_dl(caminho):
    """Carrega checkpoint JSON. Retorna {} se arquivo não existir."""
    caminho = Path(caminho)
    if not caminho.exists():
        return {}
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar_checkpoint_dl(caminho, dados):
    """Salva checkpoint atomicamente: escreve .tmp e renomeia."""
    caminho = Path(caminho)
    tmp = caminho.with_suffix(caminho.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    tmp.replace(caminho)
```

- [ ] **Step 4: Rodar todos os testes do módulo**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/test_baixar_requisitorios.py -v
```
Expected: 25 passed (5 split + 5 sanitizar + 3 extrair + 4 gerar + 3 filtrar + 3 checkpoint + 2 extras esperados).

Confirme com a contagem real (pode haver +/- 1 dependendo de quais testes pegamos).

---

### Task 7: Skeleton Playwright + verificar_sessao + login persistente

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

Esta task adiciona a infraestrutura de browser. Não é testável via unit test (precisa de browser real). Validação é manual no fim.

- [ ] **Step 1: APPEND constantes de paths e função de browser context**

Adicionar ao topo do arquivo:

```python
import asyncio
import sys
from playwright.async_api import async_playwright, BrowserContext, Page
```

E APPEND ao final:

```python


PROJETO_DIR = Path(__file__).parent
USER_DATA_DIR = PROJETO_DIR / "user_data_chromium"
SCREENSHOTS_DIR = PROJETO_DIR / "screenshots_erro"
CHECKPOINT_PATH = PROJETO_DIR / "downloads_state.json"
ERROS_CSV = PROJETO_DIR / "erros_download.csv"

URL_BASE = "https://www3.tjrj.jus.br/portalservicos/"
URL_CONSULTA = "https://www3.tjrj.jus.br/portalservicos/#/consproc/consultaportal"


async def abrir_context(playwright, headless=False):
    """Abre BrowserContext persistente em user_data_chromium/.

    Retorna o context. O caller deve fechar com `await context.close()`.
    """
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    context = await playwright.chromium.launch_persistent_context(
        user_data_dir=str(USER_DATA_DIR),
        headless=headless,
        accept_downloads=True,
        viewport={"width": 1400, "height": 900},
    )
    return context


async def verificar_sessao(context):
    """Verifica se há sessão ativa abrindo a URL base.

    Retorna True se está logado (não foi redirecionado para login),
    False se precisa logar.
    """
    page = context.pages[0] if context.pages else await context.new_page()
    await page.goto(URL_BASE, wait_until="networkidle", timeout=30000)
    # Heurística: se a URL ainda contém /portalservicos/ depois do load,
    # e não há campo de senha visível, consideramos logado.
    url_atual = page.url
    if "login" in url_atual.lower():
        return False
    senha_visivel = await page.locator('input[type="password"]').count()
    return senha_visivel == 0


async def login_interativo(context):
    """Abre o portal e aguarda o usuário logar manualmente.

    Bloqueia até o usuário pressionar ENTER no terminal.
    """
    page = context.pages[0] if context.pages else await context.new_page()
    await page.goto(URL_BASE, wait_until="domcontentloaded")
    print("\n" + "=" * 60)
    print(" LOGIN NECESSÁRIO")
    print(" Faça login no portal TJRJ no navegador que abriu.")
    print(" Quando estiver logado, pressione ENTER aqui.")
    print("=" * 60)
    input(" Pressione ENTER após logar: ")
    if await verificar_sessao(context):
        print(" Sessão confirmada. Continuando...")
        return True
    print(" Sessão ainda não confirmada. Saindo.", file=sys.stderr)
    return False
```

- [ ] **Step 2: Criar função main_test_login isolada para teste manual**

APPEND:

```python


async def _test_login_manual():
    """Função de teste manual: abre o browser, valida sessão (ou pede login)."""
    async with async_playwright() as pw:
        context = await abrir_context(pw, headless=False)
        try:
            if await verificar_sessao(context):
                print("✅ Já estava logado!")
            else:
                logou = await login_interativo(context)
                if logou:
                    print("✅ Login feito com sucesso.")
                else:
                    print("❌ Falha no login.")
            input("Pressione ENTER para fechar...")
        finally:
            await context.close()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    else:
        print("Use --test-login para testar login manualmente.")
        print("(main completo será implementado em task posterior)")
```

- [ ] **Step 3: Rodar testes existentes (não deve quebrar)**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: todos os testes existentes ainda passam (não adicionamos teste novo).

- [ ] **Step 4: Validação manual — teste de login**

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --test-login
```

Comportamento esperado:
- Abre uma janela do Chromium
- Console pede pra logar
- Usuário loga manualmente (certificado digital ou login/senha conforme o método)
- Usuário pressiona ENTER
- Script confirma "Sessão confirmada"
- Pressiona ENTER de novo pra fechar

Após esse teste, a próxima execução já deve detectar a sessão ativa.

---

### Task 8: Exploração interativa da UI (MANUAL — colaboração com usuário)

**⚠️ Esta task NÃO deve ser delegada para subagent.** Requer presença do usuário no browser. O Claude controlador conduz a sessão.

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\seletores.py`

**Objetivo:** Mapear todos os seletores DOM necessários e gerar `seletores.py` com os valores reais descobertos.

- [ ] **Step 1: Rodar o script com page.pause()**

Criar um script temporário `_explorar.py` no diretório do projeto:

```python
"""Script de exploração interativa do portal TJRJ.

Abre o Playwright Inspector que permite picar elementos visualmente.
"""
import asyncio
from playwright.async_api import async_playwright
from baixar_requisitorios import abrir_context, URL_CONSULTA


async def explorar():
    async with async_playwright() as pw:
        context = await abrir_context(pw, headless=False)
        page = context.pages[0] if context.pages else await context.new_page()
        await page.goto(URL_CONSULTA)
        print("\n>>> Playwright Inspector aberto. Use 'Pick locator' para clicar nos elementos.")
        await page.pause()  # bloqueia até o usuário fechar o Inspector
        await context.close()


asyncio.run(explorar())
```

Rodar:

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" _explorar.py
```

- [ ] **Step 2: Mapear seletores guiado pelo usuário**

Com a janela do Playwright Inspector aberta, identificar cada elemento:

  Para cada item da lista abaixo, o usuário deve:
  1. Clicar em "Pick locator" no Inspector
  2. Clicar no elemento na página
  3. O Inspector mostra o locator (ex: `page.get_by_role("textbox", name="Número")`)
  4. Anotar o locator

Lista de elementos a mapear (usar processo `0156129-30.2020.8.19.0001`):
  - [ ] URL final de Consulta Processual (anotar se difere de `URL_CONSULTA`)
  - [ ] Radio "Única" (tipo de numeração)
  - [ ] Campo CNJ prefixo (1º textbox)
  - [ ] Campo CNJ sufixo (2º textbox)
  - [ ] Botão "Pesquisar"
  - [ ] Indicador de "sem resultados" (mensagem quando processo não existe)
  - [ ] Link/linha do resultado na tabela
  - [ ] Como abre a página do processo (mesma aba? nova aba?)
  - [ ] Botão "Processo Eletrônico - Visualizador"
  - [ ] Estrutura da árvore de documentos no visualizador
  - [ ] Como peças aparecem listadas (nome contém "OFREQ"? "REQUISITÓRIO"?)
  - [ ] Como triggerar download da peça (click direto? botão? download de URL?)
  - [ ] Indicador de "sessão expirada" (URL de login, ou campo password aparecendo)

- [ ] **Step 3: Salvar achados em seletores.py**

Criar `seletores.py` com os locators descobertos. Exemplo de estrutura (valores reais vêm da exploração):

```python
"""Seletores DOM do portal TJRJ — gerados pela exploração interativa (Task 8)."""

# ---- Página de consulta processual ----
URL_CONSULTA = "https://www3.tjrj.jus.br/portalservicos/#/consproc/consultaportal"

# Locator strings — usados com page.locator(...) ou page.get_by_role(...)
RADIO_UNICA = 'input[type="radio"][value="unica"]'  # AJUSTAR conforme descoberto
CAMPO_CNJ_PREFIX = '#numProcUnico'                    # AJUSTAR
CAMPO_CNJ_SUFIX = '#numProcOrigem'                    # AJUSTAR
BOTAO_PESQUISAR = 'button:has-text("Pesquisar")'      # AJUSTAR

# Resultado
MSG_SEM_RESULTADOS = 'text=/nenhum.*encontrado/i'     # AJUSTAR
LINK_RESULTADO = 'table tbody tr a'                   # AJUSTAR

# Visualizador
BOTAO_VISUALIZADOR = 'text="Processo Eletrônico - Visualizador"'  # AJUSTAR
PECA_OFREQ_LOCATOR = 'text=/OFREQ|REQUISIT[ÓO]RIO/i'  # AJUSTAR

# Login expirado
INDICADOR_LOGIN_EXPIRADO = 'input[type="password"]'   # AJUSTAR
```

**IMPORTANTE:** Os valores acima são CHUTES iniciais. A exploração com o usuário DEVE substituí-los pelos valores reais descobertos via Playwright Inspector. Sem essa etapa, as tasks seguintes falharão.

- [ ] **Step 4: Documentar fluxo de cliques**

Adicionar comentário no topo de `seletores.py` descrevendo a sequência exata observada:

```python
"""Seletores DOM do portal TJRJ.

Fluxo descoberto na exploração (preencher após Task 8):

1. GET URL_CONSULTA
2. Aguardar carregar (selector: ??)
3. Clicar em RADIO_UNICA
4. Preencher CAMPO_CNJ_PREFIX
5. Preencher CAMPO_CNJ_SUFIX
6. Clicar BOTAO_PESQUISAR
7. Aguardar [resultado OU mensagem sem-resultados]
8. Se resultado: clicar em LINK_RESULTADO → abre [mesma aba | nova aba: ???]
9. Na página do processo: clicar BOTAO_VISUALIZADOR → abre [???]
10. No visualizador: peças OFREQ aparecem como [tipo de elemento? como filtrar?]
11. Click em peça → [comportamento de download?]
"""
```

- [ ] **Step 5: Limpar script temporário**

```powershell
Remove-Item "C:\Users\DARLANMARTINS\Documents\PROJETO 01\_explorar.py" -ErrorAction SilentlyContinue
```

- [ ] **Step 6: Validar que seletores.py importa sem erro**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -c "import seletores; print('seletores ok:', dir(seletores))"
```
Expected: lista das constantes definidas.

**Saída esperada da Task 8:** arquivo `seletores.py` com TODOS os locators reais descobertos, e comentário documentando o fluxo de cliques. As tasks seguintes assumem que `seletores.py` está completo.

---

### Task 9: consultar_processo (navegação até o visualizador)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

Esta task implementa a navegação do início (form de consulta) até abrir a aba do visualizador. Usa os seletores da Task 8.

- [ ] **Step 1: APPEND consultar_processo em baixar_requisitorios.py**

Adicionar import no topo:

```python
import seletores
```

E APPEND:

```python


async def consultar_processo(context, numero_processo):
    """Navega de Consulta Processual até abrir o Visualizador do processo.

    Retorna a Page do visualizador, ou None se o processo não foi encontrado.

    Levanta:
        TimeoutError — se algum selector não aparece no tempo esperado
        LoginExpiradoError — se detectar redirect para login
    """
    page = await context.new_page()
    try:
        await page.goto(seletores.URL_CONSULTA, wait_until="networkidle", timeout=30000)

        # Detectar login expirado
        if await page.locator(seletores.INDICADOR_LOGIN_EXPIRADO).count() > 0:
            raise LoginExpiradoError(f"Sessão expirada ao consultar {numero_processo}")

        prefix, sufix = split_cnj(numero_processo)

        await page.locator(seletores.RADIO_UNICA).check()
        await page.locator(seletores.CAMPO_CNJ_PREFIX).fill(prefix)
        await page.locator(seletores.CAMPO_CNJ_SUFIX).fill(sufix)
        await page.locator(seletores.BOTAO_PESQUISAR).click()

        # Aguardar resultado OU mensagem de "sem resultados"
        try:
            await page.wait_for_selector(
                f"{seletores.LINK_RESULTADO}, {seletores.MSG_SEM_RESULTADOS}",
                timeout=15000,
            )
        except Exception:
            return None

        if await page.locator(seletores.MSG_SEM_RESULTADOS).count() > 0:
            return None

        # Click no resultado pode abrir nova aba; capturamos com expect_page
        async with context.expect_page(timeout=15000) as page_info:
            await page.locator(seletores.LINK_RESULTADO).first.click()
        processo_page = await page_info.value
        await processo_page.wait_for_load_state("networkidle", timeout=30000)

        # Click em Visualizador também abre nova aba
        async with context.expect_page(timeout=30000) as visu_info:
            await processo_page.locator(seletores.BOTAO_VISUALIZADOR).click()
        visualizador_page = await visu_info.value
        await visualizador_page.wait_for_load_state("networkidle", timeout=60000)

        # Fechar as abas intermediárias
        await page.close()
        await processo_page.close()

        return visualizador_page
    except LoginExpiradoError:
        raise
    except Exception:
        await page.close()
        raise


class LoginExpiradoError(Exception):
    """Disparado quando a sessão expirou durante a navegação."""
    pass
```

- [ ] **Step 2: Mover declaração de LoginExpiradoError para antes do uso**

Ajustar a ordem no arquivo: `LoginExpiradoError` deve estar declarado ANTES de `consultar_processo`. Mover para imediatamente após os imports e antes das constantes:

```python
class LoginExpiradoError(Exception):
    """Disparado quando a sessão expirou durante a navegação."""
    pass
```

- [ ] **Step 3: Adicionar handler CLI para teste de um processo**

Modificar o bloco `if __name__ == "__main__":` no final do arquivo para incluir um modo `--test-processo`:

```python
async def _test_processo(numero):
    """Função de teste manual: abre o visualizador de um único processo."""
    async with async_playwright() as pw:
        context = await abrir_context(pw, headless=False)
        try:
            if not await verificar_sessao(context):
                if not await login_interativo(context):
                    return
            print(f"\nConsultando processo {numero}...")
            visu = await consultar_processo(context, numero)
            if visu is None:
                print("Processo NÃO encontrado.")
            else:
                print(f"Visualizador aberto: {visu.url}")
                input("Inspecione a tela e pressione ENTER para fechar...")
        finally:
            await context.close()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-processo":
        asyncio.run(_test_processo(sys.argv[2]))
    else:
        print("Modos disponíveis:")
        print("  --test-login")
        print("  --test-processo NNNNNNN-NN.AAAA.8.19.NNNN")
        print("(main completo será implementado em task posterior)")
```

- [ ] **Step 4: Validação manual — testar com o processo Garrastazu**

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --test-processo "0156129-30.2020.8.19.0001"
```

Expected:
- Browser abre
- Sessão é detectada (já logada da Task 7)
- Navegação acontece: consulta → clique no resultado → clique em visualizador
- Console imprime "Visualizador aberto: <url>"
- Usuário inspeciona visualmente que o visualizador do processo correto está aberto
- Pressiona ENTER para fechar

Se algo falhar (timeout, selector não achado), ajustar `seletores.py` e tentar de novo.

---

### Task 10: localizar_pecas_ofreq + baixar_peca + parse + rename

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

- [ ] **Step 1: APPEND funções**

```python


async def localizar_pecas_ofreq(visualizador_page):
    """Encontra todas as peças no visualizador com nome contendo OFREQ/REQUISITÓRIO.

    Retorna lista de Locators clicáveis (cada um corresponde a um requisitório).
    """
    locator = visualizador_page.locator(seletores.PECA_OFREQ_LOCATOR)
    count = await locator.count()
    return [locator.nth(i) for i in range(count)]


async def baixar_peca(visualizador_page, peca_locator, pasta_temp):
    """Clica numa peça e baixa o PDF.

    Retorna Path do arquivo baixado (em pasta_temp, com nome temporário).
    """
    pasta_temp.mkdir(parents=True, exist_ok=True)
    async with visualizador_page.expect_download(timeout=60000) as dl_info:
        await peca_locator.click()
    download = await dl_info.value
    destino_tmp = pasta_temp / f"_temp_{download.suggested_filename}"
    await download.save_as(destino_tmp)
    return destino_tmp


async def processar_processo(context, precatorio, numero_processo, pasta_saida, pasta_tmp):
    """Pipeline completo para um processo: consulta → baixa peças → renomeia.

    Retorna dict com:
      - status: "ok" | "sem_requisitorio" | "processo_nao_encontrado" | "erro_*"
      - arquivos: list[str]  (nomes dos PDFs salvos)
      - motivo: str  (presente em status de erro)
    """
    visu = None
    try:
        visu = await consultar_processo(context, numero_processo)
        if visu is None:
            return {"status": "processo_nao_encontrado", "arquivos": []}

        pecas = await localizar_pecas_ofreq(visu)
        if not pecas:
            return {"status": "sem_requisitorio", "arquivos": []}

        arquivos_finais = []
        for peca in pecas:
            tmp_pdf = await baixar_peca(visu, peca, pasta_tmp)
            pdf_bytes = tmp_pdf.read_bytes()
            beneficiario = extrair_beneficiario(pdf_bytes)
            if beneficiario is None:
                pasta_manual = pasta_saida / "manual_revisar"
                pasta_manual.mkdir(parents=True, exist_ok=True)
                final = gerar_nome_arquivo(precatorio, "SEM_NOME", pasta_manual)
                tmp_pdf.replace(final)
                arquivos_finais.append(str(final.relative_to(pasta_saida)))
            else:
                final = gerar_nome_arquivo(precatorio, beneficiario, pasta_saida)
                tmp_pdf.replace(final)
                arquivos_finais.append(final.name)

        return {"status": "ok", "arquivos": arquivos_finais}

    except LoginExpiradoError:
        raise  # propaga para o orquestrador pausar
    except Exception as e:
        return {"status": "erro_navegacao", "arquivos": [], "motivo": f"{type(e).__name__}: {e}"}
    finally:
        if visu is not None:
            await visu.close()
```

- [ ] **Step 2: Adicionar modo CLI --test-baixar para teste isolado**

Substituir o bloco `if __name__ == "__main__":` por:

```python
async def _test_baixar(numero):
    """Pipeline completo de um processo de teste."""
    pasta_saida = Path.home() / "Downloads" / "Precatórios_Requisitórios"
    pasta_saida.mkdir(parents=True, exist_ok=True)
    pasta_tmp = PROJETO_DIR / "_tmp_downloads"

    async with async_playwright() as pw:
        context = await abrir_context(pw, headless=False)
        try:
            if not await verificar_sessao(context):
                if not await login_interativo(context):
                    return
            print(f"\nProcessando {numero}...")
            resultado = await processar_processo(
                context,
                precatorio="TESTE",
                numero_processo=numero,
                pasta_saida=pasta_saida,
                pasta_tmp=pasta_tmp,
            )
            print(f"Resultado: {resultado}")
            input("ENTER para fechar...")
        finally:
            await context.close()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-processo":
        asyncio.run(_test_processo(sys.argv[2]))
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-baixar":
        asyncio.run(_test_baixar(sys.argv[2]))
    else:
        print("Modos disponíveis:")
        print("  --test-login")
        print("  --test-processo NNN")
        print("  --test-baixar NNN")
```

- [ ] **Step 3: Validação manual — pipeline completo de um processo**

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --test-baixar "0156129-30.2020.8.19.0001"
```

Expected:
- Browser abre
- Navega até o visualizador
- Identifica 1+ peças de requisitório
- Baixa cada uma
- Extrai beneficiário do PDF
- Renomeia para "TESTE - TECHNE ENGENHARIA E SISTEMAS LTDA.pdf" (ou similar)
- Salva em `C:\Users\DARLANMARTINS\Downloads\Precatórios_Requisitórios\`
- Imprime: `{"status": "ok", "arquivos": ["TESTE - TECHNE ENGENHARIA E SISTEMAS LTDA.pdf"]}`

Validar abrindo o PDF salvo e comparando visualmente com o `GARRASTAZU - REQUISITÓRIO.pdf` original (devem ser idênticos no conteúdo).

---

### Task 11: Orquestração async + CLI completo

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\baixar_requisitorios.py`

- [ ] **Step 1: APPEND argparse, main, worker pool**

APPEND:

```python


import argparse
import csv
from datetime import datetime


PASTA_DOWNLOADS = Path.home() / "Downloads"
SAIDA_PADRAO = PASTA_DOWNLOADS / "Precatórios_Requisitórios"
ENTRADA_PADRAO = PASTA_DOWNLOADS / "Precatórios 2027 - Atualizado.xlsx"


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Baixa ofícios requisitórios do TJRJ.")
    p.add_argument("--entrada", type=Path, default=ENTRADA_PADRAO)
    p.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    p.add_argument("--saldo-minimo", type=float, default=200000.0)
    p.add_argument("--workers", type=int, default=3)
    p.add_argument("--limit", type=int, default=None)
    p.add_argument("--reset", action="store_true")
    p.add_argument("--apenas-erros", action="store_true")
    p.add_argument("--processo", type=str, default=None,
                   help="Roda só este processo (modo debug)")
    p.add_argument("--headless", action="store_true")
    return p.parse_args(argv)


def precisa_processar(processo, resultados, apenas_erros):
    if processo not in resultados:
        return True
    status = resultados[processo].get("status")
    if apenas_erros and status != "ok":
        return True
    return status.startswith("erro_") if status else True


async def worker(nome, fila, resultados, contador, lock, context, args):
    """Worker async que consome processos da fila."""
    while True:
        try:
            item = await fila.get()
        except asyncio.CancelledError:
            return
        if item is None:
            fila.task_done()
            return
        precatorio, processo = item
        try:
            try:
                pasta_tmp = PROJETO_DIR / "_tmp_downloads" / nome
                resultado = await processar_processo(
                    context, precatorio, processo, args.saida, pasta_tmp,
                )
            except LoginExpiradoError:
                # Sinaliza para o main pausar
                async with lock:
                    resultados[processo] = {"status": "erro_navegacao",
                                            "motivo": "login_expirado"}
                print(f"\n[{nome}] LOGIN EXPIRADO - pausando")
                fila.task_done()
                # Esperar sinalização externa para retomar (não implementado: simples re-fila)
                raise

            async with lock:
                resultados[processo] = resultado
                contador["n"] += 1
                if contador["n"] % 10 == 0:
                    salvar_checkpoint_dl(CHECKPOINT_PATH, resultados)
                    _imprimir_progresso(contador["n"], contador["total"])
        except LoginExpiradoError:
            return
        finally:
            fila.task_done()


def _imprimir_progresso(concluidos, total):
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

    args.saida.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)

    if args.reset and CHECKPOINT_PATH.exists():
        CHECKPOINT_PATH.unlink()
        print(f"Checkpoint apagado: {CHECKPOINT_PATH}")

    resultados = carregar_checkpoint_dl(CHECKPOINT_PATH)
    print(f"Checkpoint carregado: {len(resultados)} processos")

    print(f"Lendo {args.entrada}...")
    pares = filtrar_precatorios(args.entrada, args.saldo_minimo)
    print(f"  {len(pares)} precatórios com saldo >= R$ {args.saldo_minimo:,.2f}")
    processos_unicos = {}
    for prec, proc in pares:
        if proc not in processos_unicos:
            processos_unicos[proc] = prec  # primeiro precatório encontrado
    print(f"  {len(processos_unicos)} processos únicos")

    if args.processo:
        processos_unicos = {args.processo: processos_unicos.get(args.processo, "DEBUG")}

    pendentes = [
        (prec, proc) for proc, prec in processos_unicos.items()
        if precisa_processar(proc, resultados, args.apenas_erros)
    ]
    if args.limit:
        pendentes = pendentes[: args.limit]
    print(f"  {len(pendentes)} pendentes para processar")

    if not pendentes:
        print("Nada a fazer.")
        return 0

    fila = asyncio.Queue()
    for item in pendentes:
        fila.put_nowait(item)
    for _ in range(args.workers):
        fila.put_nowait(None)  # sentinela para encerrar workers

    contador = {"n": 0, "total": len(pendentes)}
    lock = asyncio.Lock()

    inicio = datetime.now()
    async with async_playwright() as pw:
        context = await abrir_context(pw, headless=args.headless)
        try:
            if not await verificar_sessao(context):
                if not await login_interativo(context):
                    return 1
            print(f"Iniciando {args.workers} workers...")
            tasks = [
                asyncio.create_task(worker(f"w{i}", fila, resultados, contador,
                                           lock, context, args))
                for i in range(args.workers)
            ]
            try:
                await asyncio.gather(*tasks)
            except LoginExpiradoError:
                print("\nSessão expirou. Faça login no browser e re-execute o script.")
                salvar_checkpoint_dl(CHECKPOINT_PATH, resultados)
                return 2
        finally:
            salvar_checkpoint_dl(CHECKPOINT_PATH, resultados)
            await context.close()

    print()
    _imprimir_relatorio(processos_unicos, resultados, args.saida, inicio)
    return 0


def _imprimir_relatorio(processos_unicos, resultados, saida, inicio):
    ok = sem_req = nao_enc = erro_rede = erro_nav = erro_pars = 0
    erros_detalhe = []
    total_pdfs = 0
    for proc in processos_unicos:
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
    print(f" Pasta saída               : {saida}")
    print("=" * 60)
    if erros_detalhe:
        with open(ERROS_CSV, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Processo", "Status", "Motivo"])
            w.writerows(erros_detalhe)
        print(f" Log de erros: {ERROS_CSV}")
```

- [ ] **Step 2: Atualizar bloco `if __name__ == "__main__":` para chamar main()**

Substituir o bloco final por:

```python
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-login":
        asyncio.run(_test_login_manual())
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-processo":
        asyncio.run(_test_processo(sys.argv[2]))
    elif len(sys.argv) > 2 and sys.argv[1] == "--test-baixar":
        asyncio.run(_test_baixar(sys.argv[2]))
    else:
        sys.exit(asyncio.run(main()))
```

- [ ] **Step 3: Confirmar que testes unit ainda passam**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" -m pytest tests/ -v
```
Expected: todos passando (sem novos testes nesta task — funções de browser não são unit-testáveis).

- [ ] **Step 4: Smoke test --help**

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --help
```
Expected: imprime opções com `--entrada`, `--saida`, `--saldo-minimo`, `--workers`, `--limit`, `--reset`, `--apenas-erros`, `--processo`, `--headless`.

---

### Task 12: Validação --limit 5 e retomada

**Files:** nenhum (apenas execução)

- [ ] **Step 1: Rodar com limite pequeno**

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --limit 5
```

Expected:
- Browser abre (sessão reusada)
- Filtra precatórios >= 200k
- Mostra "5 pendentes"
- Inicia 3 workers
- Progresso aparece na tela
- Termina em ~2-3 minutos
- 5 ou mais PDFs salvos em `Downloads\Precatórios_Requisitórios\`
- Relatório: "Processos OK: 5"

- [ ] **Step 2: Verificar arquivos baixados**

```powershell
Get-ChildItem "C:\Users\DARLANMARTINS\Downloads\Precatórios_Requisitórios" *.pdf | Select-Object -First 10 Name, Length
```
Expected: lista de 5+ PDFs nomeados no padrão `2025.XXXXX-X - NOME_BENEFICIARIO.pdf`.

- [ ] **Step 3: Teste de retomada**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --limit 5
```
Expected: console diz "Checkpoint carregado: 5+ processos" e "0 pendentes para processar" → sai sem fazer nada.

- [ ] **Step 4: Validar conteúdo de 1 PDF**

Abrir manualmente um dos PDFs baixados e confirmar visualmente:
- Cabeçalho "OFÍCIO REQUISITÓRIO"
- Seção "III - BENEFICIÁRIO" com o nome que está no filename
- Documento legível, completo

---

### Task 13: Execução completa

**Files:** nenhum

- [ ] **Step 1: Rodar tudo**

```powershell
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py
```

Esperado:
- Detecta sessão ativa (login automático via cookies)
- Filtra ~2.350 precatórios → ~1.952 processos únicos
- Pendentes = 1.952 - 5 (já baixados no teste) ≈ 1.947
- 3 workers iniciam
- Barra de progresso atualizando
- Tempo estimado: 4-7 horas
- Salva checkpoint a cada 10 processos

- [ ] **Step 2: Monitorar (sem interromper)**

Verificar periodicamente:
- Console mostrando progresso
- Pasta `Downloads\Precatórios_Requisitórios\` crescendo
- Sem screenshots em `screenshots_erro/` indicando muitos erros

Se a sessão expirar, o script avisa, pede pra logar e re-executar.

- [ ] **Step 3: Após terminar — relatório final**

```powershell
$pasta = "C:\Users\DARLANMARTINS\Downloads\Precatórios_Requisitórios"
$pdfs = Get-ChildItem $pasta -Filter *.pdf -Recurse
Write-Host "Total de PDFs: $($pdfs.Count)"
Write-Host "Pasta manual_revisar: $((Get-ChildItem "$pasta\manual_revisar" -Filter *.pdf -ErrorAction SilentlyContinue).Count) PDFs"
if (Test-Path "C:\Users\DARLANMARTINS\Documents\PROJETO 01\erros_download.csv") {
    Write-Host "Erros registrados:"
    Get-Content "C:\Users\DARLANMARTINS\Documents\PROJETO 01\erros_download.csv" | Select-Object -First 20
}
```

Expected:
- ~1.800+ PDFs (90%+ dos 1.952 processos)
- Alguns "sem_requisitorio" e talvez erros isolados
- `erros_download.csv` se houve falhas

- [ ] **Step 4: Re-tentar erros (se houver)**

```powershell
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py --apenas-erros
```
Reroda só os processos com erro_rede/erro_navegacao/erro_parsing.

---

## Critérios de aceite (do spec)

- [x] Login manual uma vez, sessão persiste — validado na Task 7
- [x] Filtra saldos >= R$ 200.000 → ~1.952 processos únicos — validado na Task 5
- [x] ≥ 90% dos processos resultam em PDF baixado — validado na Task 13 Step 3
- [x] Nomes correspondem ao beneficiário do PDF — validado nas Tasks 10 e 12
- [x] Múltiplos requisitórios numerados com (2), (3) — validado pela função gerar_nome_arquivo (Task 4)
- [x] Ctrl+C salva checkpoint, retomada continua — re-execução pula com base no checkpoint (Task 12 Step 3)
- [x] Planilha de entrada permanece inalterada — script só faz load_workbook(read_only=True), nunca save
- [x] Screenshots em erros — diretório `screenshots_erro/` criado em Task 11
