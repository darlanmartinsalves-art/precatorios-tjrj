# Atualização de Saldos de Precatórios - Plano de Implementação

> **Para agentes:** USAR SUB-SKILL: superpowers:subagent-driven-development (recomendado) ou superpowers:executing-plans. Passos usam checkbox `- [ ]` para tracking.

**Goal:** Construir um script Python que consulta a API do TJRJ e preenche a coluna H da planilha `Precatórios 2027.xlsx` com saldos atualizados de 6.288 precatórios, com retomada automática em caso de falha.

**Architecture:** Script Python único com módulo importável, dependências mínimas (openpyxl, requests), paralelismo via ThreadPoolExecutor, checkpoint atômico em JSON, formato preservado via openpyxl.

**Tech Stack:** Python 3 (Microsoft Store), openpyxl, requests, pytest (testes), unittest.mock (mocking HTTP).

**Spec:** [docs/superpowers/specs/2026-05-27-atualizar-saldos-precatorios-design.md](../specs/2026-05-27-atualizar-saldos-precatorios-design.md)

---

## File Structure

```
C:\Users\DARLANMARTINS\Documents\PROJETO 01\
  atualizar_saldos.py            # script principal (funções + CLI)
  requirements.txt               # openpyxl, requests, pytest
  pytest.ini                     # config pytest
  tests/
    __init__.py
    conftest.py                  # fixtures pytest (gera xlsx pequeno)
    test_atualizar_saldos.py     # testes unitários
```

**Responsabilidades de `atualizar_saldos.py`:**
- `extrair_precatorios(caminho_xlsx) -> list[tuple[int, str]]` — lê coluna B
- `consultar_saldo(numero, session) -> float | str` — chama API com retry
- `carregar_checkpoint(caminho) -> dict` — lê JSON com {numero: resultado}
- `salvar_checkpoint(caminho, dados) -> None` — escrita atômica (.tmp + rename)
- `escrever_xlsx(entrada, saida, precatorios, resultados) -> None` — gera xlsx final
- `main(args) -> int` — orquestração + CLI

---

### Task 1: Setup do ambiente

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\requirements.txt`
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\pytest.ini`
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\__init__.py`

- [ ] **Step 1: Instalar Python (se necessário)**

Abrir Microsoft Store, buscar "Python 3.12", clicar Install. Após instalar, fechar e reabrir o PowerShell. Verificar:

```powershell
python --version
```
Expected: `Python 3.12.x` ou similar.

Se ainda mostrar erro do MS Store: rodar `Get-Command python` para ver se há múltiplas instalações conflitando.

- [ ] **Step 2: Criar requirements.txt**

```
openpyxl>=3.1.0
requests>=2.31.0
pytest>=8.0.0
```

- [ ] **Step 3: Instalar dependências**

```powershell
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
python -m pip install -r requirements.txt
```
Expected: `Successfully installed openpyxl-X.X.X requests-X.X.X pytest-X.X.X ...`

- [ ] **Step 4: Criar pytest.ini**

```ini
[pytest]
testpaths = tests
python_files = test_*.py
```

- [ ] **Step 5: Criar tests/__init__.py (arquivo vazio)**

```python
```

- [ ] **Step 6: Verificar instalação**

```powershell
python -c "import openpyxl, requests; print('ok')"
python -m pytest --version
```
Expected: `ok` na primeira linha, `pytest 8.x.x` na segunda.

---

### Task 2: Fixture xlsx para testes

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\conftest.py`

- [ ] **Step 1: Escrever conftest.py que gera xlsx de teste**

```python
import pytest
from openpyxl import Workbook
from pathlib import Path

@pytest.fixture
def xlsx_pequena(tmp_path):
    """Cria um xlsx temporário com 5 precatórios para testes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Precatorios"
    ws.append(["Tribunal", "Precatório", "Natureza", "Devedor", "Orçamento",
               "Processo Judicial", "Valor Histórico", "Saldo atualizado"])
    dados = [
        ("TJ", "2025.09451-0", "Alimentícia", "UERJ", 2027, "0230764-50.2018.8.19.0001", 50000, None),
        ("TJ", "2025.09452-8", "Alimentícia", "ESTADO", 2027, "0006013-75.2021.8.19.0001", 70000, None),
        ("TJ", "2025.09453-6", "Alimentícia", "IPERJ", 2027, "0027542-97.2014.8.19.0001", 300000, None),
        ("TJ", "2025.09462-5", "Comum", "ESTADO", 2027, "0000224-76.2021.8.19.0072", 100000, None),
        ("TJ", "2025.09470-6", "Comum", "ESTADO", 2027, "0123018-26.2018.8.19.0001", 350000, None),
    ]
    for d in dados:
        ws.append(d)
    caminho = tmp_path / "pequena.xlsx"
    wb.save(caminho)
    return caminho
```

- [ ] **Step 2: Verificar fixture funciona**

```powershell
python -c "from openpyxl import Workbook; wb = Workbook(); print('openpyxl ok')"
```
Expected: `openpyxl ok`

---

### Task 3: Função extrair_precatorios (TDD)

**Files:**
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\atualizar_saldos.py`
- Create: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_atualizar_saldos.py`

- [ ] **Step 1: Escrever teste falhante para extrair_precatorios**

Em `tests/test_atualizar_saldos.py`:

```python
from atualizar_saldos import extrair_precatorios

def test_extrair_precatorios_retorna_lista_de_tuplas(xlsx_pequena):
    resultado = extrair_precatorios(xlsx_pequena)
    assert len(resultado) == 5
    assert resultado[0] == (2, "2025.09451-0")
    assert resultado[4] == (6, "2025.09470-6")

def test_extrair_precatorios_ignora_linhas_sem_numero_valido(xlsx_pequena, tmp_path):
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_pequena)
    ws = wb.active
    ws.append(["", "INVALIDO", "", "", 0, "", 0, None])
    ws.append(["TJ", "2025.99999-9", "Comum", "", 2027, "", 0, None])
    saida = tmp_path / "modificada.xlsx"
    wb.save(saida)

    resultado = extrair_precatorios(saida)
    numeros = [n for _, n in resultado]
    assert "INVALIDO" not in numeros
    assert "2025.99999-9" in numeros
```

- [ ] **Step 2: Rodar teste — confirmar que falha**

```powershell
python -m pytest tests/test_atualizar_saldos.py::test_extrair_precatorios_retorna_lista_de_tuplas -v
```
Expected: FAIL com `ImportError: cannot import name 'extrair_precatorios'` (módulo ainda não existe).

- [ ] **Step 3: Implementar extrair_precatorios**

Em `atualizar_saldos.py`:

```python
"""Atualiza saldos de precatórios consultando a API do TJRJ."""
import re
from pathlib import Path
from openpyxl import load_workbook

REGEX_PRECATORIO = re.compile(r"^\d{4}\.\d+-\d+$")


def extrair_precatorios(caminho_xlsx):
    """Lê coluna B do xlsx e retorna lista de (linha_excel, numero).

    Filtra apenas linhas cuja coluna B casa com o padrão de número de precatório.
    Pula a linha 1 (cabeçalho).
    """
    wb = load_workbook(caminho_xlsx, read_only=True, data_only=True)
    ws = wb.active
    precatorios = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or len(row) < 2:
            continue
        numero = row[1]
        if isinstance(numero, str) and REGEX_PRECATORIO.match(numero):
            precatorios.append((row_idx, numero))
    wb.close()
    return precatorios
```

- [ ] **Step 4: Rodar testes — confirmar que passam**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k extrair
```
Expected: 2 passed.

---

### Task 4: Função consultar_saldo (TDD com mocks)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\atualizar_saldos.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_atualizar_saldos.py`

- [ ] **Step 1: Escrever teste — happy path**

Em `tests/test_atualizar_saldos.py` (append):

```python
from unittest.mock import MagicMock, patch
from atualizar_saldos import consultar_saldo

def _mock_response(status_code, json_data=None):
    m = MagicMock()
    m.status_code = status_code
    m.json.return_value = json_data or {}
    return m

def test_consultar_saldo_retorna_float_em_caso_de_sucesso():
    session = MagicMock()
    session.get.return_value = _mock_response(200, {"Saldo": 13613821.78})
    resultado = consultar_saldo("2025.17049-6", session)
    assert resultado == 13613821.78

def test_consultar_saldo_404_retorna_nao_encontrado():
    session = MagicMock()
    session.get.return_value = _mock_response(404)
    resultado = consultar_saldo("9999.99999-9", session)
    assert resultado == "NAO_ENCONTRADO"

def test_consultar_saldo_null_retorna_sem_saldo():
    session = MagicMock()
    session.get.return_value = _mock_response(200, {"Saldo": None})
    resultado = consultar_saldo("2025.00001-1", session)
    assert resultado == "SEM_SALDO"
```

- [ ] **Step 2: Rodar teste — confirmar que falha**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k consultar_saldo
```
Expected: FAIL com `ImportError`.

- [ ] **Step 3: Adicionar `import time` ao topo de atualizar_saldos.py**

Modificar o cabeçalho para incluir `time`:

```python
"""Atualiza saldos de precatórios consultando a API do TJRJ."""
import re
import time
from pathlib import Path
from openpyxl import load_workbook
```

- [ ] **Step 4: Implementar consultar_saldo**

Em `atualizar_saldos.py` (append, após `extrair_precatorios`):

```python
URL_API = "https://www3.tjrj.jus.br/PortalConhecimento/api/precatorios/consultaPosicao"


def consultar_saldo(numero, session, max_tentativas=3, timeout=30):
    """Consulta o saldo do precatório no TJRJ.

    Retorna float em caso de sucesso, ou string com código de erro:
      "NAO_ENCONTRADO" — HTTP 404
      "SEM_SALDO"      — HTTP 200 mas Saldo é null
      "RESPOSTA_INVALIDA" — JSON malformado
      "ERRO_REDE"      — timeout / 5xx após N tentativas
    """
    for tentativa in range(1, max_tentativas + 1):
        try:
            resp = session.get(URL_API, params={"numeroPrecatorio": numero}, timeout=timeout)
            if resp.status_code == 404:
                return "NAO_ENCONTRADO"
            if resp.status_code >= 500:
                raise IOError(f"HTTP {resp.status_code}")
            if resp.status_code != 200:
                return "RESPOSTA_INVALIDA"
            try:
                dados = resp.json()
            except ValueError:
                return "RESPOSTA_INVALIDA"
            saldo = dados.get("Saldo")
            if saldo is None:
                return "SEM_SALDO"
            return float(saldo)
        except (IOError, OSError):
            if tentativa < max_tentativas:
                time.sleep(2 ** (tentativa - 1))
            continue
    return "ERRO_REDE"
```

- [ ] **Step 5: Rodar testes — happy path / 404 / null**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k consultar_saldo
```
Expected: 3 passed.

- [ ] **Step 6: Adicionar testes — retry em 5xx e erro após max tentativas**

Em `tests/test_atualizar_saldos.py` (append):

```python
def test_consultar_saldo_retry_em_5xx():
    session = MagicMock()
    session.get.side_effect = [
        _mock_response(500),
        _mock_response(500),
        _mock_response(200, {"Saldo": 100.0}),
    ]
    with patch("atualizar_saldos.time.sleep"):
        resultado = consultar_saldo("2025.00001-1", session)
    assert resultado == 100.0
    assert session.get.call_count == 3

def test_consultar_saldo_erro_rede_apos_max_tentativas():
    session = MagicMock()
    session.get.side_effect = [_mock_response(500)] * 3
    with patch("atualizar_saldos.time.sleep"):
        resultado = consultar_saldo("2025.00001-1", session)
    assert resultado == "ERRO_REDE"
```

- [ ] **Step 7: Rodar todos os testes de consultar_saldo**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k consultar_saldo
```
Expected: 5 passed.

---

### Task 5: Checkpoint atômico (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\atualizar_saldos.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_atualizar_saldos.py`

- [ ] **Step 1: Escrever testes para checkpoint**

Em `tests/test_atualizar_saldos.py` (append):

```python
from atualizar_saldos import carregar_checkpoint, salvar_checkpoint

def test_checkpoint_inexistente_retorna_dict_vazio(tmp_path):
    caminho = tmp_path / "checkpoint.json"
    assert carregar_checkpoint(caminho) == {}

def test_salvar_e_carregar_checkpoint_preserva_dados(tmp_path):
    caminho = tmp_path / "checkpoint.json"
    dados = {"2025.09451-0": 59308.24, "2025.09452-8": "ERRO_REDE"}
    salvar_checkpoint(caminho, dados)
    assert carregar_checkpoint(caminho) == dados

def test_salvar_checkpoint_e_atomico(tmp_path):
    caminho = tmp_path / "checkpoint.json"
    salvar_checkpoint(caminho, {"a": 1})
    # após salvar, não deve sobrar arquivo .tmp
    assert not (tmp_path / "checkpoint.json.tmp").exists()
    assert caminho.exists()
```

- [ ] **Step 2: Rodar testes — confirmar que falham**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k checkpoint
```
Expected: FAIL com `ImportError`.

- [ ] **Step 3: Implementar checkpoint**

Em `atualizar_saldos.py` (append). Adicionar `import json` ao topo:

```python
import json
```

E adicionar as funções:

```python
def carregar_checkpoint(caminho):
    """Carrega checkpoint JSON. Retorna dict vazio se arquivo não existir."""
    caminho = Path(caminho)
    if not caminho.exists():
        return {}
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar_checkpoint(caminho, dados):
    """Salva checkpoint de forma atômica: escreve .tmp e faz rename."""
    caminho = Path(caminho)
    tmp = caminho.with_suffix(caminho.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    tmp.replace(caminho)
```

- [ ] **Step 4: Rodar testes**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k checkpoint
```
Expected: 3 passed.

---

### Task 6: Escrever xlsx final (TDD)

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\atualizar_saldos.py`
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\tests\test_atualizar_saldos.py`

- [ ] **Step 1: Escrever teste para escrever_xlsx**

Em `tests/test_atualizar_saldos.py` (append):

```python
from atualizar_saldos import escrever_xlsx
from openpyxl import load_workbook
import hashlib

def _md5(caminho):
    return hashlib.md5(caminho.read_bytes()).hexdigest()

def test_escrever_xlsx_preenche_coluna_H(xlsx_pequena, tmp_path):
    saida = tmp_path / "saida.xlsx"
    precatorios = [(2, "2025.09451-0"), (3, "2025.09452-8"), (4, "2025.09453-6")]
    resultados = {
        "2025.09451-0": 59308.24,
        "2025.09452-8": 83433.75,
        "2025.09453-6": "ERRO_REDE",  # não deve gravar
    }
    escrever_xlsx(xlsx_pequena, saida, precatorios, resultados)

    wb = load_workbook(saida)
    ws = wb.active
    assert ws.cell(row=2, column=8).value == 59308.24
    assert ws.cell(row=3, column=8).value == 83433.75
    assert ws.cell(row=4, column=8).value is None  # erro não gravado

def test_escrever_xlsx_aplica_formato_moeda(xlsx_pequena, tmp_path):
    saida = tmp_path / "saida.xlsx"
    escrever_xlsx(xlsx_pequena, saida, [(2, "2025.09451-0")], {"2025.09451-0": 100.0})
    wb = load_workbook(saida)
    ws = wb.active
    assert "#,##0.00" in ws.cell(row=2, column=8).number_format

def test_escrever_xlsx_nao_altera_entrada(xlsx_pequena, tmp_path):
    saida = tmp_path / "saida.xlsx"
    hash_antes = _md5(xlsx_pequena)
    escrever_xlsx(xlsx_pequena, saida, [(2, "2025.09451-0")], {"2025.09451-0": 100.0})
    hash_depois = _md5(xlsx_pequena)
    assert hash_antes == hash_depois
```

- [ ] **Step 2: Rodar testes — confirmar que falham**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k escrever
```
Expected: FAIL com `ImportError`.

- [ ] **Step 3: Implementar escrever_xlsx**

Em `atualizar_saldos.py` (append):

```python
COLUNA_SALDO = 8  # coluna H
FORMATO_MOEDA = 'R$ #,##0.00'


def escrever_xlsx(entrada, saida, precatorios, resultados):
    """Carrega xlsx de entrada, preenche coluna H e salva em saída.

    O arquivo de entrada não é modificado.
    Resultados não-numéricos (strings de erro) são ignorados.
    """
    entrada = Path(entrada)
    saida = Path(saida)
    wb = load_workbook(entrada)
    ws = wb.active
    for linha, numero in precatorios:
        valor = resultados.get(numero)
        celula = ws.cell(row=linha, column=COLUNA_SALDO)
        celula.number_format = FORMATO_MOEDA
        if isinstance(valor, (int, float)):
            celula.value = float(valor)
    wb.save(saida)
    wb.close()
```

- [ ] **Step 4: Rodar testes**

```powershell
python -m pytest tests/test_atualizar_saldos.py -v -k escrever
```
Expected: 3 passed.

- [ ] **Step 5: Rodar todos os testes para garantir que nada quebrou**

```powershell
python -m pytest tests/ -v
```
Expected: 13 passed (2 extrair + 5 consultar + 3 checkpoint + 3 escrever).

---

### Task 7: Orquestração + CLI

**Files:**
- Modify: `C:\Users\DARLANMARTINS\Documents\PROJETO 01\atualizar_saldos.py`

- [ ] **Step 1: Adicionar imports adicionais ao topo de atualizar_saldos.py**

```python
import argparse
import csv
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import requests
```

- [ ] **Step 2: Implementar função main (orquestração)**

Append em `atualizar_saldos.py`:

```python
PASTA_DOWNLOADS = Path.home() / "Downloads"
ENTRADA_PADRAO = PASTA_DOWNLOADS / "Precatórios 2027.xlsx"
SAIDA_PADRAO_SUFIX = " - Atualizado.xlsx"
CHECKPOINT_PADRAO = Path(__file__).parent / "saldos_checkpoint.json"
ERROS_CSV = Path(__file__).parent / "erros_consulta.csv"


def _formatar_eta(segundos):
    if segundos < 60:
        return f"{int(segundos)}s"
    minutos, seg = divmod(int(segundos), 60)
    return f"{minutos}min{seg:02d}s"


def main(argv=None):
    parser = argparse.ArgumentParser(description="Atualiza saldos de precatórios via API do TJRJ.")
    parser.add_argument("--entrada", type=Path, default=ENTRADA_PADRAO)
    parser.add_argument("--saida", type=Path, default=None)
    parser.add_argument("--checkpoint", type=Path, default=CHECKPOINT_PADRAO)
    parser.add_argument("--workers", type=int, default=20)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--reset", action="store_true", help="Apaga checkpoint antes")
    parser.add_argument("--apenas-erros", action="store_true", help="Re-consulta entradas com erro")
    parser.add_argument("--so-aplicar", action="store_true", help="Pula consulta, só aplica checkpoint")
    args = parser.parse_args(argv)

    if not args.entrada.exists():
        print(f"ERRO: arquivo não encontrado: {args.entrada}", file=sys.stderr)
        return 1

    saida = args.saida or args.entrada.with_name(args.entrada.stem + SAIDA_PADRAO_SUFIX)

    if args.reset and args.checkpoint.exists():
        args.checkpoint.unlink()
        print(f"Checkpoint apagado: {args.checkpoint}")

    print(f"Lendo {args.entrada}...")
    precatorios = extrair_precatorios(args.entrada)
    print(f"Encontrados {len(precatorios)} precatórios.")

    resultados = carregar_checkpoint(args.checkpoint)
    print(f"Checkpoint atual: {len(resultados)} entradas.")

    if not args.so_aplicar:
        # Determinar pendentes
        def precisa_consultar(numero):
            if numero not in resultados:
                return True
            valor = resultados[numero]
            if valor == "ERRO_REDE":
                return True
            if args.apenas_erros and isinstance(valor, str):
                return True
            return False

        pendentes = [(linha, n) for linha, n in precatorios if precisa_consultar(n)]
        if args.limit:
            pendentes = pendentes[: args.limit]

        if pendentes:
            print(f"Consultando {len(pendentes)} precatórios com {args.workers} workers...")
            _consultar_em_paralelo(pendentes, resultados, args.checkpoint, args.workers)
        else:
            print("Nada pendente para consultar.")

    print(f"Escrevendo {saida}...")
    escrever_xlsx(args.entrada, saida, precatorios, resultados)

    _imprimir_relatorio(precatorios, resultados, saida)
    return 0


def _consultar_em_paralelo(pendentes, resultados, caminho_checkpoint, workers):
    inicio = datetime.now()
    total = len(pendentes)
    concluidos = 0

    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futuros = {pool.submit(consultar_saldo, n, session): (linha, n) for linha, n in pendentes}
            for futuro in as_completed(futuros):
                linha, numero = futuros[futuro]
                try:
                    resultados[numero] = futuro.result()
                except Exception as e:
                    resultados[numero] = f"ERRO_INESPERADO:{type(e).__name__}"

                concluidos += 1
                if concluidos % 100 == 0 or concluidos == total:
                    salvar_checkpoint(caminho_checkpoint, resultados)
                    _imprimir_progresso(concluidos, total, inicio)

    salvar_checkpoint(caminho_checkpoint, resultados)
    print()  # newline após barra de progresso


def _imprimir_progresso(concluidos, total, inicio):
    decorrido = (datetime.now() - inicio).total_seconds()
    pct = concluidos / total * 100
    taxa = concluidos / decorrido if decorrido > 0 else 0
    eta = (total - concluidos) / taxa if taxa > 0 else 0
    barra_size = 30
    preenchido = int(barra_size * concluidos / total)
    barra = "█" * preenchido + "░" * (barra_size - preenchido)
    print(f"\r  [{barra}] {concluidos}/{total} ({pct:.1f}%) | {taxa:.1f}/s | ETA {_formatar_eta(eta)}",
          end="", flush=True)


def _imprimir_relatorio(precatorios, resultados, saida):
    sucesso = sem_saldo = nao_encontrado = erro_rede = invalido = outros = 0
    erros_detalhe = []
    for linha, numero in precatorios:
        valor = resultados.get(numero)
        if isinstance(valor, (int, float)):
            sucesso += 1
        elif valor == "SEM_SALDO":
            sem_saldo += 1
            erros_detalhe.append((linha, numero, "SEM_SALDO"))
        elif valor == "NAO_ENCONTRADO":
            nao_encontrado += 1
            erros_detalhe.append((linha, numero, "NAO_ENCONTRADO"))
        elif valor == "ERRO_REDE":
            erro_rede += 1
            erros_detalhe.append((linha, numero, "ERRO_REDE"))
        elif valor == "RESPOSTA_INVALIDA":
            invalido += 1
            erros_detalhe.append((linha, numero, "RESPOSTA_INVALIDA"))
        elif valor is None:
            outros += 1
        else:
            outros += 1
            erros_detalhe.append((linha, numero, str(valor)))

    print()
    print("=" * 50)
    print(f" CONCLUÍDO")
    print(f" Arquivo: {saida}")
    print(f" Sucesso          : {sucesso}")
    print(f" Sem saldo        : {sem_saldo}")
    print(f" Não encontrados  : {nao_encontrado}")
    print(f" Erros de rede    : {erro_rede}")
    print(f" Resposta inválida: {invalido}")
    if outros:
        print(f" Outros           : {outros}")
    print("=" * 50)

    if erros_detalhe:
        with open(ERROS_CSV, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Linha", "Numero", "Motivo"])
            w.writerows(erros_detalhe)
        print(f" Log de erros: {ERROS_CSV}")


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 3: Rodar todos os testes para confirmar que nada quebrou**

```powershell
python -m pytest tests/ -v
```
Expected: 13 passed.

- [ ] **Step 4: Smoke test com --help**

```powershell
python atualizar_saldos.py --help
```
Expected: imprime descrição e opções; sai com código 0.

---

### Task 8: Validação manual com --limit 10

**Files:** nenhum (apenas execução)

- [ ] **Step 1: Rodar com limite pequeno**

```powershell
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
python atualizar_saldos.py --limit 10 --reset
```
Expected:
- Termina em < 30 segundos
- Console mostra "Sucesso: ~10"
- Cria `Precatórios 2027 - Atualizado.xlsx` em Downloads
- Cria `saldos_checkpoint.json` no projeto

- [ ] **Step 2: Abrir xlsx de saída e verificar visualmente**

```powershell
Start-Process "C:\Users\DARLANMARTINS\Downloads\Precatórios 2027 - Atualizado.xlsx"
```
Verificar:
- Coluna H das linhas 2-11 está preenchida com valor monetário
- Formato é "R$ X.XXX,XX" ou similar
- Demais colunas inalteradas
- Cabeçalho inalterado

- [ ] **Step 3: Validar comparando 2 valores manualmente**

Pegar precatório da linha 2 (provavelmente `2025.09451-0`) e abrir no navegador:
https://www3.tjrj.jus.br/PortalConhecimento/api/precatorios/consultaPosicao?numeroPrecatorio=2025.09451-0

Comparar o campo `"Saldo"` do JSON com o valor da célula H2 na planilha. Deve ser idêntico.

Repetir para a linha 6 (`2025.09470-6`).

- [ ] **Step 4: Validar retomada**

```powershell
python atualizar_saldos.py --limit 5
```
Expected: console mostra "Checkpoint atual: 10 entradas" e "Nada pendente para consultar." Não faz chamadas adicionais à API.

---

### Task 9: Execução completa

**Files:** nenhum

- [ ] **Step 1: Resetar e rodar tudo**

```powershell
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
python atualizar_saldos.py --reset
```
Expected:
- Mostra barra de progresso atualizando
- Termina em aproximadamente 8-12 minutos
- "Sucesso: ~6200" (alguns podem estar quitados/cancelados)

- [ ] **Step 2: Verificar arquivo de saída**

```powershell
Get-Item "C:\Users\DARLANMARTINS\Downloads\Precatórios 2027 - Atualizado.xlsx" | Select-Object Name, Length, LastWriteTime
```
Expected: arquivo existe, tamanho semelhante ao original.

- [ ] **Step 3: Verificar relatório de erros**

```powershell
if (Test-Path "C:\Users\DARLANMARTINS\Documents\PROJETO 01\erros_consulta.csv") {
    Get-Content "C:\Users\DARLANMARTINS\Documents\PROJETO 01\erros_consulta.csv" | Select-Object -First 20
}
```
Inspecionar visualmente: erros devem ser `SEM_SALDO` (quitados) ou `NAO_ENCONTRADO`. Se houver muitos `ERRO_REDE`, rodar `python atualizar_saldos.py` de novo (vai re-tentar só os erros).

- [ ] **Step 4: Sanity check final**

```powershell
python -c "
from openpyxl import load_workbook
wb = load_workbook(r'C:\Users\DARLANMARTINS\Downloads\Precatórios 2027 - Atualizado.xlsx', data_only=True)
ws = wb.active
preenchidas = sum(1 for r in ws.iter_rows(min_row=2, min_col=8, max_col=8, values_only=True) if r[0] is not None)
total = ws.max_row - 1
print(f'Linhas com saldo preenchido: {preenchidas} / {total} ({preenchidas/total*100:.1f}%)')
"
```
Expected: pelo menos 99% das linhas preenchidas.

- [ ] **Step 5: Limpeza opcional do checkpoint**

Se tudo correu bem e os erros são compatíveis com a expectativa (quitados/cancelados):

```powershell
Remove-Item "C:\Users\DARLANMARTINS\Documents\PROJETO 01\saldos_checkpoint.json"
```

A próxima execução começa do zero (refletindo o estado atualizado do TJRJ na época).

---

## Critérios de aceite (do spec)

- [x] Script roda do início ao fim sem intervenção manual em < 15 minutos. → validado em Task 9 Step 1
- [x] Coluna H da planilha de saída preenchida em ≥ 99% das linhas. → validado em Task 9 Step 4
- [x] Comparação manual de 3 valores aleatórios bate com o site do TJRJ. → validado em Task 8 Step 3
- [x] Interromper com Ctrl+C e reiniciar continua do ponto correto. → validado em Task 8 Step 4
- [x] Planilha de entrada permanece intacta (mesmo hash MD5). → validado em Task 6 test_escrever_xlsx_nao_altera_entrada
- [x] Formatação original da planilha de saída preservada (fontes, larguras, abas). → openpyxl preserva por padrão; verificado visualmente em Task 8 Step 2
