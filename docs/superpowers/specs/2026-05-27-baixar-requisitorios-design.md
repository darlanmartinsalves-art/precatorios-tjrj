# Download automatizado de ofícios requisitórios do TJRJ

**Data:** 2026-05-27
**Autor:** Darlan Martins + Claude
**Status:** Design aprovado
**Etapa anterior:** [2026-05-27-atualizar-saldos-precatorios-design.md](2026-05-27-atualizar-saldos-precatorios-design.md)

## Contexto e motivação

A primeira etapa do projeto preencheu os saldos atualizados de 6.288 precatórios na coluna H da planilha `Precatórios 2027.xlsx`. Esta segunda etapa baixa o **ofício requisitório (OFREQ)** em PDF de cada processo judicial relacionado, a partir do portal do TJRJ (`www3.tjrj.jus.br/portalservicos`).

Trabalho **recorrente** — provavelmente mensal/trimestral conforme novos precatórios entram na lista.

## Escopo

**Filtro de entrada:** somente precatórios cujo `saldo atualizado >= R$ 200.000`.

- 2.350 precatórios atendem o critério
- Resultam em 1.952 processos judiciais únicos (um processo pode ter vários precatórios)
- Cobrem R$ 2,24 bilhões (≈ 86% do valor total dos precatórios da planilha)

**Não escopo:**
- Outras peças do processo além do ofício requisitório
- Alteração da planilha de entrada
- Interface gráfica
- Processos com saldo < R$ 200.000 (configurável via `--saldo-minimo`)

## Origem dos dados

A planilha **`Precatórios 2027 - Atualizado.xlsx`** (gerada na Etapa 1) tem:

| Coluna | Conteúdo |
|---|---|
| B | Número do precatório (ex: `2025.09451-0`) |
| F | Número do processo judicial padrão CNJ (ex: `0156129-30.2020.8.19.0001`) |
| H | Saldo atualizado (numérico, valor em reais) |

O script lê todas as linhas, aplica o filtro de saldo, agrupa por processo único.

## Estrutura do ofício requisitório

PDF gerado dinamicamente pelo portal TJRJ. Identificado por:
- Numeração `AAAA.NNNNN/OFREQ` (sufixo OFREQ no número do ofício)
- Cabeçalho `OFÍCIO REQUISITÓRIO DE PAGAMENTO DE VERBA [ALIMENTÍCIA|NÃO ALIMENTÍCIA]`
- Conteúdo em seções (I a XII) sempre presentes:
  - III - BENEFICIÁRIO (Nome, CNPJ/CPF, Tipo)
  - IV - DADOS PROCESSUAIS
  - VII - ADVOGADOS DO BENEFICIÁRIO
  - VIII - HONORÁRIOS
  - XII - DADOS BANCÁRIOS

Exemplo: `C:\Users\DARLANMARTINS\Downloads\GARRASTAZU - REQUISITÓRIO.pdf` (3 páginas, 133 KB).

## Arquitetura

### Stack
- **Python 3.12** (já instalado)
- **Playwright para Python** (`playwright`) — automação Chromium
- **pypdf** — extração de texto do PDF
- **openpyxl** — leitura da planilha (já instalado)
- **pytest** — testes unitários (já instalado)

### Layout de arquivos
```
C:\Users\DARLANMARTINS\Documents\PROJETO 01\
  baixar_requisitorios.py          # script principal
  requirements.txt                 # adicionar: playwright, pypdf
  user_data_chromium/              # sessão persistente Chromium (cookies/cache)
  downloads_state.json             # checkpoint atômico
  erros_download.csv               # gerado se houver erros
  screenshots_erro/                # PNGs capturados em erros (debug)
  tests/
    test_baixar_requisitorios.py
    fixtures/
      garrastazu_modelo.pdf        # PDF de referência para testar parsing

C:\Users\DARLANMARTINS\Downloads\
  Precatórios 2027 - Atualizado.xlsx          # entrada (read-only)
  Precatórios_Requisitórios\                  # destino dos PDFs
    2025.09451-0 - TECHNE ENGENHARIA.pdf
    2025.09451-0 - TECHNE ENGENHARIA (2).pdf  # se múltiplos requisitórios
    ...
    manual_revisar\                            # PDFs sem beneficiário identificado
      2025.09451-0 - SEM_NOME.pdf
```

### Princípios chave

1. **Sessão persistente:** Playwright `launch_persistent_context(user_data_chromium/)` salva cookies e localStorage entre execuções. Primeira execução pede login manual; subsequentes reusam.
2. **Idempotência:** re-execução pula processos com status `ok` ou `sem_requisitorio`.
3. **Browser visível por padrão:** `headless=False` para acompanhar visualmente e lidar com captcha/sessão expirada.
4. **Planilha intocada:** somente leitura.
5. **Checkpoint atômico:** mesmo padrão da Etapa 1 (`.tmp` + rename).

## Fluxo de dados

### Fase 1 — Setup e filtro
```
1. Ler "Precatórios 2027 - Atualizado.xlsx"
2. Filtrar: saldo (coluna H) >= --saldo-minimo (padrão 200.000)
3. Construir lista [(precatorio, processo), ...]
4. Agrupar por processo_judicial único
5. Carregar downloads_state.json
6. Filtrar pendentes: status != "ok" (e != "sem_requisitorio" salvo --apenas-erros)
```

### Fase 2 — Login (primeira execução)

Detecta se já há sessão válida:
- Tenta abrir uma URL autenticada (ex: `/portalservicos/#/painel`)
- Se redireciona pra login → sessão inválida

Se sessão inválida:
```
1. Abre browser visível em https://www3.tjrj.jus.br/portalservicos/
2. Console: "Faça login no portal. Quando estiver logado, pressione ENTER aqui."
3. Aguarda input do usuário
4. Confirma com nova checagem de autenticação
5. Salva flag (timestamp) e prossegue
```

### Fase 3 — Download paralelo

Pool com `--workers` (padrão 3) tarefas async; cada uma com sua page (tab) compartilhando o mesmo context:

```
para cada processo na fila:
  1. page.goto("https://www3.tjrj.jus.br/portalservicos/#/consproc/consultaportal")
  2. wait_for_selector(formulário CNJ)
  3. selecionar radio "Única" (numeração)
  4. split_cnj("0156129-30.2020.8.19.0001") → ("0156129-30.2020", "0001")
  5. preencher 2 campos do CNJ
  6. click "Pesquisar"
  7. wait_for_selector(tabela de resultados ou mensagem "sem resultados")
  8. se "sem resultados" → status=processo_nao_encontrado, próximo
  9. click no link do processo
  10. context.expect_page() captura nova aba
  11. nova aba: click "Processo Eletrônico - Visualizador"
  12. context.expect_page() captura aba do visualizador
  13. no visualizador: localizar todas as peças com "OFREQ" ou "REQUISITÓRIO" no nome
  14. para cada peça encontrada:
        a. page.expect_download() + click pra baixar
        b. salvar em temp/ com nome temporário
        c. pypdf: extrair texto da página 1
        d. regex: "III - BENEFICIÁRIO.*?Nome:\\s*(.+)"  → beneficiário
        e. sanitizar nome (substituir /\\:*?"<>| por _)
        f. nome final: "{precatorio} - {beneficiario}.pdf"
        g. se já existe arquivo com mesmo nome: incrementar " (2)", " (3)"...
        h. mover do temp/ para Precatórios_Requisitórios/
  15. fechar abas extras (visualizador e processo); manter só a aba de consulta para próximo
  16. checkpoint[processo] = {"status": "ok", "arquivos": [...], "ts": now}
  17. a cada 10 processos: salvar checkpoint atomicamente
```

### Fase 4 — Relatório

```
Console:
  Processos OK            : X  (Y PDFs baixados)
  Sem requisitório        : Z
  Processo não encontrado : W
  Erros de rede           : E1
  Erros de navegação      : E2
  Erros de parsing        : E3
  Tempo total             : HH:MM:SS

Se E1+E2+E3 > 0: gera erros_download.csv com (precatorio, processo, status, motivo, screenshot)
```

## Tratamento de erros

### Tabela de status

| Status no checkpoint | Significado | Re-tenta em re-execução? |
|---|---|---|
| `ok` | Baixou 1+ PDFs com sucesso | Não |
| `sem_requisitorio` | Processo existe, sem peça OFREQ | Só com `--apenas-erros` |
| `processo_nao_encontrado` | Portal retornou sem resultados | Só com `--apenas-erros` |
| `erro_rede` | Timeout / conn refused | **Sim, automaticamente** |
| `erro_navegacao` | Selector não encontrado, página mudou | **Sim, automaticamente** |
| `erro_parsing` | Baixou PDF mas não extraiu beneficiário | **Sim, automaticamente** (PDF fica em manual_revisar/) |
| ausente | Não processado ainda | Sim |

### Comportamento por erro

**Erros de rede (timeout / conn refused / 5xx):**
- Retry 3x com backoff 5s → 10s → 20s
- Se persistir: status=`erro_rede`, salva screenshot, próximo processo

**Selector não encontrado (UI mudou):**
- Recarrega a página, retry 1x
- Se persistir: status=`erro_navegacao`, salva screenshot

**Login expirado durante execução:**
- Detecta redirect para `/login` ou aparição de campo senha
- Pausa todos os workers
- Console: "Sessão expirou. Faça login no browser visível e pressione ENTER"
- Retoma após confirmação

**Captcha:**
- Detecta iframe reCAPTCHA ou similar
- Pausa workers, pede pra resolver manual, ENTER pra continuar

**PDF baixou mas beneficiário não foi extraído:**
- Não re-tenta o download (o PDF está ok)
- Salva como `{precatorio} - SEM_NOME.pdf` em `Precatórios_Requisitórios/manual_revisar/`
- Status final: `erro_parsing` (não conta como sucesso, mas o PDF está disponível)

**Ctrl+C:**
- Captura SIGINT
- Workers terminam tarefa atual
- Salva checkpoint
- Sai com código 0

## Interface CLI

```
python baixar_requisitorios.py [opções]

  --entrada PATH        xlsx fonte (padrão: ~/Downloads/Precatórios 2027 - Atualizado.xlsx)
  --saida PATH          Pasta destino (padrão: ~/Downloads/Precatórios_Requisitórios/)
  --saldo-minimo VALOR  Filtro de saldo mínimo em reais (padrão: 200000)
  --workers N           Abas paralelas (padrão: 3)
  --limit N             Processa apenas N processos pendentes (testes)
  --reset               Apaga checkpoint
  --apenas-erros        Re-tenta tudo que não é "ok" (inclui sem_requisitorio e nao_encontrado)
  --processo NNN        Roda só um processo específico (debug)
  --headless            Sem janela visível (recomendado só após primeira execução)
```

## Sub-componentes (testáveis)

### Funções puras (unit tests)

1. **`split_cnj(numero: str) -> tuple[str, str]`**
   - `"0156129-30.2020.8.19.0001"` → `("0156129-30.2020", "0001")`
   - Lança `ValueError` se não casar com padrão CNJ

2. **`extrair_beneficiario(pdf_bytes: bytes) -> str | None`**
   - Usa `pypdf` para extrair texto
   - Regex: `r"III\s*-\s*BENEFICI[ÁA]RIO.*?Nome:\s*(.+?)(?:\n|CNPJ|CPF)"`
   - Retorna nome limpo ou `None` se não encontrar

3. **`sanitizar_nome(nome: str) -> str`**
   - Remove `/\\:*?"<>|`, substitui por `_`
   - Colapsa múltiplos espaços
   - Limita a 100 caracteres
   - Strip de espaços nas pontas

4. **`gerar_nome_arquivo(precatorio, beneficiario, ordinal=None, pasta=None) -> Path`**
   - `("2025.09451-0", "TECHNE ENG", None, /tmp)` → `Path("/tmp/2025.09451-0 - TECHNE ENG.pdf")`
   - Se `ordinal=2`: `Path("/tmp/2025.09451-0 - TECHNE ENG (2).pdf")`
   - Se arquivo já existe: incrementa ordinal automaticamente

5. **`filtrar_precatorios(linhas: list, saldo_minimo: float) -> list[tuple[str, str]]`**
   - Recebe lista de linhas (precatório, processo, saldo)
   - Retorna `[(precatorio, processo)]` com saldo >= mínimo

6. **`carregar_checkpoint(path) / salvar_checkpoint(path, dados)`**
   - Mesmo padrão atômico da Etapa 1

### Funções de integração (testes manuais com portal real)

- `consultar_processo(page, processo) -> dict` — abre consulta e clica até abrir o visualizador
- `localizar_pecas_ofreq(page) -> list` — encontra peças com "OFREQ" no nome
- `baixar_peca(page, peca, destino) -> Path` — clica e captura o download

## Fase de exploração da UI (pré-requisito)

Antes de implementar a navegação, é necessária uma **sessão de exploração interativa** com o portal logado, para mapear:

1. URL e selector CSS/XPath do form de consulta processual
2. Selector dos 2 campos do CNJ (prefixo e sufixo)
3. Selector do radio "Única"
4. Selector do botão "Pesquisar"
5. Selector do link/botão para abrir o processo na nova aba
6. Selector do botão "Processo Eletrônico - Visualizador" (segunda nova aba)
7. Estrutura DOM do visualizador (como peças/movimentos são listados)
8. Como identificar peças com "OFREQ" no nome ou descrição
9. Como triggerar o download da peça (clique direto? botão de download? URL de PDF embutida?)
10. Comportamento quando o processo não tem requisitório

Essa exploração é feita com Playwright em modo `--processo XXX` interativo, com pausas para inspeção DOM, e o resultado vira código no script. Esse trabalho consome ~30min mas só precisa ser feito uma vez.

## Validação antes da execução completa

1. **Teste 1 — `--processo 0156129-30.2020.8.19.0001`** (Garrastazu/TECHNE)
   - Esse é o processo de referência (já temos o PDF baixado manualmente)
   - Deve baixar 1 PDF nomeado "2025.NNNNN-N - TECHNE ENGENHARIA E SISTEMAS LTDA.pdf"
   - Comparar visualmente com o PDF manual

2. **Teste 2 — `--limit 5`**
   - Roda 5 processos diversos
   - Verifica que paralelismo funciona, checkpoint persiste, nomes corretos

3. **Teste 3 — Retomada**
   - `--limit 10`, interromper aos 5 com Ctrl+C
   - Re-executar: deve pular os 5 já baixados, completar os 5 restantes

4. **Teste 4 — Execução completa**
   - Sem `--limit`: roda os ~1.952 processos
   - Estimativa: 4-7h com 3 workers

## Critérios de sucesso

- [ ] Login manual uma vez, sessão persiste entre execuções
- [ ] Filtra corretamente saldos >= R$ 200.000 (resultando em ~1.952 processos únicos)
- [ ] ≥ 90% dos processos resultam em PDF baixado com sucesso
- [ ] Nomes de arquivo correspondem ao beneficiário do PDF (seção III)
- [ ] Múltiplos requisitórios no mesmo processo são salvos com numeração `(2)`, `(3)`...
- [ ] Ctrl+C salva checkpoint; re-execução continua de onde parou sem duplicar
- [ ] Planilha de entrada permanece inalterada (mesmo hash MD5)
- [ ] Screenshots gerados em erros para diagnóstico

## Out of scope (futuras melhorias)

- Extrair dados estruturados do PDF (valor bruto, datas, advogado, conta bancária) para enriquecer a planilha
- OCR para PDFs escaneados
- Notificação por e-mail quando termina
- Re-validação periódica (verificar se algum PDF foi atualizado)
- Suporte a outras peças do processo
- Integração com sistemas de gestão de escritório
