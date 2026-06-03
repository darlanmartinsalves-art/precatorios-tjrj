# Pipeline Unificado TJRJ — Download + Extração + Atualização da Planilha

**Data:** 2026-05-27
**Autor:** Darlan Martins + Claude
**Status:** Design aprovado
**Substitui:** [2026-05-27-baixar-requisitorios-design.md](2026-05-27-baixar-requisitorios-design.md) (Etapa 2 original, agora redesenhada)
**Próxima etapa:** Etapa 4 (telefones via confirmeonline) — escopo separado

## Contexto

A Etapa 2 original previa apenas download dos PDFs. Após análise, o escopo foi expandido para também extrair dados estruturados do PDF (beneficiário e advogado) e atualizar a planilha em **uma única passagem por processo**. Isso evita re-navegar o TJRJ depois.

Tasks 1-7 da Etapa 2 anterior continuam válidas e foram concluídas (36 testes passando). Tasks 8-13 são redesenhadas neste spec.

## Mudança crítica: Chrome via CDP

O portal TJRJ exige **código verificador** a cada login (provavelmente 2FA). Usar Playwright com Chromium bundled e persistent context não resolve isso de forma confiável. Solução:

1. Usuário abre Chrome com flag de debugging: `chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\temp\chrome_debug`
2. Usuário loga manualmente no TJRJ (resolve código verificador uma vez)
3. Script Python conecta no Chrome aberto via `playwright.chromium.connect_over_cdp("http://localhost:9222")`
4. Script abre novas abas e automatiza tudo, **reutilizando a sessão autenticada**

Vantagens:
- Resolve código verificador (login manual real)
- Site não detecta como bot (é Chrome de verdade)
- Sessão dura enquanto Chrome estiver aberto

## Escopo

**Filtro de entrada (inalterado):** precatórios com `saldo atualizado >= R$ 200.000`.
- 2.350 precatórios → 1.952 processos judiciais únicos

**Processamento por processo:**
1. Navega no TJRJ até o visualizador
2. Baixa todos os PDFs de ofício requisitório (OFREQ) encontrados
3. Extrai do PDF baixado:
   - Seção III: Nome + Documento (CPF ou CNPJ) do Beneficiário
   - Seção VII: Nome + CPF + OAB do **primeiro** advogado (ignora demais)
4. Salva PDF: `~/Downloads/Precatórios_Requisitórios/{precatório} - {beneficiário sanitizado}.pdf`
5. Atualiza colunas N–R da planilha:
   | Coluna | Conteúdo |
   |---|---|
   | N | Beneficiário Nome |
   | O | Beneficiário Doc (CPF formatado xxx.xxx.xxx-xx ou CNPJ xx.xxx.xxx/xxxx-xx) |
   | P | Advogado Nome |
   | Q | Advogado CPF (formatado) |
   | R | Advogado OAB |

**Não escopo (Etapa 4 separada):**
- Busca de telefones no confirmeonline (depende de O e Q)
- Atualização de colunas S, T

## Origem dos dados de entrada

Planilha: `Precatórios 2027 - Atualizado.xlsx` (gerada na Etapa 1)
- B = número do precatório
- F = número do processo judicial CNJ
- H = saldo atualizado (usado para filtro)
- I-M = Cedente/CPF/Celular/E-mail/OBSERVAÇÃO (inalterados)
- **Novas colunas a adicionar:** N, O, P, Q, R

## Estrutura do PDF (referência)

PDF de exemplo: `GARRASTAZU - REQUISITÓRIO.pdf`. Seções relevantes:

```
III - BENEFICIÁRIO
Nome: TECHNE ENGENHARIA E SISTEMAS LTDA
CNPJ: 50737766000121
Tipo do beneficiário: Beneficiario

VII - ADVOGADOS DO BENEFICIÁRIO
RJ185918 - ARTUR GARRASTAZU GOMES FERREIRA - 33394784068
```

Padrão da linha VII: `{OAB} - {NOME} - {CPF}`

Para beneficiário PF, o campo será `CPF: XXXXXXXXXXX`. Para PJ, `CNPJ: XXXXXXXXXXXXXX`.

## Arquitetura

### Stack
- **Python 3.12** (já instalado)
- **Playwright** com `connect_over_cdp` (já instalado)
- **pypdf** (já instalado)
- **openpyxl** (já instalado)
- **Chrome real** (não Chromium bundled) com debug port

### Layout de arquivos
```
C:\Users\DARLANMARTINS\Documents\PROJETO 01\
  baixar_requisitorios.py        # arquivo único — funções + orquestração
  seletores.py                   # selectors DOM (gerado na fase de exploração)
  iniciar_chrome_debug.ps1       # NOVO: script PowerShell que abre Chrome com debug port
  downloads_state.json           # checkpoint
  erros_download.csv             # log
  screenshots_erro/              # debug
  tests/test_baixar_requisitorios.py   # já tem 23 testes; vai crescer pras novas funções
  tests/conftest.py              # já tem fixtures

C:\Users\DARLANMARTINS\Downloads\
  Precatórios 2027 - Atualizado.xlsx              # entrada (read-only)
  Precatórios 2027 - Atualizado - Etapa2.xlsx     # SAÍDA — entrada com colunas N-R preenchidas
  Precatórios_Requisitórios\
    2025.09451-0 - TECHNE ENGENHARIA.pdf
    ...
```

### Princípios

1. **Chrome real via CDP:** script não lança browser, apenas conecta
2. **Planilha de saída separada:** entrada permanece intocada (`Precatórios 2027 - Atualizado.xlsx`), saída é `... - Etapa2.xlsx`
3. **Checkpoint per processo:** se cair na metade, retoma onde parou
4. **Salvamento de planilha incremental:** a cada 10 processos, regrava o xlsx (não perde dados se algo travar)

## Fluxo

### Fase 1 — Bootstrap
```
1. Carregar planilha de entrada
2. Filtrar saldo >= 200000 → lista de (precatorio, processo)
3. Construir mapa processo → lista_precatorios_que_apontam_pra_ele
   (alguns processos têm múltiplos precatórios)
4. Carregar checkpoint downloads_state.json
5. Filtrar pendentes
```

### Fase 2 — Conectar no Chrome
```
1. Tentar conectar via CDP em http://localhost:9222
2. Se falhar: exibir instruções de como abrir Chrome com debug port
3. Verificar se há sessão TJRJ ativa (abrir URL_BASE, checar)
4. Se não estiver logado: instruir usuário a logar e pressionar ENTER
```

### Fase 3 — Pipeline por processo (3 abas paralelas async)
```
para cada processo pendente:
  1. Nova aba: navegar URL_CONSULTA
  2. Preencher CNJ split (prefix, sufix)
  3. Clicar Pesquisar
  4. Aguardar resultado ou "sem resultados"
  5. Se sem resultados → status="processo_nao_encontrado", próximo
  6. Clicar no link do processo (abre nova aba)
  7. Nessa nova aba, clicar "Processo Eletrônico - Visualizador" (abre outra aba)
  8. No visualizador: localizar peças OFREQ
  9. Se não houver: status="sem_requisitorio", próximo
  10. Para cada peça OFREQ:
        a. Baixar PDF
        b. Extrair texto (pypdf)
        c. extrair_beneficiario(texto) → (nome, doc)
        d. extrair_advogado(texto) → (nome, cpf, oab)  ← NOVO
        e. Salvar PDF com nome {precatorio} - {beneficiario sanitizado}.pdf
        f. Para CADA precatório que aponta pra este processo:
            - Gravar nas colunas N-R do row do precatório
  11. Fechar abas intermediárias
  12. Atualizar checkpoint
  13. A cada 10 processos: salvar checkpoint + regravar planilha de saída
```

### Fase 4 — Relatório
```
Console:
  Processos OK            : X (Y PDFs baixados)
  Sem requisitório        : Z
  Não encontrados         : W
  Erros (rede/nav/parsing): E
  Linhas planilha atualiz.: K
  Tempo total             : HH:MM:SS

Se erros: gera erros_download.csv com detalhes
```

## Funções novas

### `extrair_beneficiario_completo(pdf_bytes) -> dict | None`
Substitui o atual `extrair_beneficiario` (que só retorna nome). Retorna:
```python
{
  "nome": "TECHNE ENGENHARIA E SISTEMAS LTDA",
  "doc": "50737766000121",     # raw, só dígitos
  "tipo_doc": "CNPJ"           # "CPF" ou "CNPJ"
}
```
Ou `None` se não conseguir extrair.

A função atual `extrair_beneficiario` é mantida (retrocompat com testes) — pode delegar para `extrair_beneficiario_completo`.

### `extrair_advogado(pdf_bytes) -> dict | None`
Extrai o **primeiro** advogado da seção VII. Retorna:
```python
{
  "nome": "ARTUR GARRASTAZU GOMES FERREIRA",
  "cpf": "33394784068",         # raw, só dígitos
  "oab": "RJ185918"
}
```
Ou `None` se não conseguir.

Padrão da linha: `{OAB} - {NOME} - {CPF (11 dígitos)}`

### `formatar_doc(doc: str, tipo: str) -> str`
- CPF: `33394784068` → `333.947.840-68`
- CNPJ: `50737766000121` → `50.737.766/0001-21`

### `conectar_chrome_cdp(porta=9222) -> BrowserContext`
Substitui `abrir_context`. Conecta no Chrome rodando em `localhost:{porta}`. Levanta `RuntimeError` se Chrome não estiver acessível.

### `atualizar_planilha(caminho_entrada, caminho_saida, dados_extraidos)`
- Abre planilha de entrada
- Para cada (precatório, nome_benef, doc_benef, nome_adv, cpf_adv, oab_adv), localiza a linha pelo número do precatório e preenche N–R
- Salva como planilha de saída
- Chamada periodicamente durante a execução (não só no final)

## Tratamento de erros

Mantido similar ao spec original, com adição:

| Status | Significado | Re-tenta? |
|---|---|---|
| `ok` | Baixou PDFs e extraiu dados | Não |
| `sem_requisitorio` | Não encontrou OFREQ | Só com `--apenas-erros` |
| `processo_nao_encontrado` | Portal sem resultados | Só com `--apenas-erros` |
| `erro_rede` | Timeout / conn refused | Sim |
| `erro_navegacao` | Selector mudou | Sim |
| `erro_parsing` | Baixou PDF mas falhou extrair beneficiário | Sim |
| `erro_chrome_desconectado` | Chrome foi fechado | **Não retoma silencioso** — usuário precisa reabrir Chrome |

## Script de bootstrap do Chrome (`iniciar_chrome_debug.ps1`)

Para facilitar o usuário, criar script PowerShell que:
1. Mata processos Chrome existentes (opcional, com prompt)
2. Inicia Chrome com `--remote-debugging-port=9222` e `--user-data-dir=C:\temp\chrome_debug`
3. Imprime instruções: "Faça login no TJRJ e no confirmeonline. Depois rode o script Python."

## CLI

```
python baixar_requisitorios.py [opções]

  --entrada PATH          xlsx fonte (padrão: ~/Downloads/Precatórios 2027 - Atualizado.xlsx)
  --saida PATH            xlsx saída (padrão: <entrada com ' - Etapa2' antes de .xlsx)
  --pasta-pdfs PATH       Pasta de PDFs (padrão: ~/Downloads/Precatórios_Requisitórios/)
  --saldo-minimo VALOR    Filtro (padrão: 200000)
  --workers N             Abas paralelas (padrão: 3)
  --cdp-port N            Porta debug do Chrome (padrão: 9222)
  --limit N               Limite de processos (testes)
  --reset                 Apaga checkpoint
  --apenas-erros          Re-tenta tudo que não é ok
  --processo NNN          Só este processo (debug)
```

## Validação

1. **Teste 1 — funções puras (TDD):** novos `extrair_beneficiario_completo`, `extrair_advogado`, `formatar_doc` com fixture GARRASTAZU
2. **Teste 2 — `atualizar_planilha`:** xlsx fixture, preenche N-R, verifica saída
3. **Teste 3 — conexão CDP:** abrir Chrome com debug port manualmente, rodar `--test-cdp` que apenas conecta e confirma
4. **Teste 4 — `--processo` único** (Garrastazu): pipeline completo num único processo
5. **Teste 5 — `--limit 5`** com 3 workers
6. **Teste 6 — execução completa**

## Critérios de sucesso

- [ ] Conecta no Chrome aberto via CDP (não lança browser próprio)
- [ ] Filtra >= R$ 200k → ~1.952 processos
- [ ] ≥ 90% dos processos resultam em PDF + dados extraídos
- [ ] Planilha saída tem colunas N–R preenchidas para esses 90%+
- [ ] Planilha entrada permanece inalterada (hash MD5 igual)
- [ ] Documentos formatados: CPF como `xxx.xxx.xxx-xx`, CNPJ como `xx.xxx.xxx/xxxx-xx`
- [ ] Múltiplos requisitórios numerados `(2)`, `(3)`
- [ ] Ctrl+C salva checkpoint e planilha; re-execução continua

## Regra de desempate (múltiplos OFREQ no mesmo processo)

Quando um processo tem mais de um ofício requisitório (ex: "Definitivo" e "Provisório", ou um por precatório se houver mais de um):

- **Todos os PDFs são baixados** (com sufixo `(2)`, `(3)`...)
- **Para extração de dados que vão pra planilha:** prioridade do que tem `Definitivo OFÍCIO` no cabeçalho > primeiro encontrado
- Se mesmo assim houver ambiguidade (vários definitivos), usa o primeiro encontrado e registra a ambiguidade em `erros_download.csv` (status="ok" mas com observação)

## Out of scope

- Etapa 4 (telefones via confirmeonline) — spec separado
- Múltiplos advogados (só extrair o primeiro)
- Honorários, dados bancários, datas (seções IV, VI, VIII, XII)
- OCR de PDFs escaneados
