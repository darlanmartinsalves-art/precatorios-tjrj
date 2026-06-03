# Design: Mapeamento correto de precatório → ofício via OFREQ

**Data:** 2026-06-01 (revisado após descoberta do documento de vínculo DEPRE/DEPJU)
**Arquivo afetado principal:** `baixar_requisitorios.py`

## Problema

Um processo judicial pode conter **vários precatórios** (`2025.NNNNN-D`), com
**beneficiários diferentes** + precatórios de **honorários** dos advogados. Cada
precatório tem seu próprio ofício requisitório Definitivo na árvore do visualizador.

O código atual (`processar_processo` + `_baixar_e_extrair_pecas_ofreq`):

1. Para no **primeiro** ofício que extrai dados válidos (`break` em `:1038`).
2. Carimba esse **único** beneficiário em **todos** os precatórios do processo
   (`dados = {p: dados_proc for p in precatorios_do_processo}`, `:1144`).
3. Nomeia o PDF com `precatorios_do_processo[0]` — um precatório arbitrário.

Resultado: em processos multi-precatório, as colunas N–R saem com dados **errados**,
silenciosamente. Confirmado com evidência: o OFREQ `2025.06478` (escritório Carraro e
Guimarães) foi nomeado `2025.09453-6`, mas na verdade gerou o precatório `2025.06209-0`.

## A chave de ligação: número OFREQ

Descoberto com evidência (PDFs reais):

- O **ofício requisitório** (Definitivo) tem, no cabeçalho, `Definitivo OFÍCIO Nº:
  AAAA.NNNNN/OFREQ`, além do beneficiário (seção III) e advogado (seção VII).
- Existe, **sempre** (confirmado pelo usuário), um **Ofício DEPRE/DEPJU** dentro do
  processo, emitido na geração do precatório, com a frase:
  > "Ofício **{AAAA.NNNNN}/OFREQ** ... gerou o precatório **{AAAA.NNNNN-D}**."
- O **número OFREQ** aparece nos dois documentos → é a chave de join.

```
requisitório:   OFREQ ──► beneficiário + advogado + docs
vínculo DEPRE:  OFREQ ──► número do precatório
join por OFREQ: precatório ──► beneficiário/advogado corretos
```

Mapeamento **determinístico e automático** — sem chute, sem busca manual.

### Regexes validados contra PDFs reais

- Classificar requisitório: texto contém `OFÍCIO REQUISITÓRIO DE PAGAMENTO`.
- Classificar vínculo: texto contém `gerou o precatório`.
- OFREQ (ambos): `(\d{4}\.\d+)\s*/\s*OFREQ`.
- Vínculo (ofreq, precatório): `Of[íi]cio\s+(\d{4}\.\d+)\s*/\s*OFREQ.*?gerou\s+o\s+precat[óo]rio\s+(\d{4}\.\d+-\d+)` (DOTALL|IGNORECASE).

Verificado: vínculo → `(2025.06478, 2025.06209-0)`; fixture requisitório → OFREQ
`2025.14235`; os dois tipos não se confundem na classificação.

## Objetivo

Para cada processo, baixar tanto os ofícios requisitórios quanto os ofícios de vínculo
DEPRE/DEPJU, casar por OFREQ, e preencher as **colunas N–R por precatório com os dados
corretos**. Precatórios sem vínculo encontrado → marca `REVISAR`.

## Mudanças

### 1. Novos extratores (funções puras, testáveis)

- `extrair_numero_ofreq(pdf_bytes) -> str | None` — `(\d{4}\.\d+)\s*/\s*OFREQ`.
- `extrair_vinculo_ofreq_precatorio(pdf_bytes) -> (str, str) | None` — retorna
  `(ofreq, precatorio)` do ofício DEPRE/DEPJU, ou `None`.
- `eh_documento_requisitorio(texto) -> bool` — contém `OFÍCIO REQUISITÓRIO DE PAGAMENTO`.
- `eh_documento_vinculo(texto) -> bool` — contém `gerou o precatório`.

Reutilizam `_extrair_texto_pdf`.

### 2. Coleta — `_baixar_e_extrair_pecas_ofreq`

- **Remover o `break`** (`:1038`): iterar todos os candidatos e baixar/classificar cada PDF.
- Para cada PDF baixado, classificar:
  - **requisitório** → guardar `{ofreq, beneficiario, advogado, pdf_path, pdf_bytes}`.
  - **vínculo** → guardar no mapa `vinculos[ofreq] = precatorio`.
  - escaneado/sem texto → `manual_revisar` (comportamento atual mantido).
- **Deduplicar** requisitórios por `ofreq` (a árvore pode listar a mesma peça 2x).
- Pode ser preciso elevar `LIMITE_CANDIDATOS` (processos com muitos precatórios têm
  mais peças). Definir valor no plano.
- Retorna `(requisitorios: list, vinculos: dict[ofreq -> precatorio])`.

### 3. Join — `processar_processo`

- Receber `(requisitorios, vinculos)`.
- Para cada requisitório, achar seu precatório: `precatorio = vinculos.get(req.ofreq)`.
- Montar `dados = {precatorio: {beneficiario_nome, beneficiario_doc, advogado_nome,
  advogado_cpf, advogado_oab}}` — agora **um beneficiário por precatório, correto**.
- Renomear cada PDF para `"{precatorio} - {beneficiario}.pdf"` (precatório agora é o
  **verdadeiro**, vindo do join).
- Requisitório cujo OFREQ não tem vínculo → registrar com chave/flag `REVISAR`
  (não descartar; salvar PDF para conferência).

### 4. Saída — manter `atualizar_planilha` (N–R)

- A função `atualizar_planilha` (já testada) **continua sendo usada**, agora alimentada
  com o `dados` correto por precatório. Será necessário estender `COLUNAS_NOVAS` para
  incluir a coluna **S = Status** (coluna 19).
- Precatórios do processo (na planilha filtrada) que **não** receberam vínculo →
  deixar N–R em branco e gravar `"REVISAR"` na coluna **S (Status)** dedicada, e logar.
  (Decisão do usuário: coluna de status separada, mantendo N–R limpas.)

### 5. Checkpoint (`downloads_state.json`)

- Estrutura por processo mantém `dados` (dict precatório→dados), agora correto.
- `dados_acumulados` em `main()` permanece `{precatorio: dados}`.
- O checkpoint atual (18 processos, com os 3 "ok" errados) deve ser descartado:
  **rodar com `--reset`** após a mudança.

## Componentes e isolamento

| Unidade | Faz | Testável isolada? |
|---|---|---|
| `extrair_numero_ofreq` | regex OFREQ no texto | ✅ fixture |
| `extrair_vinculo_ofreq_precatorio` | regex vínculo | ✅ (fixture do doc DEPRE/DEPJU — adicionar) |
| `eh_documento_requisitorio` / `eh_documento_vinculo` | classificação por texto | ✅ |
| `_baixar_e_extrair_pecas_ofreq` | navega, baixa, classifica, dedup | ⚠️ browser; lógica de dedup/join extraível p/ helper puro |
| join em `processar_processo` | casa OFREQ → precatório | ✅ se extraído p/ função pura |
| `atualizar_planilha` | grava N–R (existente) | ✅ (já testada) |

## Testes (TDD)

- **Adicionar fixture** `tests/fixtures/vinculo_depre_modelo.pdf` — cópia de
  `C:\Users\DARLANMARTINS\Downloads\00275429720148190001.pdf` (aprovado pelo usuário).
- `extrair_vinculo_ofreq_precatorio` → `("2025.06478", "2025.06209-0")`.
- `extrair_numero_ofreq` (fixture requisitório) → `"2025.14235"`.
- `eh_documento_requisitorio` / `eh_documento_vinculo`: separam corretamente os 2 tipos.
- Extratores com PDF inválido/vazio → `None`.
- Helper de join puro (se extraído): dado `requisitorios` + `vinculos`, produz
  `{precatorio: dados}` correto; OFREQ sem vínculo → `REVISAR`; dedup por OFREQ.
- `atualizar_planilha`: testes existentes continuam verdes.

## Fora de escopo

- Busca por número do precatório (Numeração Antiga + Origem Precatórios Judiciais) como
  fluxo automático — fica como **conferência manual** do usuário, não codificada.
- Diagnóstico dos `erro_navegacao` (frente separada).
- Limpezas de código morto (`--workers`, `SCREENSHOTS_DIR`) — frente separada.
