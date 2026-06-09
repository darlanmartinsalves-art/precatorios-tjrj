# Design — Localizar requisitório + vínculo DEPJU classificando pela PRÉVIA (Abordagem 2 + fallback 1)

**Data:** 2026-06-09
**Status:** Aprovado pelo usuário (brainstorming). Próximo: plano de implementação.

## Problema

Para cada precatório-alvo (saldo ≥ R$ 200k), o pipeline precisa achar 2 documentos no
visualizador do processo TJRJ:

1. **Ofício requisitório** (`OFÍCIO REQUISITÓRIO DE PAGAMENTO`, com `Definitivo OFÍCIO Nº
   AAAA.NNNNN/OFREQ` e `III - BENEFICIÁRIO`) → fonte de beneficiário/CPF e advogado; é o PDF entregável.
2. **Ofício de vínculo DEPRE/DEPJU** (`Ofício AAAA.NNNNN/OFREQ ... gerou o precatório AAAA.NNNNN-D`)
   → mapa `OFREQ → número do precatório`. Não é entregável; só liga requisitório ao precatório.

### Causa raiz do problema atual (regressão de 03/06)

O rótulo desses nós na árvore é **inconsistente**, até dentro do mesmo processo
(confirmado em screenshots do usuário):
- requisitório pode aparecer como `"<Nome do Beneficiário> Def"`, `"Def <Nome>"`,
  `"PREC DEFINITIVO ..."`, ou `"NNN - <docnum> - Petição"` genérico;
- vínculo DEPJU costuma ser `"NNN - <docnum> - Petição"` genérico.

A seleção introduzida em 03/06 passou a **escolher qual nó clicar pelo RÓTULO**
(`PADROES_PECA_REQUISITORIO` no passo 1, `PADROES_DOC_GENERICO` no passo 2). Rótulos
fora desses padrões (ex: `"Def Roberto"`) nunca eram baixados → `sem_requisitorio` falso.
O fix de 2026-06-09 (commit d0790f0) já removeu o filtro de rótulo do fallback (varre
todas as folhas do fim) — corrige a CORRETUDE, mas classifica **baixando cada PDF**
(~5-8s/nó), o que fica lento nos processos onde os alvos estão "no meio".

### Premissas validadas com o usuário

- A versão antiga (rápida) **lia do FIM pro início e abria poucos** documentos.
- Os 2 documentos ficam **sempre/quase sempre no fim** da árvore (eventos mais recentes,
  perto da geração do precatório); em **alguns processos antigos** podem estar mais no meio.
- O sinal confiável é o **CONTEÚDO** do documento, não o rótulo.

## Objetivo

Localizar os 2 documentos de forma **rápida E confiável** ("acha sempre"), inclusive nos
casos "no meio", reduzindo o custo por nó inspecionado.

## Abordagem escolhida: classificar pela PRÉVIA, baixar só os alvos confirmados

Quando um nó é clicado, o visualizador renderiza o documento na **prévia à direita**.
Ler esse texto via DOM (~1-2s) é muito mais barato que baixar o PDF (~5-8s). Classificamos
pela prévia e **só baixamos os documentos confirmados como requisitório**.

### A. Fluxo geral (uma passada, do fim pro início)

Para cada FOLHA (documento real; sem nós-pai de juntada), varrendo do **fim** pro início:

1. Clica o nó → lê o **texto da prévia** (DOM), sem baixar.
2. Classifica o texto com `eh_documento_vinculo` / `eh_documento_requisitorio` (já existentes).
3. Age conforme o tipo (seção B).
4. **goal-stop**: para assim que todos os `precatorios_alvo` do processo têm vínculo
   (`OFREQ → precatório`) **e** requisitório daquele OFREQ. Alcance do fim como teto de
   segurança (generoso, pois cada nó agora é barato → cobre os casos "no meio").

### B. Decisão por nó

| Classificação (pela prévia) | Ação |
|---|---|
| **Vínculo DEPJU** | Extrai `OFREQ → precatório` **do texto da prévia**. **Não baixa** (não é entregável). |
| **Requisitório** | **Baixa o PDF** (necessário p/ beneficiário/CPF/advogado e é o entregável). Extrai como hoje (dedup por OFREQ, beneficiário, advogado; nomeia `{precatório} - {beneficiário}.pdf`). |
| **Nenhum dos dois** | Pula. |
| **Prévia vazia / imagem** (sem texto extraível) | **Fallback Abordagem 1**: baixa aquele nó e classifica pelo PDF. Degradação graciosa, por nó. |

No caso comum, **só requisitórios são baixados** (1-2 por processo); o resto é lido de graça.

### C. Refactors necessários

Os extratores de OFREQ/vínculo hoje recebem `pdf_bytes`. Para classificar pela prévia,
criar versões **por texto** (puras, testáveis):

- `extrair_numero_ofreq_de_texto(texto) -> str | None`
- `extrair_vinculo_ofreq_precatorio_de_texto(texto) -> (ofreq, precatorio) | None`

Rodam os mesmos `REGEX_OFREQ` / `REGEX_VINCULO`. As funções `extrair_numero_ofreq(pdf_bytes)`
e `extrair_vinculo_ofreq_precatorio(pdf_bytes)` passam a **delegar** para elas após
`_extrair_texto_pdf(pdf_bytes)` — sem mudar a assinatura pública nem quebrar testes existentes.
Os classificadores `eh_documento_*` já operam sobre texto (sem mudança).

A extração de beneficiário/advogado (`extrair_beneficiario_completo`, `extrair_advogado`)
**continua exigindo o PDF baixado** (precisão), e isso só acontece para requisitórios confirmados.

### D. Spike de de-risco (1º passo da implementação)

O seletor do contêiner de texto da prévia no DOM do visualizador é **desconhecido**.

Passo 1 da implementação = ao vivo, identificar o seletor e confirmar que a prévia é
**texto extraível** (não canvas/imagem). Critério de go/no-go:

- **Prévia é texto** → segue a Abordagem 2 plena.
- **Prévia é imagem/canvas** → classificação-por-prévia não rende; **aborta a Abordagem 2**
  e mantém a Abordagem 1 ajustada (já em produção, commit d0790f0). Sem retrabalho cego.

Mesmo no caminho "go", o fallback por download (seção B, linha "prévia vazia/imagem") cobre
nós individuais que renderizem como imagem.

### E. Testes

**Unit (TDD):**
- `extrair_numero_ofreq_de_texto` e `extrair_vinculo_ofreq_precatorio_de_texto` (texto real
  transcrito das screenshots: `Definitivo OFÍCIO Nº: 2024.16127/OFREQ`, `Ofício 2025.10413/OFREQ
  ... gerou o precatório 2025.16841-6`); e que os wrappers `..._bytes` continuam funcionando.
- Decisão por nó: vínculo → não baixa; requisitório → baixa; prévia vazia → fallback download.
  (mock do leitor de prévia e do downloader, no estilo dos testes de `_baixar_e_classificar_indices`.)

**Ao vivo:**
- `0090006-26.2015` (Def Roberto, 1 alvo `2025.13367-1`): acha o requisitório + vínculo,
  baixa só o requisitório, mede nós-abertos vs nós-baixados.
- Um multi-precatório (ex: `0110128-60.2015`, Carraro, 4 alvos): acha os 4, baixa só os 4
  requisitórios, goal-stop dispara.

## Decisões abertas deixadas para o plano

- **Alcance/teto do fim**: como cada nó ficou barato, o teto pode subir bastante (ou virar
  "todas as folhas até goal-stop"); definir valor no plano após medir no spike.
- **Passo 1 por rótulo**: avaliar se ainda vale como atalho (quando o rótulo é bom, resolve
  sem varrer) ou se a passada única por prévia o torna redundante. Decidir no plano.

## Não-objetivos (YAGNI)

- OCR de PDFs escaneados (segue indo pra `manual_revisar/`, decisão prévia do usuário).
- Abordagem 3 (salto por data/índice) — descartada por fragilidade.
- Etapa ejud / 2ª instância (`.0000`) — projeto separado.
