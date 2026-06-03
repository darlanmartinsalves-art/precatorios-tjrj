# Atualização automática de saldos de precatórios TJRJ

**Data:** 2026-05-27
**Autor:** Darlan Martins + Claude
**Status:** Design aprovado

## Contexto e motivação

Há uma planilha `Precatórios 2027.xlsx` (6.288 linhas) cujo campo "Saldo atualizado" (coluna H) precisa ser preenchido a partir da consulta individual no portal do TJRJ. A consulta manual é inviável; é necessário automatizar.

Trabalho **recorrente** — será executado periodicamente (mensal/trimestral) para atualizar os saldos.

## Descoberta importante: API REST do TJRJ

A página `https://www.tjrj.jus.br/web/precatorios` usa internamente um endpoint REST não documentado publicamente:

```
GET https://www3.tjrj.jus.br/PortalConhecimento/api/precatorios/consultaPosicao
    ?numeroPrecatorio={numero}
```

Retorno (JSON):
```json
{
  "Posicao": 34118.0,
  "Saldo": 13613821.78,
  "ColocacaoAtual": 202703553,
  "AnoAtual": 2027,
  "AnoPagamento": 2021,
  "Natureza": "Comum",
  "Origem": "Estado Do Rio De Janeiro",
  "Quitado": false,
  "Cancelado": false,
  "SaldoFinal": 13613821.78
}
```

Campo de interesse: `Saldo` (decimal em reais).

## Origem dos dados (importante)

O número de precatório de cada linha **vem da própria planilha de entrada (coluna B)**. O usuário **não precisa informar nada manualmente**. Para cada linha de dados:

1. Lê coluna B → número do precatório (ex.: `2025.17049-6`)
2. Consulta a API do TJRJ com esse número
3. Grava o saldo retornado na coluna H **da mesma linha**

Exemplo:
```
Entrada (linha 2):
  B2 = "2025.09451-0"  →  consulta API  →  retorna Saldo: 59308.24
Saída (linha 2):
  B2 = "2025.09451-0"  (inalterada)
  H2 = 59308.24        (preenchida com formato monetário)
```

## Não-objetivos

- Não atualizar outras colunas além de H (saldo atualizado).
- Não criar interface gráfica.
- Não publicar ou compartilhar a planilha.
- Não recalcular o saldo localmente — sempre consultar o TJRJ.
- Não modificar a planilha de entrada (output é arquivo novo).
- Não solicitar números de precatório ao usuário — a planilha é a fonte.

## Arquitetura

### Stack
- **Python 3** (instalado via Microsoft Store)
- **openpyxl** — manipulação de xlsx preservando formatação
- **requests** — chamadas HTTP
- **concurrent.futures.ThreadPoolExecutor** — paralelismo (biblioteca padrão)

### Layout de arquivos
```
C:\Users\DARLANMARTINS\Documents\PROJETO 01\
  atualizar_saldos.py              # script principal
  requirements.txt                 # openpyxl>=3.1, requests>=2.31
  saldos_checkpoint.json           # gerado automaticamente
  erros_consulta.csv               # gerado se houver falhas
  execucao.log                     # log detalhado de cada execução

C:\Users\DARLANMARTINS\Downloads\
  Precatórios 2027.xlsx            # entrada (read-only)
  Precatórios 2027 - Atualizado.xlsx  # saída (novo arquivo)
```

### Princípio de invariantes
1. O arquivo de entrada **nunca** é modificado.
2. O checkpoint é a única fonte de verdade durante a execução.
3. Operações de I/O para o checkpoint são **atômicas** (escreve `.tmp` → `rename`).
4. Saldo só é gravado na planilha se for `int` ou `float` válido.

## Fluxo de dados

### Fase 1 — Carregar estado
```
1. Verifica que entrada existe; se não, aborta.
2. Abre xlsx em modo read_only (rápido).
3. Itera linhas 2..N, extrai (linha_excel, numero_precatorio) da coluna B.
4. Filtra apenas números que casam com /^\d{4}\.\d+-\d$/.
5. Se checkpoint.json existir: carrega dict {numero: resultado}.
6. Calcula lista de pendentes = todos − (consultados com sucesso).
```

### Fase 2 — Consultar API em paralelo
```
ThreadPoolExecutor(max_workers=20):
  para cada numero pendente:
    futuro = pool.submit(consultar_saldo, numero)
  
  para cada futuro concluído (as_completed):
    resultado = futuro.result()
    checkpoint[numero] = resultado
    
    se (concluidos % 100) == 0:
      escrever checkpoint atomicamente
      atualizar barra de progresso

consultar_saldo(numero):
  para tentativa em [1, 2, 3]:
    try:
      resp = GET endpoint, timeout=30s
      if status == 200: return json["Saldo"]  # pode ser float ou null
      if status == 404: return "NAO_ENCONTRADO"
      else: raise para retry
    except (timeout, conn_error, 5xx):
      sleep(2 ** (tentativa-1))  # 1s, 2s, 4s
  return "ERRO_REDE"
```

### Fase 3 — Escrever planilha
```
1. wb = load_workbook(entrada)  # carrega normalmente (não read_only)
2. ws = wb.active
3. Aplicar number_format='R$ #,##0.00' à coluna H das linhas de dados
4. Para cada (linha, numero) extraído:
     resultado = checkpoint.get(numero)
     se resultado é numérico:
       ws.cell(row=linha, column=8).value = resultado
5. wb.save(saida)
```

### Fase 4 — Relatório
```
Console:
  Consultados com sucesso: X
  Sem saldo (quitado/cancelado): Y
  Erros de rede: Z
  Não encontrados: W
  Tempo total: HH:MM:SS

Se Z+W+falhas > 0: gera erros_consulta.csv com (linha, numero, motivo)
```

## Tratamento de erros

| Cenário | Comportamento | Anotação no checkpoint |
|---|---|---|
| Timeout / conn refused | Retry 3x backoff exponencial | `"ERRO_REDE"` (re-tenta na próxima execução) |
| HTTP 5xx | Retry 3x backoff exponencial | `"ERRO_REDE"` |
| HTTP 404 | Sem retry | `"NAO_ENCONTRADO"` |
| HTTP 200, `Saldo: null` | Sem retry | `"SEM_SALDO"` (provavelmente quitado/cancelado) |
| JSON inválido | Sem retry | `"RESPOSTA_INVALIDA"` |
| Sucesso | — | `<float>` (valor do saldo) |
| Ctrl+C / kill | Checkpoint preserva trabalho | — |

**Regra de re-consulta:** apenas entradas marcadas como `"ERRO_REDE"` são re-tentadas na próxima execução. Outras anotações (`SEM_SALDO`, `NAO_ENCONTRADO`, `RESPOSTA_INVALIDA`) são consideradas finais e não consomem tempo de rede em re-execuções.

A flag `--apenas-erros` força re-tentativa de todas as entradas não-numéricas (caso o usuário queira investigar).

## Interface de linha de comando

```
python atualizar_saldos.py [opções]

  --entrada PATH       Caminho do xlsx (padrão: ~/Downloads/Precatórios 2027.xlsx)
  --saida PATH         Caminho do xlsx de saída (padrão: <entrada> + " - Atualizado.xlsx")
  --workers N          Threads paralelas (padrão: 20)
  --limit N            Processa apenas N precatórios pendentes (testes)
  --reset              Apaga checkpoint antes de começar
  --apenas-erros       Re-consulta entradas marcadas como erro
  --so-aplicar         Pula consulta, só aplica checkpoint atual à planilha
```

## Validação antes do uso completo

1. **Sanity check (30s)**: `python atualizar_saldos.py --limit 10`
   - Verificar coluna H preenchida nas 10 primeiras linhas
   - Verificar formato monetário aplicado
   - Comparar 1-2 valores com consulta manual no site
2. **Teste de retomada (1min)**: rodar com `--limit 50`, Ctrl+C aos 25, rodar de novo. Confirmar que termina os 25 restantes.
3. **Execução completa**: rodar sem `--limit` (~8 minutos).

## Métricas de progresso (console)

```
Consultando precatórios:
  Pendentes: 6288  |  Concluídos: 2345 (37.3%)
  Taxa: 14.2/s  |  ETA: 4min38s
  Erros: 3  |  Workers: 20
  [████████████░░░░░░░░░░░░░░░░░░]
```

## Critérios de sucesso

- [ ] Script roda do início ao fim sem intervenção manual em < 15 minutos.
- [ ] Coluna H da planilha de saída preenchida em ≥ 99% das linhas.
- [ ] Comparação manual de 3 valores aleatórios bate com o site do TJRJ.
- [ ] Interromper com Ctrl+C e reiniciar continua do ponto correto.
- [ ] Planilha de entrada permanece intacta (mesmo hash MD5).
- [ ] Formatação original da planilha de saída preservada (fontes, larguras, abas).

## Out of scope (futuras melhorias)

- Suporte a múltiplas abas (a planilha tem 2 abas mas só a primeira tem precatórios).
- Notificação por e-mail quando termina.
- Agendamento via Task Scheduler.
- Detecção de mudanças (alertar quando o saldo de um precatório muda muito).
- Histórico de saldos ao longo do tempo.
