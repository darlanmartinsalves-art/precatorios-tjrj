# Como rodar o pipeline TJRJ (Etapa 2-Unificada)

## Visão geral
Para cada precatório com saldo ≥ R$ 200.000:
1. Navega no portal TJRJ (consulta processual → detalhes → visualizador)
2. Identifica a peça com o ofício requisitório (heurística "Definitivo *")
3. Baixa o PDF e renomeia como `{precatório} - {beneficiário}.pdf`
4. Extrai do PDF: Beneficiário (Nome + CPF/CNPJ) e Advogado (Nome + CPF + OAB)
5. Atualiza a planilha em `Downloads/Precatórios 2027 - Atualizado - Etapa2.xlsx` (colunas N–R)

Pode pausar e retomar a qualquer momento — o checkpoint `downloads_state.json` rastreia tudo.

---

## Passo 1 — Abrir Edge com debug port

Em um PowerShell, **rode primeiro**:

```powershell
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9223 --user-data-dir=C:\temp\edge_debug
```

OU use o atalho que já existe:

```powershell
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& .\iniciar_edge_debug.ps1
```

> ⚠️ Esse comando trava o PowerShell (Chrome continua aberto). Use **outra janela do PowerShell** pro próximo passo.

---

## Passo 2 — Fazer login no Edge aberto

Na janela do Edge que abriu:

1. Vai em `https://www.tjrj.jus.br/`
2. Clica em "Advogado" no menu
3. Clica em "Processo Eletrônico" → "Portal de Serviços"
4. Faz login (CPF + senha + código verificador)
5. No modal "Alterar Perfil", seleciona **"Advogado"** e clica Entrar
6. **(Opcional)** Vá também ao confirmeonline e faça login lá: `https://confirme30.confirmeonline.com.br/search`

**Importante:** mantenha o Edge aberto enquanto o script rodar.

---

## Passo 3 — Rodar o script

Em **outra janela** do PowerShell:

```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
$env:PYTHONIOENCODING = "utf-8"
cd "C:\Users\DARLANMARTINS\Documents\PROJETO 01"
& "C:\Users\DARLANMARTINS\AppData\Local\Programs\Python\Python312\python.exe" baixar_requisitorios.py
```

Vai processar todos os ~1.952 processos pendentes. **Pode levar 15-25 horas.**

### Variantes úteis

```powershell
# Resetar tudo e começar do zero
... baixar_requisitorios.py --reset

# Só testar com 5 processos
... baixar_requisitorios.py --limit 5

# Re-tentar só os que deram erro
... baixar_requisitorios.py --apenas-erros

# Testar 1 processo específico
... baixar_requisitorios.py --test-processar "0027542-97.2014.8.19.0001"

# Filtrar saldo maior
... baixar_requisitorios.py --saldo-minimo 500000

# Não fechar PowerShell se sessão TJRJ expirar
#   Edge precisa estar logado o tempo todo
```

---

## Como saber se está rodando bem

O console vai mostrar uma barra de progresso:

```
[██████░░░░░░░░░░░░░░░░░░░░░░░░] 350/1952 (17.9%)  ETA 4min23s
```

**Tempos esperados por processo:**
- ~30s pra processos rápidos
- ~60-90s pra processos com modal Solicitação de Acesso
- ~5s pra processos PJe (pula)

**Arquivos gerados em tempo real:**
- `C:\Users\DARLANMARTINS\Downloads\Precatórios_Requisitórios\*.pdf` — PDFs baixados
- `C:\Users\DARLANMARTINS\Downloads\Precatórios 2027 - Atualizado - Etapa2.xlsx` — planilha com colunas N-R
- `C:\Users\DARLANMARTINS\Documents\PROJETO 01\downloads_state.json` — checkpoint (a cada 10 processos)
- `C:\Users\DARLANMARTINS\Documents\PROJETO 01\erros_download.csv` — log de erros (no final)

---

## Solução de problemas

### Edge precisa ser logado de novo
A sessão TJRJ expira após algum tempo de inatividade. Se vir muitos `processo_nao_encontrado` em sequência:
1. **Pare o script** (Ctrl+C no PowerShell)
2. Vá no Edge, faça login novamente (passo 2)
3. **Continue de onde parou:** `python baixar_requisitorios.py` (sem `--reset`)

### Antivirus bloqueia
Se o Edge não abrir com debug port (geralmente porta 9223 não responde), o antivirus pode estar bloqueando. Geralmente o Edge passa pelos antivirus corporativos (diferente do Chrome).

### Script trava
Pode acontecer se aparecer modal não tratado:
1. Ctrl+C pra parar o script
2. Tira print do que aparece na tela
3. Re-execute (`python baixar_requisitorios.py` sem reset, vai retomar)

---

## Etapa 4 (próxima — telefones via confirmeonline)

Ainda não implementada. Depende da Etapa 2 estar completa (com colunas O e Q da planilha preenchidas com os CPFs).
