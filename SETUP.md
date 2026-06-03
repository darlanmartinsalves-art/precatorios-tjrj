# SETUP — Rodar a automação de precatórios TJRJ numa máquina nova

Guia passo a passo para instalar e executar o pipeline de download de requisitórios
do TJRJ em qualquer computador Windows.

---

## 1. Pré-requisitos (instalar uma vez)

| Item | Como obter |
|------|------------|
| **Git** | https://git-scm.com/download/win (vem com o gerenciador de credenciais) |
| **Python 3.12** | https://www.python.org/downloads/ — marque **"Add Python to PATH"** na instalação |
| **Microsoft Edge** | Já vem no Windows (o script usa o Edge do sistema) |

Confirme no PowerShell:
```powershell
git --version
python --version    # deve mostrar Python 3.12.x
```

---

## 2. Baixar o código

```powershell
git clone https://github.com/darlanmartinsalves-art/precatorios-tjrj.git
cd precatorios-tjrj
```

## 3. Instalar as dependências Python

```powershell
pip install -r requirements.txt
```
(Instala: openpyxl, requests, playwright, pypdf, pytest. Não precisa de
`playwright install` — o script usa o Edge já instalado no sistema.)

## 4. Colocar a planilha de entrada

Copie **`Precatórios 2027 - Atualizado.xlsx`** para a pasta **Downloads** do usuário
(`C:\Users\<voce>\Downloads\`). É a planilha com os precatórios a consultar.

---

## 5. Rodar

```powershell
$env:PYTHONIOENCODING = "utf-8"
python baixar_requisitorios.py
```

1. O **Edge abre sozinho**. Faça login: **Processo Eletrônico → Portal de Serviços →
   CPF/senha/código verificador → perfil "Advogado"**.
2. **Logue na janela que o script abriu** (não na sua janela normal do Edge) e
   **mantenha o Edge aberto** durante todo o processamento.
3. O script detecta o login sozinho e começa. Salva progresso a cada 10 processos.

**Saídas geradas:**
- PDFs: `Downloads\Precatórios_Requisitórios\` (nomeados `{precatório} - {beneficiário}.pdf`)
- Planilha preenchida: `Downloads\Precatórios 2027 - Atualizado - Etapa2.xlsx`
- Checkpoint (retomável): `downloads_state.json` (na pasta do projeto)
- PDFs escaneados (sem texto) p/ conferência manual: `manual_revisar\`

---

## 6. A sessão do TJRJ cai? (normal)

O TJRJ derruba a sessão de tempos em tempos. Quando isso acontece, o script para
sozinho (circuit breaker) e **salva tudo**. Para retomar:

```powershell
python baixar_requisitorios.py        # SEM --reset — continua de onde parou
```
Relogue no Edge e ele segue. Nada é perdido (o checkpoint guarda o progresso).

## 7. No FINAL do lote (passo único)

Os processos que falharam por queda de sessão (`erro_navegacao`) não são refeitos
nos resumes normais. Quando o lote terminar, rode **uma vez**:

```powershell
python baixar_requisitorios.py --apenas-erros
```
Isso recupera esses processos (quase todos são falsos erros de sessão e dão certo no
retry).

---

## 8. Continuar o MESMO lote em outra máquina

O checkpoint (`downloads_state.json`) **não** vai pro GitHub (contém dados). Para
continuar o mesmo lote noutro PC, **copie manualmente** o `downloads_state.json` (e,
se quiser, a pasta `Precatórios_Requisitórios` com os PDFs) para a pasta do projeto.
Sem o checkpoint, use `--reset` para começar do zero.

## 9. Variantes úteis

```powershell
python baixar_requisitorios.py --reset            # apaga o checkpoint e refaz tudo
python baixar_requisitorios.py --limit 10         # processa só 10 (teste)
python baixar_requisitorios.py --apenas-erros     # refaz só os que não ficaram "ok"
python baixar_requisitorios.py --test-processar "0110128-60.2015.8.19.0001"  # 1 processo, com log
```

## 10. Atualizar o código

```powershell
git pull        # baixa as atualizações mais recentes
```

---

## Rodar os testes (opcional, valida o código)

```powershell
python -m pytest -q
```
