"""Seletores DOM do portal TJRJ — gerados pela exploração interativa (E2U Task 7).

Fluxo descoberto:
================

1. Página: https://www3.tjrj.jus.br/portalservicos/#/consproc/consultaportal
   - Esta é a página OUTER. O conteúdo real está em iframe interno
     com URL contendo "/consultaprocessual/".
   - Use `INNER_FRAME_URL_FRAGMENT` para identificar o frame interno.

2. Tela de Consulta Processual (no iframe interno):
   - Radio "Tipo de numeração": "Única" deve estar marcado (RADIO_UNICA)
   - 2 campos para o CNJ split em (prefix, sufix) — separados por ".8.19."
     visualmente. O .8.19. é fixo, não preencher.
   - Botão Pesquisar — há múltiplos (um por aba), pegar o VISÍVEL.

3. Após pesquisar, tabela "Resultado" aparece embaixo. Link clicável é o
   CNJ azul (anchor com `href="javascript:void(0)"`).
   - Clicar nele navega o SPA para a página de Detalhes do Processo
     (mesma aba, URL do iframe muda para /consultar/detalhes-processo...).

4. Página Detalhes do Processo:
   - Botão "Processo Eletrônico - Visualizador" no topo direito.
   - Clicar abre NOVA ABA com URL https://www3.tjrj.jus.br/visproc/#/<hash>
   - Se a usuária NÃO está vinculada aos autos, antes da nova aba aparece
     um MODAL pedindo senha provisória. Bypass: preencher qualquer texto
     no campo e clicar "Visualizar Processo".

5. Visualizador (nova aba): https://www3.tjrj.jus.br/visproc/#/...
   - Árvore de peças à esquerda (mat-nested-tree-node com role="treeitem")
   - Painel central renderiza o PDF (gerado dinamicamente pela peça selecionada)
   - Toolbar tem botão "Salvar Cópia do Documento em Exibição" — baixa
     APENAS a peça que está selecionada (não o processo todo).

6. Identificação da peça correta:
   - O REQUISITÓRIO geralmente está em peça nomeada "PREC DEFINITIVO AUTOR"
     (ou similar contendo "PREC DEFINITIVO").
   - Pode também haver "Definitivo OFÍCIO Nº: AAAA.NNNNN/OFREQ" como peça
     separada (de data posterior, é uma comunicação/anúncio do TJ).
   - Estratégia robusta: iterar candidatos (PREC DEFINITIVO, OFREQ, REQUISIT),
     clicar, baixar, e verificar com extrair_beneficiario_completo() — o
     que retornar dados válidos é o requisitório real.
"""

# ===== URLs =====
URL_CONSULTA = "https://www3.tjrj.jus.br/portalservicos/#/consproc/consultaportal"
URL_BASE = "https://www3.tjrj.jus.br/portalservicos/"

# Fragmento na URL do iframe interno (onde estão os elementos reais da consulta)
INNER_FRAME_URL_FRAGMENT = "consultaprocessual"

# Fragmento na URL do visualizador (nova aba)
VISUALIZADOR_URL_FRAGMENT = "visproc"


# ===== Tela de Consulta (iframe interno) =====
RADIO_UNICA = "#numeracaoUnica"
RADIO_ANTIGA = "#numeracaoAntiga"
CAMPO_CNJ_PREFIX = 'input[name="numeroProcesso"]'
CAMPO_CNJ_SUFIX = 'input.sufixo-codigo'
# Há múltiplos #botaoPesquisarProcesso (um por aba); usar `.locator(...).all()` e filtrar
# por `await b.is_visible()`.
BOTAO_PESQUISAR = "#botaoPesquisarProcesso"

# Link de resultado: anchor azul com texto contendo o CNJ
# Usar text=/{prefix.split('-')[0]}.*{sufix}/ para casar dinamicamente


# ===== Tela Detalhes do Processo =====
BOTAO_VISUALIZADOR = 'text=/Processo Eletr[oô]nico.*Visualizador/i'


# ===== Modal de Acesso ao Processo (aparece se não vinculada aos autos) =====
# Quando o advogado NÃO está vinculado ao processo, o portal exibe um modal
# "Visualização de documentos do processo eletrônico" com DUAS opções:
#
# 1. SENHA PROVISÓRIA (campo + botão "Visualizar Processo" btn-sm)
#    - Funciona digitando QUALQUER texto no campo e clicando o botão btn-sm.
#
# 2. SOLICITAÇÃO DE ACESSO (campo Motivo + botão "Visualizar Processo" float-right)
#    - Conforme Resolução 121 CNJ. Mais formal — gera registro de acesso.
#    - Funciona digitando QUALQUER texto no Motivo e clicando o botão float-right.
#
# Escolhemos a opção 2 (SOLICITAÇÃO DE ACESSO) por ser o caminho mais formal.

MODAL_ACESSO_MOTIVO = '#motivo'  # textarea Motivo (Solicitação de Acesso)
# Botão "Visualizar Processo" com class float-right (do bloco Solicitação de Acesso).
# Há mais um botão "Visualizar Processo" no mesmo modal (do bloco Senha Provisória),
# então o seletor PRECISA ser específico com float-right.
MODAL_ACESSO_BOTAO_OK = 'button.float-right:has-text("Visualizar Processo")'

# Texto padrão usado no Motivo (irrelevante o conteúdo)
MOTIVO_PADRAO = "Consulta de informacoes do precatorio"

# Selectors alternativos do bloco SENHA PROVISÓRIA (não usado — mantido como referência)
MODAL_SENHA_CAMPO = '#senhaProvisoria'
MODAL_SENHA_BOTAO_OK = 'button.btn-alinhado:has-text("Visualizar Processo")'


# ===== Visualizador (nova aba) =====
# Itens da árvore de peças
ARVORE_ITEM_TREEITEM = 'mat-nested-tree-node'
# Para encontrar peça por texto:  ARVORE_ITEM_TREEITEM + f':has-text("{padrao}")'

# Padrões para identificar a peça correta do requisitório
# Heurística descoberta empiricamente — peças costumam se chamar:
#   "PREC DEFINITIVO AUTOR"  (Garrastazu)
#   "PREC DEFINITIVO HONORÁRIOS SUCUMBENCIAIS"
#   "Definitivo <Nome do Beneficiário>"  (ex: "Definitivo Elisabeth Souza da Cunha")
# O padrão genérico que pega tudo é "Definitivo " no início do nome.
PADROES_PECA_REQUISITORIO = [
    "PREC DEFINITIVO",     # mais específico — Garrastazu, etc
    "Definitivo ",         # genérico — pega "Definitivo Nome do Benef" e variantes
    "OFÍCIO REQUISITÓRIO", # nome explícito (raro nos labels da árvore)
    "OFREQ",               # numeração do ofício
    "REQUISITÓRIO",        # fallback
    # Documento de VÍNCULO (liga OFREQ -> nº do precatório). Rótulo real na árvore:
    # "OFÍCIO DEPJU - PRECATÓRIO GERADO" (um por precatório gerado). Sem capturá-lo
    # não há como mapear o requisitório ao precatório certo.
    "PRECATÓRIO GERADO",   # rótulo do ofício DEPJU de vínculo
    "DEPJU",               # variações do rótulo do ofício de vínculo
]
# Limite de candidatos a tentar por processo (evita estourar tempo em processos
# com muitas peças "Definitivo"). Inclui nós-filhos (capa de juntada -> documento)
# e os ofícios de vínculo DEPJU. Processos multi-precatório (vários herdeiros +
# honorários) têm muitas peças, por isso o teto é folgado.
LIMITE_CANDIDATOS = 50

# Fallback por CONTEÚDO: alguns cartórios arquivam o requisitório e o ofício DEPJU
# como "Petição"/"Documento" genérico (sem "Definitivo"/"OFREQ"/"DEPJU" no rótulo da
# árvore) — ex: processo 0110128-60.2015. A busca por rótulo não casa nada e o
# processo vira "sem_requisitório" falso. Quando o passo por rótulo acha ZERO, fazemos
# um segundo passo baixando nós de documento genéricos e classificando pelo conteúdo
# (que já detecta OFREQ / "OFÍCIO REQUISITÓRIO" / "gerou o precatório").
# Stems com e sem acento, pra tolerar variação de rendering no DOM.
PADROES_DOC_GENERICO = ["Petiç", "Petic", "Documento", "Ofíci", "Ofici"]
# Teto do fallback — cada candidato custa ~5-8s de clique+download. Com o goal-stop
# (para assim que todos os precatórios-alvo do processo são resolvidos), no caso comum
# o loop para muito antes; este teto só limita o PIOR caso (alvo realmente ausente ou
# escaneado). Como o fallback varre do FIM da árvore (onde ficam os requisitórios/DEPJU),
# 15 alcança os alvos com folga e corta o desperdício pela metade vs. o antigo 30.
LIMITE_FALLBACK_GENERICO = 15

# Early-stop do fallback: no 0110128, os requisitórios/DEPJU formam um BLOCO contíguo no
# fim da árvore; depois deles só há petições antigas inúteis. Após ACHAR o bloco, se
# vierem N downloads irrelevantes seguidos, paramos (economiza ~2/3 dos downloads nos
# processos que acham cedo). Folga de 8 tolera um scan/petição no meio do bloco sem
# cortar cedo demais.
FALLBACK_EARLY_STOP_MISSES = 8

# Botão de download — pega apenas a peça em exibição
BOTAO_SALVAR_COPIA = 'button[aria-label="Salvar Cópia do Documento em Exibição"]'

# Botão alternativo (NÃO USAR — baixa processo inteiro): "Baixar o processo atual em PDF"


# ===== Modal "Mensagem Processo do PJe" (processo está no PJe, não acessível) =====
# Aparece quando o número CNJ pertence ao sistema PJe (separado do Portal de Serviços).
# Não temos acesso ao PJe — pular esses processos.
MODAL_PJE = '#modal-aviso-pje'
MODAL_PJE_BOTAO_FECHAR = '#modal-aviso-pje div.rodape-cancela'


# ===== Modal "Alterar Perfil" (aparece após login/re-autenticação) =====
# Modal: <div role="dialog" class="modalWide"> com título "Alterar Perfil"
# Tem app-dropdown id="dropdownPerfil" e botão Entrar.
# Sempre selecionamos "Advogado".

MODAL_ALTERAR_PERFIL = '[role="dialog"].modalWide:has-text("Alterar Perfil")'
MODAL_ALTERAR_PERFIL_DROPDOWN = '#dropdownPerfil input[role="combobox"]'
MODAL_ALTERAR_PERFIL_OPCAO_ADVOGADO = 'span.color-def:has-text("Advogado")'
MODAL_ALTERAR_PERFIL_BOTAO_ENTRAR = 'button:has-text("Entrar")'


# ===== Modal "Aviso de sessão inativa" (aparece a cada ~5 min de inatividade) =====
# Estrutura capturada do TJRJ Portal de Serviços:
#
#   <div class="modal fade show" id="modalUserIdleTimeout" role="dialog">
#     <h4 class="modal-title cabecalho-modal">Aviso de sessão inativa</h4>
#     <span>Sua sessão inativa terminará em 00:04:37. Deseja prolongar o tempo da sessão?</span>
#     <a><div class="rodape-confirma">[check] Prolongar sessão</div></a>
#     <a><div class="rodape-cancela">[ban]   Encerrar sessão</div></a>
#   </div>
#
# Aparece na aba do Portal de Serviços (consultaportal). Se ignorado por ~5min,
# desloga a sessão.
#
# Bot deve detectar e clicar "Prolongar sessão" automaticamente a cada N segundos.

MODAL_SESSAO_INATIVA = '#modalUserIdleTimeout'  # detecção
# IMPORTANTE: escopar ao modal. 'div.rodape-confirma' sozinho casa ~13 elementos na
# página (classe genérica de botão "confirmar") e o clique acerta o errado, deixando
# a sessão cair. Confirmado ao vivo: escopado casa exatamente 1 = "Prolongar sessão".
MODAL_SESSAO_PROLONGAR = '#modalUserIdleTimeout div.rodape-confirma'  # botão "Prolongar sessão"
MODAL_SESSAO_ENCERRAR = 'div.rodape-cancela'    # NÃO clicar — desloga

# Intervalo de polling: o modal tem timer de ~5min. Verificar a cada 60s
# garante que clicamos antes de expirar.
INTERVALO_POLLING_MODAL_SEG = 60


# ===== Indicador de login expirado =====
# Heurística: presença de campo password OU URL contém "login"
INDICADOR_LOGIN_EXPIRADO = 'input[type="password"]'


# ===== Comportamento das novas abas =====
ABRE_NOVA_ABA_APOS_PESQUISAR = False  # navega NO MESMO iframe (SPA)
ABRE_NOVA_ABA_VISUALIZADOR = True     # botão Visualizador abre NOVA ABA
