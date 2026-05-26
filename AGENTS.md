# AGENTS.md — briefing para agentes (Codex CLI, Claude Code, etc.)

Este projeto **opera ofícios SSG no e-TCM** (Tribunal de Contas do Município de
São Paulo) para a fila APO-PEN (aposentadoria / pensão). Não é uma webapp: é uma
automação Playwright que dirige o navegador real do operador, gera DOCX, anexa,
conclui o ato e pede assinatura. **Tudo é PROD por padrão** (sem ambiente de
homologação efetivo para esses processos — ver `MEMORY.md → etcm-ambiente`).

Para setup do venv/Playwright veja [README.md](README.md). Este doc cobre o que
um agente precisa saber para **não pedir o básico de novo**.

---

## 1. Identidade & escopo

- e-TCM: `https://etcm.tcm.sp.gov.br/`. Login automático via `ETCM_USERNAME` /
  `ETCM_PASSWORD` (env vars de USER-scope, já configuradas na máquina do operador).
- Operador padrão / assinante: **Roseli Chaves** (`ASSINANTE_NOME=Roseli Chaves`).
  Cleanup só apaga **rascunhos dela** (`SAFE_DELETE_OWN_DRAFTS=true`).
- Plataforma: **Windows 10/11 + PowerShell + Playwright HEADFUL** (browser visível
  na tela). Não roda em sandbox cloud sem desktop.

## 2. Pipeline (o que `src/main.py` faz por processo)

```
   APO-PEN grid
        ↓
1. Cleanup destrutivo (apaga comunicação + Ofício SSG antigos do robô)
        ↓
2. Cria Comunicação Processual (descrição = "dilação" | "conhecimento/providências" | ...)
        ↓
3. Gera DOCX a partir do template, anexa VINCULADO à comunicação
        ↓
4. Conclui o Ato Ofício SSG
        ↓
5. Solicita assinatura (Roseli Chaves)
        ↓
6. Tramita (atualmente DESLIGADO via SKIP_TRAMITACAO=true)
```

Cada etapa loga `"<verbo concluído> para <processo>"`. O marcador final de
sucesso é `"Assinatura solicitada para Roseli Chaves no processo TC/..."`.

## 3. 4 tipos × 3 secretarias

| Tipo (canônico) | Descrição da comunicação | Pasta de template |
|---|---|---|
| `UTAP`        | conhecimento/providências   | `modelos_utap/`        |
| `DILACAO`     | dilação                     | `modelos_dilacao/`     |
| `REITERACAO`  | reiteração                  | `modelos_reiteracao/`  |
| `JUIZO`       | decisão de juízo singular   | `modelos_juizo/`       |

Cada pasta tem 3 arquivos: `... - Saúde.{docx,dotx}`, `... - Educação.*`,
`... - Geral.*`. Para processos fora de Saúde/Educação (ex.: Urbanismo) o
template usado é o `Geral`, **mas o destinatário da comunicação é a UG real**
(ver gotcha §6.3).

Classificador automático: `_classify_tipo_from_text_and_piece` em
[src/main.py](src/main.py). Override por env: `FORCE_TIPO=UTAP|DILACAO|REITERACAO|JUIZO`.

## 4. Convenção de scripts (`run_*.ps1`)

- **Batch:** `run_<N>_processos_<LOTE>_<YYYY>_<MM>_<DD>.ps1` (ex.: `run_32_processos_DILACAO_2026_05_25.ps1`).
  Orquestra `pytest tests -q` → `main.py` em lote → retry automático até 3x →
  relatório `logs/RELATORIO_*.txt`.
- **Retry isolado:** `run_retry_<TC>_<YYYY>.ps1`. Mesmo flow para 1 processo só,
  usado quando um caiu no batch (ver §6.5).
- Todo script começa carregando credenciais de User-env e abortando se ausentes.
- Todo script faz `Set-Location -Path $PSScriptRoot` (importante quando lançado
  pelo Task Scheduler).
- `FORCE_TIPO` é setado **APÓS** o pytest, senão vaza override para os testes
  de classificação.

## 5. Env vars que importam (não decorar — usar como referência)

Estruturais (sempre):
```
ENVIRONMENT=producao   HEADLESS=false   SHOW_BROWSER=true
WATCH_MODE=true   SLOWMO_MS=200
ETCM_USERNAME / ETCM_PASSWORD   (User-scope)
ETCM_URL=https://etcm.tcm.sp.gov.br/paginas/login.aspx
```

Autoriza PROD destrutivo (sem isso, cleanup é no-op):
```
RUN_PROD_DESTRUCTIVE_CLEANUP=true
SAFE_DELETE_OWN_DRAFTS=true
USE_CAIXA_CORREIO=true
FORCE_DELETE_OLD_OFICIO_SSG=true
FORCE_RECREATE_COMUNICACAO=true
FORCE_REVOKE_PENDING_ROSELI_SIGNATURE=true
ONLY_PROCESSOS_AUTHORIZED=TC/...,TC/...   (lista explícita, gating de segurança)
```

Comunicação processual:
```
COMUNICACAO_PRAZO_DIAS=60
COMUNICACAO_REFERENCIA="gerado automaticamente"   (marcador p/ cleanup futuro)
STATUS_ENTREGA=Normal
```

Ofício:
```
OFICIO_TEMPLATE_MODE=auto
OFICIO_PRESERVE_AT_TOKENS=true                    (run-preservante; obrigatório)
OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA=true     (UTAP) | false (DILACAO)
OFICIO_ENCAMINHA_FONT_NAME="Times New Roman"
OFICIO_ENCAMINHA_FONT_SIZE_PT=12
DATA_OFICIO=DD/MM/YYYY   OFICIO_DATA=DD/MM/YYYY
FORCE_TIPO=DILACAO   (só para batches dilação — setar APÓS pytest)
```

Assinatura + tramitação (preferência atual):
```
REQUEST_SIGNATURE=true   ASSINANTE_NOME="Roseli Chaves"
SKIP_TRAMITACAO=true     TRAMITAR_DESTINO=""
```

## 6. Gotchas (cada um já queimou alguém — leia)

### 6.1 Pasta com acento quebra encoding em bash/cmd
A raiz é `C:\Users\20386\OneDrive - tcm.sp.gov.br\Ofícios\oficio_automation\`.
O "í" em "Ofícios" quebra ao passar por `cmd.exe` ou inline-arg de `powershell.exe`
via bash. **Use o short-path 8.3** quando precisar de caminho ASCII puro
(Task Scheduler, args inline):

```
C:\Users\20386\ONEDRI~1.BR\OFCIOS~1\OFICIO~1\
```

Resolva em runtime com:
```powershell
$fso = New-Object -ComObject Scripting.FileSystemObject
$fso.GetFile($absPath).ShortPath
```

### 6.2 Tokens `@@` nos modelos DOCX estão quebrados em múltiplos runs
Os modelos têm `@@processo`, `@@nome_relator`, etc. No XML, esses tokens estão
**partidos** em vários `<w:r><w:t>` (Word splita o run quando o operador edita).
A automação só funciona com substituição "run-preservante"
(`OFICIO_PRESERVE_AT_TOKENS=true`). **Nunca retipe o corpo de um modelo à mão** —
você quebra os tokens. Para rebrandar, transplante header/footer; preserve o
body byte-a-byte.

Tokens canônicos (validação `_validate_oficio_at_tokens`):
- UTAP/Dilação/Reiteração: `@@numero_oficio @@natureza_processo @@processo @@Tipo_Processo @@Nome_interessado @@processoexterno @@nome_relator @@instancia @@data_extenso`
- Juízo (acima +): `@@data_publicacao_ata @@pag_publicacao_ata`

### 6.3 Destinatário da comunicação = Unidade Gestora REAL, não o rótulo do bucket
O campo `cbbUsuarios` ("Destinatário") tem ~5 opções fixas. Use o texto de
`ucTabInfo_pgcInfoProc_lblUG` da página do processo (ex.: "Secretaria Municipal
de Urbanismo e Licenciamento (*)"). Passar "Geral" (rótulo do template) dá
no-match → save rejeitado em silêncio → grid não avança → popup não fecha.
Já corrigido no `criar_comunicacao_processual` (commit `42dbfd3`).

### 6.4 Dilação em batch — Referência e Ofício SSG são por processo
A automação auto-extrai por processo:
- **Referência (cabeçalho):** 1ª linha "Ofício nº .../..." da última peça REQUERIMENTO.
- **Nº do Ofício SSG (corpo):** número no NOME da peça logo após o último MANUTAP-OF.

**Em batch, NÃO setar** `OFICIO_REFERENCIA_TEXT` nem `OFICIO_SSG_REF` globalmente —
todos os processos receberiam o mesmo valor (bug grave). Esses env vars existem
apenas como fallback em retry de **1 processo só**.

### 6.4b Reiteração — auto-extração do nome das peças (Saúde e Educação)
Briefing Gilson Nóbrega 2026-05-25. Quando `FORCE_TIPO=REITERACAO`, o robô
auto-extrai TUDO do **nome das peças** (não baixa PDFs extras):

- **Referência (cabeçalho):** `Ofício SSG <num/ano>, encaminhado eletronicamente em <DD/MM/AAAA>.`
  — `<num/ano>` e data vêm da peça SSG **logo após o último MANUTAP-OF** (o ofício
  que está sendo reiterado).
- **Encaminha:** `Cópia das peças <MANUTAP> e <DES> dos autos.` — números (de
  display, não index_ato) da última peça **MANUTAP-OF** e da última peça **DES**
  (despacho do conselheiro pedindo reiteração).
- **Corpo do ofício (SSG):** mesmo `<num/ano>` da Referência.

Implementado em `_extract_reiteracao_data_from_pieces(pieces)` em [main.py](src/main.py).
A função preenche três env vars que o pipeline consome adiante:
- `OFICIO_SSG_REF`
- `OFICIO_REFERENCIA_TEXT`
- `OFICIO_ENCAMINHA_PIECE_NUMBERS` (lista CSV, ex.: "4,9") — consumida pelo bloco
  de Encaminha que chama `format_encaminha_from_piece_numbers` e propaga para
  `OFICIO_ENCAMINHA_TEXT` (para o de-bolding kick in via
  `set_encaminha_text_without_bold`).

Em batch, valem as MESMAS regras da dilação: **não** setar manualmente nenhum
desses três env vars (todos os processos receberiam o mesmo valor). Para
**run isolado**, é OK setar como fallback se a auto-extração falhar — mas o
padrão é deixar o `_extract_reiteracao_data_from_pieces` cuidar de tudo.

Setando `REITERACAO_REQUIRE_AUTO_EXTRACT=true` força aborto cedo se algum
campo crítico (SSG, Referência, peças MANUTAP+DES) não foi auto-extraído —
melhor que gerar ofício incompleto.

### 6.5 DevExpress da Comunicação tem bug intermitente (várias manifestações)
O save do form às vezes não confirma na grid em batch (popup não fecha). Não é
determinístico em forma única — observamos pelo menos 3 manifestações:

1. **Confirm timeout puro**: popup demora a fechar/grid demora a atualizar.
   Mitigado em 2026-05-25: timeout subiu de 45s → **120s**.
2. **`destinatario-fix` items_count=0** (dropdown vazio): o callback que
   popula `cbbUsuarios` não termina dentro da janela de espera, e a auto-match
   recebe lista vazia. Mitigado em 2026-05-25: o `time.sleep(2)` virou um
   **wait-loop** que enquete `GetItemCount` a cada 300ms até items > 0 ou 10s.
3. **Falha persistente em alguns processos** (ex.: TC/008918/2022, TC/011771/2022):
   reproduzível ≥4 tentativas seguidas. Hipótese aberta — possivelmente
   `rows_before≥7` interage mal com algo. Sem fix automático até agora; nesses
   casos o operador completa a comm manualmente.

Distinto do "no-match" de §6.3 (esse é determinístico — `destinatário='Geral'`
ou alvo inválido no dropdown).

Heurística:
- `no-match` com `items_count > 0` → bug de dados (§6.3).
- `no-match` com `items_count == 0` → race do callback (já mitigado pelo wait-loop).
- timeout sem `no-match` → bug do form (varies); retry isolado às vezes resolve.

### 6.6 Cleanup destrutivo apaga ANTES de criar
Se a criação da comm/ofício falhar no meio, o processo fica **incompleto**
(antigos apagados, novos não criados). Único caminho de recuperação é
re-rodar (após corrigir a causa). Sempre cheque o log para identificar em
qual passo travou antes de re-rodar.

### 6.7 PDF é a fonte de verdade da lista de processos
A planilha `Processos_*.pdf` é exportada do e-TCM. **Coluna "N° Processo"
(página 2)** é a única fonte autoritativa — não a primeira página (que tem
agrupamentos visuais). Veja `MEMORY.md → reference-planilha-processos`.

### 6.8 Scripts `.ps1` DEVEM ser 100% ASCII (sem acentos, sem em-dash)
PowerShell 5.1 (`powershell.exe`, o que o Task Scheduler usa) lê arquivos
`.ps1` SEM BOM com o codepage do sistema — Windows-1252 em pt-BR. Caracteres
multi-byte UTF-8 (acentos, em-dash `—` U+2014, aspas tipográficas) decodificam
para bytes que incluem U+201D (`"` aspas direitas) e outros chars que o lexer
trata como delimitador de string. Resultado: parse error 7+ linhas depois do
caractere ofensor, mesmo se ele estiver dentro de uma string ou comentário.

**Sintoma operacional:** task agendada dispara mas `powershell.exe` sai em
**<1s com exit code 2147942401 (=0x80070001)** sem escrever nenhum log. Event
Viewer (Microsoft-Windows-TaskScheduler/Operational) registra o `código de
retorno 2147942401` na action. Aconteceu com o `run_LOTE_2026_05_26_*.ps1`
em 26/05/2026: em-dash dentro de `Log-Both "WRAPPER ... — INICIO ..."` quebrou
o parse; nenhum dos 13 processos do lote foi tocado.

**Regra:** todo `.ps1` neste repo é escrito em ASCII puro — `dilacao` (não
`dilação`), `--` (não `—`), `Saude` (não `Saúde`). Strings que serão impressas
em logs/relatórios podem ter acentos APENAS se gravadas com BOM UTF-8 explícito
(raro; geralmente ASCII serve).

**Diagnóstico rápido** (rodar quando uma task agendada sair em <1s):

```powershell
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    "<caminho.ps1>", [ref]$null, [ref]$errors
)
$errors | ForEach-Object {
    "  L{0}:{1}  {2}" -f $_.Extent.StartLineNumber, $_.Extent.StartColumnNumber, $_.Message
}
```

Se aparecer erro de "string sem terminador" ou "token X inesperado" em uma
linha que parece OK, suspeitar de não-ASCII em alguma linha acima. Caçar com:

```powershell
Select-String -Path "<caminho.ps1>" -Pattern '[^\x00-\x7F]' -AllMatches
```

## 7. Agendamento (Windows Task Scheduler, não scheduler remoto)

O fluxo é headful local com credenciais locais — **scheduler remoto na nuvem
não serve**. Convenção:

- Nome: `eTCM_<TIPO>_<N>_<YYYY>_<MM>_<DD>` (ex.: `eTCM_DILACAO_32_2026_05_25`).
- `LogonType=Interactive`, `RunLevel=Limited` (headful precisa de sessão
  logada; sem senha armazenada).
- **`StartWhenAvailable=False` de propósito**: se o PC estiver desligado às
  00:10, a janela é perdida — **não** roda em horário aleatório depois
  (mais seguro para fluxo destrutivo).
- `WakeToRun=True` para acordar de sleep leve.
- `ExecutionTimeLimit` generoso (8–10h para batches grandes).

Registrar via `Register-ScheduledTask` (cmdlet, não `schtasks.exe`); cancelar com
`Unregister-ScheduledTask -TaskName <nome> -Confirm:$false`.

Mais detalhes: `MEMORY.md → project-agendamento-execucoes`.

## 8. Smoke test antes de qualquer batch

```powershell
.\.venv\Scripts\python -m pytest tests -q          # sempre verde antes
.\.venv\Scripts\python -m py_compile src/main.py   # rápido, pega syntax error
```

Os scripts `run_*.ps1` já chamam `pytest` no início e abortam se falhar.

## 9. Estrutura de arquivos

```
src/                           # código (main.py é o entrypoint do pipeline)
tests/                         # pytest
modelos_utap/                  # templates por tipo × secretaria
modelos_dilacao/
modelos_reiteracao/
modelos_juizo/
output/                        # DOCX/PDF gerados (volumoso, não commitar)
logs/                          # logs por execução + RELATORIO_*.txt
artifacts/evidence/            # HTML salvo quando o save da comm falha
run_<N>_processos_*.ps1        # batches (na raiz)
run_retry_<TC>_*.ps1           # retries isolados
storage_state.json             # sessão Playwright (pode expirar; login auto rebcria)
README.md / AGENTS.md          # docs
```

## 10. Autonomia & segurança (regras do operador)

- O operador autoriza execução autônoma para este fluxo (não parar pedindo
  permissão a cada passo).
- **Mas** continuar fazendo diligência: antes de criar tarefa nova, listar as
  existentes para evitar duplicar disparo destrutivo (já aconteceu).
- **Nunca** disparar PROD destrutivo por suposição. Em ambiguidade real
  (ex.: "rode agora" vs. agendamento já solicitado), perguntar.
- Antes de commitar mudanças, verificar `git status`: o `src/main.py` costuma
  ter WIP acumulado da branch; mencionar explicitamente o que está sendo
  empacotado junto.
- Push só quando o operador pedir.

## 11. Quando perguntar (não decidir sozinho)

- Instruções ambíguas onde uma interpretação é destrutiva/irreversível.
- Primeira ocorrência de secretaria nova ou tipo novo (a lógica de
  destinatário/template pode não cobrir — vide caso Urbanismo).
- Decisão de commit que mistura fix com WIP anterior — quem é o autor sabe
  se aquilo deve subir junto.
- Pedido de tramitar (`SKIP_TRAMITACAO=false`) — hoje está desligado por
  preferência operacional; ligar sem confirmação muda o que sai do tribunal.

---

Memória de projeto adicional fica em `~/.claude/projects/<slug>/memory/MEMORY.md`
(Claude Code). Codex CLI pode espelhar o conteúdo em `.codex/` ou ler este arquivo.
