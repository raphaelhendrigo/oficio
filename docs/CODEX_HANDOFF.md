# Handoff Claude → Codex — Projeto Euclides (interface web)

> Documento de transição. Quando você (Codex) abrir este projeto, **comece lendo
> este arquivo**, depois `web/README.md`. Status atualizado em 2026-05-29.

## O que já está pronto e funcionando

- Interface web FastAPI em `web/` ([web/app.py](../web/app.py),
  [web/runner.py](../web/runner.py), [web/store.py](../web/store.py),
  [web/report.py](../web/report.py), [web/excel_parser.py](../web/excel_parser.py)).
- Páginas: upload do Excel, confirmação de tipo por processo (dropdown),
  agendamento (default 00:10 do próximo dia útil), tela de execução com log
  ao vivo via WebSocket, relatório final com exports DOCX e PDF.
- Launcher: `run_web.ps1` (manual) e `scripts/start_web_hidden.ps1` (silent).
- Auto-start 24/7: `scripts/install_web_autostart.ps1` registra Task Scheduler
  `ProjetoEuclidesWeb` (boot + heartbeat de 5 min).
- Acesso: **http://10.20.1.214:8080** (rede interna; regra de firewall
  "Projeto Euclides Web" já criada).
- Credenciais e-TCM: variáveis de usuário `ETCM_USERNAME=20386` e
  `ETCM_PASSWORD` (já configuradas). O runner promove para o processo.

## Mudanças aplicadas no último ajuste (2026-05-29 PM)

1. **Headless verdadeiro** no `web/runner.py`: `HEADLESS=true`,
   `SHOW_BROWSER=false`, `WATCH_MODE=false`, `DEVTOOLS=false`,
   `SLOWMO_MS=0`. Não abre mais janela do Chrome quando disparado pela web.
2. **Login mais rápido**: `LOGIN_MANUAL_WAIT_MS=15000` (era 60s),
   `PAUSE_AFTER_LOGIN_MS=1500` (era 5s). Como o login é por env vars, não
   precisa de tempo para humano.
3. **Log streaming corrigido para Windows**: `_stream_subprocess` agora lê em
   chunks de 2 KB com flush de buffer parcial (resolve full-buffering do
   stdout do child no Windows), insere `-u` no python.exe e usa
   `CREATE_NO_WINDOW` para esconder console do subprocesso.
4. **24/7**: scripts em `scripts/install_web_autostart.ps1` e
   `scripts/start_web_hidden.ps1`. O `install_*` precisa rodar **uma vez como
   admin** para registrar a task.

## Bugs / pendências conhecidos

### 🔴 1. Não envia para assinatura da Roseli Chaves (alto)

**Sintoma:** No último teste (TC/016628/2024, REITERACAO, executado pela web),
o robô **criou a Comunicação Processual mas não disparou o pedido de
assinatura** para Roseli Chaves. As env vars enviadas são
`REQUEST_SIGNATURE=true`, `ASSINANTE_NOME=Roseli Chaves`,
`SIGNER_NAME=Roseli Chaves`, `SKIP_SIGNATURE=false`, `SKIP_TRAMITACAO=true` —
ou seja, deveria pedir assinatura e não tramitar.

**Hipóteses:**
- O fluxo pós-conclusão do Ofício SSG em `src/main.py` pode estar parando antes
  do `_request_signature_*` quando rodando em headless puro. Procurar por
  `STOP_AFTER_OFICIO_CONCLUIDO` e `_post_conclusion_policy` em
  `src/main.py:48-63` e seguir o caminho até onde a assinatura é solicitada.
- O headless pode estar disparando proteções anti-bot do e-TCM no passo final
  (popup que não abre, modal que precisa de viewport). Olhar logs da execução
  em `web/jobs/<job_id>.json` e procurar pela última ação antes de o robô
  finalizar.
- O pytest pode ser rodado pelos runners PROD antes do main.py (ver
  `run_4_DILACAO_GILSON_2026_05_25.ps1:104-109`) e a web não roda. Talvez
  algum side-effect do pytest seja necessário — checar.

**Próximos passos sugeridos:**
1. Reproduzir com `HEADLESS=false` (botão "Disparar agora") para ver onde o
   robô para no fluxo de assinatura.
2. Comparar o caminho de execução de `web/runner.py` vs.
   `run_4_DILACAO_GILSON_2026_05_25.ps1` — env vars idênticas? Falta alguma?
3. Adicionar evento parseável no log de "Assinatura solicitada" para que o
   `_scan_for_events` em `web/runner.py` reflita o status corretamente.

### 🟡 2. Log ao vivo apareceu vazio no primeiro teste (médio)

**Sintoma:** Tela `run.html` mostrou "batch atual: REITERACAO" e status running,
mas a caixa de log ficou completamente vazia.

**Causa provável:** Buffering do stdout do subprocesso no Windows (já
corrigido pelas mudanças do item 3 acima). Validar reexecutando.

**Se persistir:** Inspecionar `web/jobs/<job_id>.json` — se `log_lines` está
populado, o problema é no WebSocket/JS. Se está vazio, ainda é buffering do
child — considerar trocar pipe por arquivo temporário (`tee`) ou usar
`win32console` para criar pseudo-tty.

### 🟢 3. Retry automático ausente

O runner atual roda cada batch UMA vez. Os wrappers PowerShell do projeto
fazem até 3 tentativas. Quando o item 1 estiver resolvido, replicar a lógica
de `run_4_DILACAO_GILSON_2026_05_25.ps1:178-207` (tabela `pending` +
`maxAttempts`) dentro do `runner.run_job`.

### 🟢 4. Limpeza dos uploads / jobs antigos

`web/uploads/` e `web/jobs/` crescem indefinidamente. Adicionar um job
periódico (ou rota admin) que apaga > 30 dias.

## Como o robô se integra

Por tipo selecionado (UTAP / DILACAO / REITERACAO), `web/runner.py` chama
`.venv/Scripts/python.exe -u src/main.py` como subprocess com env vars
montadas. **Cada tipo é um subprocess separado** (decisão técnica: o robô usa
`FORCE_TIPO` global e mudar entre tipos no mesmo processo vazaria estado).
Ordem fixa: UTAP → DILACAO → REITERACAO.

A página `/run/{job_id}` mantém WebSocket aberto que envia snapshots do
`store.Job` a cada 0.5s (log novo + status por processo + status geral).
Eventos parseados do stdout em `web/runner.py:RE_*` (linha ~96) atualizam o
status de cada processo.

## Como rodar manualmente

```powershell
# Subir manualmente (janela visível) para depurar:
powershell -ExecutionPolicy Bypass -File .\run_web.ps1

# Subir em background (igual ao Task Scheduler):
powershell -ExecutionPolicy Bypass -File .\scripts\start_web_hidden.ps1

# Registrar 24/7 no boot (UMA VEZ, como admin):
powershell -ExecutionPolicy Bypass -File .\scripts\install_web_autostart.ps1

# Ver status da task:
Get-ScheduledTask -TaskName ProjetoEuclidesWeb | Format-List State, LastRunTime, LastTaskResult
```

## Estado da árvore

- Branch atual: `fix/apopen-stop-after-oficio-concluido` (30 commits à frente
  do `main`).
- Working tree tem o `.venv` recém-criado (não commitar) e os novos arquivos
  do `web/` (commitar quando o item 1 estiver resolvido).
- Tag pertinente: `checkpoint-2026-05-15-assinatura-solicitada` no `origin`.

## Prompt sugerido para iniciar a sessão Codex

```
Estou retomando o Projeto Euclides (TCM-SP, robô APO-PEN) no Codex. O Claude
me deixou um documento de handoff em `docs/CODEX_HANDOFF.md` — leia ele
primeiro, depois `web/README.md`, `web/runner.py` e `src/main.py`.

Objetivo desta sessão: **fazer o robô disparado pela interface web em
http://10.20.1.214:8080 chegar até o pedido de assinatura para a Roseli
Chaves**, exatamente como os runners PowerShell PROD já fazem (ver
`run_4_DILACAO_GILSON_2026_05_25.ps1`).

Sintoma reportado pelo usuário: ao disparar pela web (processo
TC/016628/2024, REITERACAO), o robô criou a Comunicação Processual mas
*não* enviou o ofício para assinatura. As env vars enviadas pelo
`web/runner.py:_default_env_for_tipo` incluem `REQUEST_SIGNATURE=true`,
`ASSINANTE_NOME=Roseli Chaves`, `SKIP_SIGNATURE=false`,
`SKIP_TRAMITACAO=true`. Comparar com o que os runners .ps1 enviam e
identificar a diferença.

Antes de fazer mudanças invasivas, rode o robô com `HEADLESS=false`
(botão "Disparar agora" na interface) e observe onde o fluxo encerra.
Consulte `web/jobs/<job_id>.json` para ver o log persistido do último
disparo. Não toque na infraestrutura do Task Scheduler nem nos
scripts em `scripts/start_web_hidden.ps1` — eles estão estáveis.

Critério de pronto: um disparo pela interface web de um processo
REITERACAO termina com o evento "Assinatura solicitada para Roseli Chaves"
no log, igual ao que os runners .ps1 produzem.
```
