# Projeto Euclides — interface web

Interface HTTP para o Gilson disparar o robô APO-PEN a partir de uma planilha
Excel, marcando o tipo de cada processo (Reiteração, Dilação ou Manutap/UTAP).

## Subir o servidor

```powershell
powershell -ExecutionPolicy Bypass -File .\run_web.ps1
```

A primeira execução instala as dependências web (`fastapi`, `uvicorn`,
`python-multipart`, `jinja2`). O servidor sobe em `0.0.0.0:8080` e imprime o
link de rede.

Variável opcional: `EUCLIDES_WEB_PORT` para mudar a porta.

## Liberar firewall (Windows, uma vez só)

```powershell
# Em PowerShell como Administrador
New-NetFirewallRule -DisplayName "Projeto Euclides Web" `
    -Direction Inbound -Protocol TCP -LocalPort 8080 -Action Allow `
    -Profile Domain,Private
```

Se o Gilson estiver em outra subnet/perfil, ajustar `-Profile`.

## Fluxo de uso pelo Gilson

1. Acessa `http://<ip-da-vm>:8080` pela máquina dele.
2. Envia o arquivo Excel (.xlsx ou .xls) com a lista de processos.
3. Confirma o tipo de cada processo no dropdown (REITERACAO / DILACAO / UTAP / pular).
4. Clica em **Disparar fluxo**.
5. Acompanha o log ao vivo e a tabela de status por processo na tela de execução.

## Arquitetura

- `web/app.py` — FastAPI: rotas `/`, `/upload`, `/confirm/{id}`, `/run/{id}`,
  `/api/job/{id}` e WebSocket `/run/{id}/ws`.
- `web/excel_parser.py` — varre todas as células do .xlsx/.xls procurando o
  padrão `TC/xxxxx/aaaa`.
- `web/store.py` — estado dos jobs em memória + persistência em
  `web/jobs/<job_id>.json` para auditoria. Garante 1 job rodando por vez.
- `web/runner.py` — orquestra sub-batches por tipo. Para cada tipo com
  processos selecionados, dispara `python src/main.py` como subprocesso com o
  `FORCE_TIPO` adequado (UTAP / DILACAO / REITERACAO) e faz streaming do
  stdout. Replica as env vars dos runners `run_*_PROD_*.ps1`.
- `web/templates/` — Jinja2: `_base.html`, `index.html`, `confirm.html`,
  `run.html`.
- `web/static/style.css` — UI moderna (Inter + JetBrains Mono, gradientes,
  dark theme).
- `web/uploads/` — planilhas recebidas (auditoria).
- `web/jobs/` — snapshot de cada job em JSON (auditoria).

## Por que sub-batches separados?

O robô (`src/main.py`) usa `FORCE_TIPO` para escolher a pasta de modelos
(`modelos_dilacao`, `modelos_reiteracao`, `modelos_utap`). Misturar tipos numa
única execução exigiria reescrever a seleção de template. O caminho mais seguro
hoje, igual ao wrapper PowerShell `run_LOTE_2026_05_26_DILACAO_E_REITERACAO.ps1`,
é encadear subprocessos — cada um com seu `FORCE_TIPO` isolado.

## Limites e atenção

- Roda 1 job por vez (lock). Disparar 2 em paralelo abriria 2 navegadores
  Playwright no mesmo perfil — proibido.
- Sem autenticação: confiar na rede interna.
- O browser do robô abre na sessão Windows da própria VM. A interface web só
  inicia/observa; o login no e-TCM ainda usa credenciais via env vars como nos
  runners PROD.
- Sem retries automáticos (diferente do `run_*_GILSON*.ps1`, que faz 3
  tentativas). Versão 1 — adicionar depois se necessário.
