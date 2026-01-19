# e-TCM Oficios - Automacao baseada em video

Este repositorio transforma um fluxo gravado (video) em uma automacao web robusta
com Playwright, logs e evidencias por etapa.

## Requisitos
- Windows 10/11
- Python 3.10+
- PowerShell
- ffmpeg no PATH (para extrair frames)
- Playwright (instala os navegadores via comando)
- Chrome/Chromium (Playwright instala os binarios necessarios)

## Setup rapido
```powershell
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\python -m playwright install
```

## A) Extrair frames do video
```powershell
# Intervalo (1s)
.\.venv\Scripts\python tools\extract_frames.py --mode interval --every 1

# Mudanca de cena (threshold mais sensivel)
.\.venv\Scripts\python tools\extract_frames.py --mode scene --scene-threshold 0.05

# Gerar a folha de contato
.\.venv\Scripts\python tools\make_contact_sheet.py
```

Abra `docs/video_frames/index.html` para revisar rapidamente.

## B) (Opcional) OCR dos frames
Requer Tesseract instalado e no PATH.
```powershell
.\.venv\Scripts\pip install pytesseract pillow
.\.venv\Scripts\python tools\ocr_frames.py --lang por
```

## C) Roteiro e steps
- Roteiro humano: `docs/roteiro_video.md`
- Steps executaveis: `docs/steps.yaml`

Edite `docs/steps.yaml` para ajustar seletores e confirmar campos antes de rodar em producao.

## D) Rodar o bot
```powershell
# Execucao normal
.\.venv\Scripts\python src\bot.py --mode run

# Debug (pausa a cada step)
.\.venv\Scripts\python src\bot.py --mode debug

# Dry-run (nao clica, apenas valida presenca)
.\.venv\Scripts\python src\bot.py --mode dry-run
```

### Variaveis de ambiente (.env)
Veja `.env.example` para um modelo completo. Principais:
- `ETCM_USER`, `ETCM_PASS`
- `BASE_URL`
- `DOWNLOAD_DIR`
- `PROCESSOS_LIST` ou `PROCESS_ALL`
- `MODE` (run/debug/dry-run)

## Saidas e evidencias
- `artifacts/downloads/`: planilhas baixadas
- `artifacts/evidence/`: screenshots por etapa
- `artifacts/html/`: HTML da pagina em caso de erro
- `logs/`: logs da execucao

## Scripts legados
Existe automacao anterior em `src/main.py` e `src/etcm_oficios_apo_pen.py`.
O fluxo atual usa `src/bot.py` + `docs/steps.yaml`.

## Automacao legada (src/main.py)
Fluxo Playwright que faz login no e-TCM, abre processos, baixa PDF, gera DOCX e tenta anexar no portal.
Observacao de credenciais: use `.env` (modelo em `.env.example`) ou variaveis de ambiente. Nunca versione senhas.

### Instalar e executar (legado)
```powershell
# 1) Dentro da pasta do projeto
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
# 2) Instalar os navegadores do Playwright
.\.venv\Scripts\python -m playwright install
# 3) Opcional: crie seu .env a partir do modelo
Copy-Item .env.example .env
# 4) Edite .env com usuario e senha (ou defina por variavel de ambiente)
# 5) Rodar
.\.venv\Scripts\python .\src\main.py
```

### Execucao one-liner (PowerShell, legado)
```powershell
python -m venv .venv; `
.\.venv\Scripts\pip install -r requirements.txt; `
.\.venv\Scripts\python -m playwright install; `
$env:ETCM_USERNAME="<seu_usuario>"; $env:ETCM_PASSWORD="<sua_senha>"; `
.\.venv\Scripts\python .\src\main.py
```

### Variaveis (legado)
- `ETCM_URL` (padrao: login da homologacao)
- `ETCM_USERNAME`, `ETCM_PASSWORD` (obrigatorias)
- `HEADLESS` (true/false, padrao false)
- `SHOW_BROWSER`/`WATCH_MODE` (true) forca janela visivel mesmo se HEADLESS=true; `SLOWMO_MS` ajusta o delay entre acoes; `PAUSE_AFTER_LOGIN_MS` mantem pausa apos login para acompanhar/solucionar captcha; `LOGIN_MANUAL_WAIT_MS` define quanto tempo esperar o login manual; `DEVTOOLS` abre o DevTools junto com o navegador.
- `OFICIO_TEMPLATES_DIR` diretorio com modelos .docx
- `OFICIO_TEMPLATE` nome do arquivo .docx dentro do diretorio (opcional)
- `ETCM_VIEWER_URL` URL direta do visualizador (opcional, fast-path)
- `PROCESSOS_LIST` lista de processos (separados por virgula ou ponto-e-virgula)
- `PROCESS_ALL` (true) usa todos os processos exportados da planilha do grid
- `MAX_PROCESSOS` limita a quantidade em lote

Observacao sobre modelos: python-docx nao abre arquivos .dotx (template do Word). Se seus modelos estao em "Modelos Oficios" como .dotx, salve uma copia em .docx e aponte via `OFICIO_TEMPLATES_DIR`/`OFICIO_TEMPLATE`. Quando nenhum .docx for encontrado, o script gera um oficio simples (fallback) preenchendo os campos extraidos do PDF.

### Caminho de modelos (exemplo)
`C:\Users\20386\OneDrive - tcm.sp.gov.br\Oficios\oficio_automation\Modelos Oficios`

### Saidas (legado)
- `output/apos_apo_pen.png` screenshot de conferencia da tela de APO-PEN
- `output/<processo>-ultimo-ato.pdf` ultimo PDF capturado
- `output/oficio_<processo>.docx` oficio gerado

### Atualizacoes .dotx (legado)
- Suporte a modelos .dotx para gerar .docx automaticamente.
- Para manter o template sem alteracoes (apenas converter .dotx -> .docx), defina `OFICIO_CONVERT_ONLY=true` no .env.
- Continua valendo o fluxo normal de preenchimento de placeholders quando `OFICIO_CONVERT_ONLY=false`.
