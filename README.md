# e-TCM Oficios - Automacao baseada em video

Este repositorio transforma um fluxo gravado (video) em uma automacao web robusta
com Playwright, logs e evidencias por etapa.

## Requisitos
- Windows 10/11
- Python 3.10+
- PowerShell
- ffmpeg no PATH (para extrair frames)
- Playwright (instala os navegadores via comando)

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
