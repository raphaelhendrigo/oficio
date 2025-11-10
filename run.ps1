# Execução completa em uma linha (edite as credenciais se quiser).
python -m venv .venv; `
.\.venv\Scripts\pip install -r requirements.txt; `
.\.venv\Scripts\python -m playwright install; `
$env:ETCM_USERNAME="20386"; $env:ETCM_PASSWORD="rhg#1004"; `
.\.venv\Scripts\python .\src\main.py
