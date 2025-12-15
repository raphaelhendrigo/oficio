# Execução completa em uma linha (use .env ou defina variáveis de ambiente).
python -m venv .venv; `
.\.venv\Scripts\pip install -r requirements.txt; `
.\.venv\Scripts\python -m playwright install; `
.\.venv\Scripts\python .\src\main.py
