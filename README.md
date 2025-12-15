# Automação e-TCM (Homologação)

Fluxo automatizado com Playwright (Python):
- Faz login no e‑TCM (homologação).
- Abre Processos → Em confecção APO‑PEN (grid).
- Pesquisa pelo campo “N° Processo” no grid e clica na lupa (como solicitado).
- Abre o visualizador (VisualizarDocsProtocolo), pega o último PDF, extrai texto.
- Gera um DOCX de resposta a partir de modelo (.docx) ou usa layout simples quando não houver modelo.
- Tenta anexar o DOCX ao processo pelo próprio portal.

Observação de credenciais: use `.env` (modelo em `.env.example`) ou variáveis de ambiente. Nunca versione senhas.

## Requisitos
- Windows 10/11
- Python 3.10+
- PowerShell
- Chrome/Chromium (Playwright instala os binários necessários)

## Instalação e execução
```powershell
# 1) Dentro da pasta do projeto
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
# 2) Instalar os navegadores do Playwright
.\.venv\Scripts\python -m playwright install
# 3) Opcional: crie seu .env a partir do modelo
Copy-Item .env.example .env
# 4) Edite .env com usuário e senha (ou defina por variável de ambiente)
# 5) Rodar
.\.venv\Scripts\python .\src\main.py
```

## Execução “one‑liner” (PowerShell)
```powershell
python -m venv .venv; `
.\.venv\Scripts\pip install -r requirements.txt; `
.\.venv\Scripts\python -m playwright install; `
$env:ETCM_USERNAME="<seu_usuario>"; $env:ETCM_PASSWORD="<sua_senha>"; `
.\.venv\Scripts\python .\src\main.py
```

## Variáveis (.env)
- `ETCM_URL` (padrão: login da homologação)
- `ETCM_USERNAME`, `ETCM_PASSWORD` (obrigatórias)
- `HEADLESS` (true/false, padrão false)
- `SHOW_BROWSER`/`WATCH_MODE` (true) força janela visível mesmo se HEADLESS=true; `SLOWMO_MS` ajusta o delay entre ações; `PAUSE_AFTER_LOGIN_MS` mantém pausa após login para acompanhar/solucionar captcha; `LOGIN_MANUAL_WAIT_MS` define quanto tempo esperar você concluir o login manualmente se a tela de login continuar aparecendo; `DEVTOOLS` abre o DevTools junto com o navegador.
- `OFICIO_TEMPLATES_DIR` diretório com modelos .docx
- `OFICIO_TEMPLATE` nome do arquivo .docx dentro do diretório (opcional)
- `ETCM_VIEWER_URL` URL direta do visualizador (opcional, fast‑path)
- `PROCESSOS_LIST` lista de processos (vírgula “,” ou “;”) para processamento em lote
- `PROCESS_ALL` (true) usa todos os processos exportados da planilha do grid
- `MAX_PROCESSOS` limita a quantidade em lote

Observação sobre modelos: python‑docx não abre arquivos .dotx (template do Word). Se seus modelos estão em “Modelos Ofícios” como .dotx, salve uma cópia em .docx e aponte via `OFICIO_TEMPLATES_DIR`/`OFICIO_TEMPLATE`. Quando nenhum .docx for encontrado, o script gera um ofício simples (fallback) preenchendo os campos extraídos do PDF.

### Caminho de modelos (exemplo)
`C:\Users\20386\OneDrive - tcm.sp.gov.br\Ofícios\ofício_automation\Modelos Ofícios`

## Como funciona
- Grid: preenche “N° Processo” e aciona a lupa da primeira linha retornada.
- Visualizador: identifica a última peça e tenta baixar o PDF embutido (ou pela “nova janela”).
- Texto do PDF: extraído com pypdf; campos como Processo, Assunto, Interessado e data são detectados por regex.
- DOCX: substitui placeholders ({{NUM_PROCESSO}}, {{DATA}}, {{ASSUNTO}}, etc.). Sem modelo .docx, cria um documento básico.
- Anexo: tenta localizar input de arquivo ou botão de upload e confirmar (Salvar/Gravar/Enviar/Confirmar).

## Saídas
- `output/apos_apo_pen.png` screenshot de conferência da tela de APO‑PEN
- `output/<processo>-ultimo-ato.pdf` último PDF capturado
- `output/oficio_<processo>.docx` ofício gerado

## Dicas
- Caso o visualizador demore, o fluxo usa timeouts com fallback automático.
- Se o HTML do portal mudar, ajuste seletores em `src/main.py` (grid/lupa/visualizador).

## Estrutura
```
oficio_automation/
  .env.example
  requirements.txt
  README.md
  src/
    main.py
  Modelos Ofícios/  (opcional, seus .docx)
  output/
```


## Atualizacoes (.dotx e conversao sem alteracoes)
- Suporte a modelos .dotx para gerar .docx automaticamente.
- Para manter o template sem alteracoes (apenas converter .dotx -> .docx), defina OFICIO_CONVERT_ONLY=true no .env.
- Continua valendo o fluxo normal de preenchimento de placeholders quando OFICIO_CONVERT_ONLY for alse.
