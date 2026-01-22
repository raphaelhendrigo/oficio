# Video silvana - gaps

## 1) Visao e filtro da grid
- Como esta hoje no codigo: nao garante selecao de "Visoes" nem filtro de "Distribuido para".
- Como o video mostra: "Visoes" em "Aposentadoria" e filtro de "Distribuido para" preenchido.
- Mudanca necessaria: adicionar steps para selecionar "Visoes" e aplicar filtro quando informado.
- Arquivos impactados: `docs/steps.yaml`, `src/bot.py`, `src/config.py`, `.env.example`.

## 2) Operacoes + carregamento
- Como esta hoje no codigo: abre dropdown "Operacoes" e clica "Cadastrar comunicacao processual"; espera apenas classes DevExpress.
- Como o video mostra: dropdown "Operacoes" e overlay "Carregando" apos a selecao.
- Mudanca necessaria: reforcar wait para texto "Carregando" (alem das classes).
- Arquivos impactados: `src/bot.py`, `docs/steps.yaml`.

## 3) Modal de comunicacao processual
- Como esta hoje no codigo: preenche destinatario, relator, descricao, status, prazo e anexa arquivo.
- Como o video mostra: modal nao visivel nos frames atuais.
- Mudanca necessaria: validar os campos reais e atualizar seletores/ordem se necessario.
- Arquivos impactados: `docs/steps.yaml`, `src/bot.py`, `src/selectors.py`.

## 4) Exportacao de planilha
- Como esta hoje no codigo: sempre tenta exportar a planilha (step de download).
- Como o video mostra: exportacao nao aparece nos frames atuais.
- Mudanca necessaria: manter como opcional ou validar se exportacao ainda e parte do fluxo.
- Arquivos impactados: `docs/steps.yaml`.
