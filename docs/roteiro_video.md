# Roteiro do video (versao inicial)

Fonte: `video_silvana.mp4` (frames em `docs/video_frames/`).

Resumo do que esta visivel nos frames:
- O video inicia ja na Mesa de Trabalho, com a grid de processos de "Em confeccao APO-PEN".
- O menu esquerdo mostra "Processos > Unidade Tecnica de Oficios > Em confeccao APO-PEN".
- O dropdown "Operacoes" no topo aparece aberto com a opcao "Cadastrar comunicacao processual".
- Ha momentos de "Carregando", indicando acao disparada/submit.
- Em alguns momentos o painel de log inferior aparece expandido.

Frames chave (exemplos):
- 00:00:00 (interval_frame_000001_t=00-00-00.000.png): grid APO-PEN ja carregada.
- 00:07:35 (interval_frame_000456_t=00-07-35.000.png): dropdown Operacoes aberto.
- 00:07:37 (interval_frame_000458_t=00-07-37.000.png): overlay "Carregando".
- 00:13:00 (interval_frame_000781_t=00-13-00.000.png): painel de log inferior expandido.
- 00:15:11 (interval_frame_000912_t=00-15-11.000.png): dropdown com "Cadastrar comunicacao processual".
- 00:15:13 (interval_frame_000914_t=00-15-13.000.png): overlay "Carregando".
- 00:23:04 (interval_frame_001385_t=00-23-04.000.png): Operacoes setado para "Cadastrar comunicacao processual".
- 00:23:05 (interval_frame_001386_t=00-23-05.000.png): overlay "Carregando".

Pontos nao visiveis (precisam confirmacao):
- Tela de login (campos e botao).
- Modal/janela "Nova comunicacao processual".
- Campos de formulario (destinatario, relator, descricao, status, prazo).
- Fluxo de anexos/upload.

Plano B se faltar detalhe nos frames:
- Enviar 3-5 screenshots adicionais (login, grid, modal nova comunicacao, tela de anexos).
- Ou habilitar OCR e rodar `tools/ocr_frames.py` (Tesseract + Pillow).

Observacao: o `docs/steps.yaml` abaixo marca passos com `review: true` onde o video nao mostra claramente.
