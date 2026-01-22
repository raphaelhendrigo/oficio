# Video silvana - index

Fonte: `video_silvana.mp4`

Artefatos:
- Frames: `docs/video_silvana/frames`
- Manifest: `docs/video_silvana/frames_manifest.json`
- Metadata: `docs/video_silvana/metadata.json`
- Transcript: `docs/video_silvana/transcript.txt` (vazio)

Observacoes rapidas:
- Mesa de trabalho aberta em "Em confeccao APO-PEN" com grid carregada.
- Filtro de coluna "Distribuido para" aparece preenchido no grid.
- "Visoes" selecionado como "Aposentadoria".
- Dropdown "Operacoes" usado para escolher "Cadastrar comunicacao processual".
- Overlay "Carregando" aparece apos a escolha.

Frames chave:
- 00:00:00 `docs/video_silvana/frames/00-00-00.png`
- 00:07:34 `docs/video_silvana/frames/00-07-34.png`
- 00:07:35 `docs/video_silvana/frames/00-07-35.png`
- 00:07:37 `docs/video_silvana/frames/00-07-37.png`
- 00:15:11 `docs/video_silvana/frames/00-15-11.png`
- 00:23:04 `docs/video_silvana/frames/00-23-04.png`
- 00:23:05 `docs/video_silvana/frames/00-23-05.png`
- 00:23:10 `docs/video_silvana/frames/00-23-10.png`
- 00:23:20 `docs/video_silvana/frames/00-23-20.png`

Limitacoes:
- ffmpeg nao encontrado e opencv-python nao instalado; frames gerados por reuso de `docs/video_frames`.
- Os frames atuais cobrem ate 00:24:27; o video completo tem 00:36:56.
- O video `video_silvana.mp4` e mais recente que os frames existentes; reextrair quando houver ffmpeg/opencv.

TODO validacao:
- [00:00:00] Confirmar se a tela de login ocorre antes do inicio do video (frame `docs/video_silvana/frames/00-00-00.png`).
- [00:23:05] Confirmar abertura do modal de "Cadastrar comunicacao processual" e seus campos (frame `docs/video_silvana/frames/00-23-05.png`).
- [00:23:05] Confirmar fluxo de anexo/upload e mensagem de sucesso (frame `docs/video_silvana/frames/00-23-05.png`).
