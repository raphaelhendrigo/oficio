# Relatório de execução — 5 processos APO-PEN

| Item | Valor |
|---|---|
| Data da execução | 2026-05-13 11:11–11:28 (PROD) |
| Ambiente | **PRODUÇÃO** (`https://etcm.tcm.sp.gov.br/`) |
| Branch | `fix/preserve-at-tokens-modelos-apopen` |
| Commits relevantes | `693db63`, `c551649`, `54d794e`, `03ff264`, `7caf9cd`, `b90ba1b`, `9a9c28a` |
| Configurações principais | `OFICIO_TEMPLATE_MODE=auto`, `OFICIO_PRESERVE_AT_TOKENS=true`, `SAFE_DELETE_OWN_DRAFTS=false`, `ENVIRONMENT=producao`, `REUSE_EXISTING_OFICIO=false`, `ASSINANTE_NOME=Roseli Chaves`, `TRAMITAR_DESTINO=Em assinatura` |
| Log da execução | `logs/run_5_processos_PROD_20260513_111139.log` |
| Output do background | `bfe9n1x4z.output` (exit code 0) |

## Resultado consolidado

**Nenhuma alteração foi feita em PROD nesta execução.** O robô fez login com sucesso (`Sessão salva em: storage_state.json`), abriu o visualizador de cada processo, mas em **todos os 5** o visualizador não retornou peças. Sem PDF, o robô não chega a gerar DOCX, anexar, pedir assinatura ou tramitar — nada foi escrito.

| Processo | Visualizador abriu | Peças encontradas | Próxima etapa atingida | Pendência |
|---|---|---|---|---|
| TC/007902/2022 | ✅ | ❌ Nenhuma | (parou aqui) | Investigar fila atual |
| TC/008636/2022 | ✅ | ❌ Nenhuma | (parou aqui) | Investigar fila atual |
| TC/008084/2023 | ✅ | ❌ Nenhuma | (parou aqui) | **Já tramitado em 12/05 ontem** (ver log antigo abaixo) |
| TC/013838/2023 | ✅ | ❌ Nenhuma | (parou aqui) | Investigar fila atual |
| TC/018149/2024 | ✅ | ❌ Nenhuma | (parou aqui) | Investigar fila atual |

## Por que "Nenhuma peça encontrada"

A função `click_last_piece_and_open_pdf` ([src/main.py:2255](src/main.py#L2255)) procura na árvore do visualizador (`#splLeitorDocumentos_pgcPecas_trePecas`) por âncoras `a[onclick*='LerPDF']` ou similares. Quando nenhuma é detectada em ~30s, ela retorna `(None, None)` com o aviso registrado.

Possíveis causas (em ordem de probabilidade):

1. **Os processos já foram tramitados em rodadas anteriores** e não estão mais no fluxo APO-PEN com peças disponíveis. Há **evidência direta para TC/008084/2023** em [logs/prod_remaining2_20260512_182657.log](logs/prod_remaining2_20260512_182657.log) (12/05 às 18:26):

   ```
   [1/3] Tratando processo: TC/008084/2023
   Processo TC/008084/2023: baixando ultima peca/PDF.
   PDF salvo em: ...TC_008084_2023-ultimo-ato.pdf
   Oficio gerado (DOTX via python-docx) em: ...oficio_TC_008084_2023.docx
   Anexo do DOCX concluido.
   Assinante selecionado para TC/008084/2023: ROSELI DE MORAIS CHAVES
   Assinatura solicitada para Roseli Moraes Chaves no processo TC/008084/2023.
   Processo TC/008084/2023 tramitado para 'Em assinatura' e localizado em Em Assinatura.
   ```

   Esse processo está, segundo o log anterior do próprio Codex, **completo** desde 12/05 — provavelmente por isso não tem mais peça para a fila APO-PEN.

2. **Visualizador abre mas com `iframe` ainda carregando** ao chegar no seletor — a função tem timeout de 30s; é possível que em PROD o carregamento exceda esse limite em alguns processos. Mas o fato de ser **0 peças em 5/5** torna essa hipótese menos provável; o cenário 1 explica melhor.

3. **Mudança de fila** — depois de tramitado para "Em assinatura", o processo deixa a fila APO-PEN. O script ainda filtra na grid APO-PEN ("Lista de processos fornecida; abrindo pasta APO-PEN sem exportar planilha") e o visualizador não encontra peças no contexto certo.

## Pré-condições verificadas (positivas)

- **Testes pytest**: 81/81 verdes antes da execução (passo `[1/2]` do runner).
- **Sanitização do repositório**: senha removida do working tree e do histórico via `git filter-repo` + force-push (commit `693db63`).
- **Preservação de tokens `@@`**: garantida em código via `docx_utils.assert_at_tokens_preserved` pós-geração (validado com 5 testes end-to-end contra modelos reais Educação/Saúde/Geral).
- **Login automático em PROD**: bem-sucedido (`Sessao salva em: storage_state.json`). Sem captcha, sem erro de credencial.
- **Sem efeito destrutivo**: `SAFE_DELETE_OWN_DRAFTS=false`, e ninguém criou comunicação/anexou/derrubou nada nesta execução.

## Pendências e riscos abertos

1. **Credencial de PROD exposta anteriormente** — a senha literal foi removida deste relatório e não deve constar em arquivos, logs ou commits. Se houver suspeita de exposição externa, a troca da senha deve ser tratada fora do repositório.
2. **Localização real dos 5 processos hoje** — não confirmada. Pode ser que estejam:
   - já em "Em assinatura" (caso de TC/008084/2023, comprovado)
   - em fila diferente após tramitação manual feita por outro operador
   - ainda em APO-PEN, mas com peças que o visualizador atual não enxerga
3. **Limpeza de minutas anteriores em PROD** — desativada por padrão (`SAFE_DELETE_OWN_DRAFTS=false`). Mantém comportamento conservador: nada é derrubado automaticamente.
4. **Regra de dilação** — extração de "dias concedidos" do despacho do Conselheiro **não implementada** nesta sessão. Quando algum processo de dilação rodar, a linha `Referência` deve ser revisada manualmente.
5. **Linha `Encaminha` (peça MANUTAP-OF)** — comportamento atual segue o WIP anterior do Codex, sem regras novas adicionadas.

## Evidências e artefatos

- Log desta execução: [logs/run_5_processos_PROD_20260513_111139.log](logs/run_5_processos_PROD_20260513_111139.log)
- Log antigo que mostra TC/008084/2023 completo: [logs/prod_remaining2_20260512_182657.log](logs/prod_remaining2_20260512_182657.log)
- Pasta de screenshots (vazia nesta execução): [artifacts/evidence/](artifacts/evidence/)
- Pasta de output DOCX (vazia nesta execução): [output/](output/)
- Backup pré-`filter-repo`: `backup_pre_filter_repo/oficio_pre_filter_repo_*.bundle`

## Próximos passos recomendados

1. **Verificar manualmente no e-TCM** o estado real dos 5 processos:
   - em qual fila/situação estão hoje
   - se já têm Ofício SSG em "Em assinatura" ou já assinado
   - se ainda há providência pendente para o robô
2. Após confirmar quais processos ainda precisam de Ofício SSG nesta rodada, voltar a rodar com a lista reduzida.
3. **Trocar a senha do e-TCM** assim que possível, mesmo que o repo agora seja privado.
4. Abrir PR a partir da branch `fix/preserve-at-tokens-modelos-apopen` para revisão e merge das melhorias offline (preservação `@@`, descrição correta, helper de assinante).
5. Em sessão futura com supervisão visual, plumbar a regra de dilação (dias do despacho) e a regra `Encaminha` com peça MANUTAP-OF por nome.
6. Pedir ao GitHub Support a purga dos orphan refs do repo (ou esperar ~30 dias para GC automático).
