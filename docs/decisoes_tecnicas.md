# Decisoes tecnicas

## Por que Playwright
- Bom suporte a DevExpress e SPAs (waits explicitos e auto-wait).
- Controle simples de downloads, popups e storage_state.
- Facil de rodar em modo headless ou visivel.

## Estrategia de seletores
- Priorizar id/name/data-* quando disponiveis.
- Usar texto de labels como fallback (id_or_label).
- Evitar XPath fragil; usar CSS e roles quando possivel.
- Manter o mapeamento em `src/selectors.py` e `docs/steps.yaml`.

## DevExpress
- Elementos costumam ter sufixos `_I` (inputs) e `_B-1` (botoes).
- Loading panels sao comuns (`.dxlpLoadingPanel`, `.dxgvLoadingPanel`).
- O bot espera o fim do loading antes de prosseguir.

## Roteiro em YAML
- `docs/steps.yaml` e a fonte de verdade do fluxo.
- Campos com `review: true` indicam pontos que precisam de ajuste apos revisar os frames.
- O runner suporta `scope: per_process` para passos repetidos.

## Observabilidade e resiliencia
- Logs estruturados em `logs/`.
- Evidencias por etapa em `artifacts/evidence/`.
- HTML de erro em `artifacts/html/`.
- Retries simples para evitar falhas por timing.
