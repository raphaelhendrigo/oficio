# Pipeline de aprendizado — classificação automática (UTAP / DILACAO / REITERACAO)

> Objetivo: substituir gradualmente a decisão manual do Gilson por um
> classificador treinado com o histórico de decisões dele. Esta versão
> mantém **intacto** o fluxo atual e só adiciona captura de dados.

## Estado em 2026-06-09

- ✅ **Captura de labels** ativa: cada confirmação na tela `/confirm/...`
  grava uma linha em `web/labels/labels.jsonl` via
  [web/labels.py](../web/labels.py). É additivo e não interfere no robô.
- ✅ **Painel de progresso** em `/labels` (`http://10.20.1.214:8080/labels`).
- ⏳ **Captura de features** — pendente. Só com label + features dá para
  treinar.
- ⏳ **Treino e inferência** — pendente.

## Por que só os labels não bastam

Hoje gravamos `(processo, tipo)`. O número do processo isolado não diz
nada para um classificador — `TC/016628/2024` não tem padrão semântico
visível. O Gilson decide olhando para o que está **dentro** do processo:
último ato, secretaria, conselheiro relator, número de ofícios SSG
anteriores, presença de "manutap" ou "providências", data da última
movimentação etc.

Essas são as **features**. Precisam ser capturadas no momento em que o
robô abre o processo no e-TCM (já é parte do fluxo de
[src/main.py](../src/main.py)). Plano:

1. Identificar no `src/main.py` o ponto onde o processo já foi aberto e
   os metadados estão disponíveis (após `_extract_referencia_*`,
   `_get_secretaria`, leitura do último ato).
2. Criar `web/features.py` análogo a `web/labels.py` — append-only JSONL
   em `web/labels/features.jsonl`, chaveado por número do processo.
3. Chamar `features.record(processo, **dados)` a partir do main.py em
   modo "fire-and-forget" (try/except silencioso, idêntico aos labels).
4. Em treino offline, fazer JOIN entre `labels.jsonl` e `features.jsonl`
   por `processo`.

Features candidatas (refinar com o Gilson):

| Campo                                | Tipo     | Por que ajuda                                              |
| ------------------------------------ | -------- | ---------------------------------------------------------- |
| `secretaria`                         | categ.   | Saúde/Educação/Geral mudam a distribuição de tipos.        |
| `ultimo_ato_tipo`                    | categ.   | "Despacho", "Encaminhamento", "Decisão"... separa muito.   |
| `ultimo_ato_origem`                  | categ.   | UTAP, Gabinete, Conselheiro Relator.                       |
| `dias_desde_ultimo_ato`              | numérico | Reiteração tende a ter intervalo maior.                    |
| `n_oficios_ssg_anteriores`           | numérico | 0 = primeira tentativa; >0 forte sinal de REITERACAO.      |
| `tem_manutap_referenced`             | bool     | Palavra-chave forte para UTAP/manutap.                     |
| `tem_decisao_julgada`                | bool     | Separa juízo singular / dilação.                           |
| `conselheiro_relator`                | categ.   | Distribuição empírica por relator.                         |
| `presenca_palavra_chave_dilacao`     | bool     | "prazo", "dilação", "60 dias".                             |

## Timeline honesta

Premissas:
- Gilson roda lotes em dias úteis. Ritmo observado nos runners do
  projeto (run_4, run_9, run_32) sugere **15–30 labels/dia útil**.
- 22 dias úteis por mês → **330–660 labels/mês**.
- 3 classes, features tabulares relativamente separáveis.
- Modelo: começa com regras + gradient boosting (LightGBM/sklearn). Não
  precisa de deep learning para esse volume.

| Marco                              | Labels  | A 15/dia        | A 30/dia        |
| ---------------------------------- | ------- | --------------- | --------------- |
| Sinal mínimo (testa hipótese)      | ~100    | ~7 dias úteis   | ~4 dias úteis   |
| Baseline assistido (sugere e confirma) | ~400    | ~27 dias úteis  | ~14 dias úteis  |
| Auto-classificação confiável       | ~1.200  | ~80 dias úteis  | ~40 dias úteis  |
| Substituição quase completa        | ~2.500  | ~170 dias úteis | ~85 dias úteis  |

Em **calendário corrido** (considerando feriados, férias, dias sem
disparo):

- **~1 mês** para começar a sugerir tipo na interface (Gilson confirma).
- **~3–4 meses** para auto-classificar os casos de alta confiança e
  Gilson revisar só os incertos.
- **~6–9 meses** para substituição quase completa.

Esses números só valem se:
1. **As features forem boas.** Se a captura ficar pobre, nunca chega lá.
   Por isso o próximo passo crítico é o `features.py`.
2. **As classes forem coerentes.** Se REITERACAO domina (digamos 70%),
   o modelo aprende a "chutar reiteração" e parece bom. Validar com
   `precision`/`recall` por classe, não só acurácia.
3. **Gilson seguir disparando pela web.** Se voltar para os
   `run_*.ps1`, perdemos o sinal — esses runners não passam pelo
   `confirm_post`.

## Como acompanhar

- Painel: `http://10.20.1.214:8080/labels`
- API: `GET /api/labels/stats` (retorna JSON com totais + estimativas).
- Dump bruto: `web/labels/labels.jsonl` (uma linha JSON por decisão).

## Treino — quando chegar a hora

```python
import json, pandas as pd
from pathlib import Path

labels = pd.DataFrame([json.loads(l) for l in Path("web/labels/labels.jsonl").read_text(encoding="utf-8").splitlines() if l.strip()])
features = pd.DataFrame([json.loads(l) for l in Path("web/labels/features.jsonl").read_text(encoding="utf-8").splitlines() if l.strip()])
df = labels.merge(features, on="processo", how="inner")
# X, y, fit ...
```

Modelo inicial: `sklearn.ensemble.GradientBoostingClassifier` ou
`lightgbm.LGBMClassifier`. Métrica: `f1_macro` (porque as classes serão
desbalanceadas). Validação: split estratificado por mês para simular
deploy real.

## O que NÃO vamos fazer

- Não treinamos enquanto não tivermos features. Treinar com só o número
  do processo é exercício inútil.
- Não substituímos a decisão do Gilson antes do marco "auto-confiável".
  Modelo pré-treino tem mais chance de ensinar mau hábito do que ajudar.
- Não tocamos no fluxo atual de execução. A coleta é puramente
  observacional até segunda ordem.
