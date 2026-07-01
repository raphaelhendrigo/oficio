"""Classificador automático de processos (UTAP / DILACAO / REITERACAO).

Treina em cima do JOIN entre `web/labels/labels.jsonl` (decisão do Gilson)
e `web/labels/features.jsonl` (peças do processo capturadas pelo robô).

Pipeline:
  - Features de texto: TF-IDF dos nomes das peças (word + char n-grams).
  - Features booleanas/numéricas: flags has_* + pieces_count + posições.
  - Modelo: LogisticRegression com class_weight='balanced'
    (lida com forte desbalanceamento DILACAO 72% / UTAP 20% / REITERACAO 8%).

Persistência: `web/labels/model.joblib` + `web/labels/model_meta.json`
(métricas de treino, tamanho do dataset, timestamp).

Inferência: `predict(features_dict) -> (tipo, prob_dict)`. Robusto a chaves
faltantes — o robô grava features novas em campos extras e nada quebra.
"""
from __future__ import annotations

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import joblib
import numpy as np
from sklearn.compose import ColumnTransformer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    classification_report,
    confusion_matrix,
    f1_score,
)
from sklearn.model_selection import StratifiedKFold, cross_val_predict
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

from . import features as _features
from . import labels as _labels

ROOT = Path(__file__).resolve().parents[1]
LABELS_DIR = ROOT / "web" / "labels"
MODEL_PATH = LABELS_DIR / "model.joblib"
META_PATH = LABELS_DIR / "model_meta.json"

CLASSES = ("UTAP", "DILACAO", "REITERACAO")

# Lista CANÔNICA de flags numéricas/booleanas usadas pelo modelo. Quando
# adicionarmos novas features em `web/features.py`, basta acrescentar aqui
# que o pipeline pega no próximo treino.
NUMERIC_FIELDS = (
    "pieces_count",
    "pos_last_manutap_of",
    "pos_last_ssg",
    "ssg_after_manutap",
    "has_manutap_of",
    "has_manutap",
    "has_ssg",
    "has_oficio_ssg",
    "has_requerimento_dilacao",
    "has_dilacao_word",
    "has_reiteracao_word",
    "has_decisao",
    "has_acordao",
    "has_despacho",
    "has_encaminha",
    "has_julga",
    "has_aposentadoria",
    "has_pensao",
    "has_providencias",
)


def _coerce(val: Any) -> float:
    if isinstance(val, bool):
        return 1.0 if val else 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if val is None:
        return 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _row_to_text(row: dict) -> str:
    names = row.get("pieces_names") or []
    if isinstance(names, list):
        return " | ".join(str(n) for n in names)
    return str(names)


def _row_to_numeric(row: dict) -> list[float]:
    return [_coerce(row.get(k)) for k in NUMERIC_FIELDS]


def build_dataset() -> tuple[list[dict], list[str]]:
    """Une labels + features pelo número do processo. Para processos com
    múltiplos labels (raro), mantém o MAIS RECENTE — reflete a decisão atual
    do Gilson.
    """
    labs = _labels.load_all()
    feats = _features.load_all()
    if not labs or not feats:
        return [], []
    # Index features por processo (mais recente vence).
    feat_by_proc: dict[str, dict] = {}
    for f in feats:
        proc = (f.get("processo") or "").strip().upper()
        if not proc:
            continue
        prev = feat_by_proc.get(proc)
        if not prev or (f.get("ts", 0) > prev.get("ts", 0)):
            feat_by_proc[proc] = f
    # Index labels por processo (mais recente vence).
    label_by_proc: dict[str, dict] = {}
    for l in labs:
        proc = (l.get("processo") or "").strip().upper()
        tipo = (l.get("tipo") or "").strip().upper()
        if not proc or tipo not in CLASSES:
            continue
        prev = label_by_proc.get(proc)
        if not prev or (l.get("ts", 0) > prev.get("ts", 0)):
            label_by_proc[proc] = l
    X_rows: list[dict] = []
    y: list[str] = []
    for proc, lab in label_by_proc.items():
        feat = feat_by_proc.get(proc)
        if not feat:
            continue
        X_rows.append(feat)
        y.append(lab["tipo"])
    return X_rows, y


def _build_pipeline() -> Pipeline:
    """Pipeline com TF-IDF dos nomes + features numéricas escaladas."""
    text_vec = TfidfVectorizer(
        analyzer="word",
        ngram_range=(1, 2),
        min_df=2,
        max_df=0.95,
        max_features=4000,
        lowercase=True,
        strip_accents="unicode",
    )
    char_vec = TfidfVectorizer(
        analyzer="char_wb",
        ngram_range=(3, 5),
        min_df=2,
        max_df=0.95,
        max_features=4000,
        lowercase=True,
        strip_accents="unicode",
    )
    pre = ColumnTransformer(
        transformers=[
            ("text_word", text_vec, "text"),
            ("text_char", char_vec, "text"),
            ("num", StandardScaler(with_mean=False), NUMERIC_FIELDS),
        ],
        remainder="drop",
        sparse_threshold=0.3,
    )
    clf = LogisticRegression(
        max_iter=2000,
        class_weight="balanced",
        C=1.0,
        solver="liblinear",
    )
    return Pipeline([("pre", pre), ("clf", clf)])


def _to_dataframe(rows: list[dict]):
    """Constrói o frame em formato esperado pelo ColumnTransformer."""
    import pandas as pd
    data = {"text": [_row_to_text(r) for r in rows]}
    for k in NUMERIC_FIELDS:
        data[k] = [_coerce(r.get(k)) for r in rows]
    return pd.DataFrame(data)


def train_and_save(verbose: bool = True) -> dict:
    """Treina, avalia com CV estratificado e salva. Retorna metadata."""
    LABELS_DIR.mkdir(parents=True, exist_ok=True)
    X_rows, y = build_dataset()
    n = len(X_rows)
    if n < 30:
        raise RuntimeError(
            f"Dataset pequeno demais ({n} linhas) — espere mais labels antes de treinar."
        )
    df = _to_dataframe(X_rows)
    y_arr = np.array(y)

    pipe = _build_pipeline()

    # CV estratificado para metricas justas. Se a classe minoritária for
    # menor que k, reduzimos k.
    class_counts = {c: int((y_arr == c).sum()) for c in CLASSES}
    min_class = min(class_counts.values()) if class_counts else 0
    k = min(5, max(2, min_class))
    skf = StratifiedKFold(n_splits=k, shuffle=True, random_state=42)
    y_pred_cv = cross_val_predict(pipe, df, y_arr, cv=skf, method="predict")
    report = classification_report(y_arr, y_pred_cv, labels=list(CLASSES), output_dict=True, zero_division=0)
    cm = confusion_matrix(y_arr, y_pred_cv, labels=list(CLASSES))
    f1_macro = float(f1_score(y_arr, y_pred_cv, labels=list(CLASSES), average="macro", zero_division=0))

    # Treino final em TODO o dataset para o modelo de producao.
    pipe.fit(df, y_arr)
    joblib.dump(pipe, MODEL_PATH)

    meta = {
        "trained_at": datetime.now().isoformat(timespec="seconds"),
        "n_samples": n,
        "class_counts": class_counts,
        "cv_folds": k,
        "f1_macro_cv": f1_macro,
        "report_cv": report,
        "confusion_matrix_cv": {
            "labels": list(CLASSES),
            "matrix": cm.tolist(),
        },
        "model_path": str(MODEL_PATH),
        "sklearn_version": __import__("sklearn").__version__,
    }
    META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    if verbose:
        print(f"=== Treino concluido ===")
        print(f"n_samples = {n}  classes = {class_counts}")
        print(f"f1_macro (CV {k}-fold) = {f1_macro:.3f}")
        print("Por classe (CV):")
        for c in CLASSES:
            r = report.get(c, {})
            print(f"  {c:11s} precision={r.get('precision',0):.2f}  "
                  f"recall={r.get('recall',0):.2f}  f1={r.get('f1-score',0):.2f}  "
                  f"support={int(r.get('support',0))}")
        print("Confusao (linhas=verdadeiro, colunas=predito):")
        print(f"  {' '*11} " + " ".join(f"{c:>11s}" for c in CLASSES))
        for i, c in enumerate(CLASSES):
            print(f"  {c:11s} " + " ".join(f"{int(v):>11d}" for v in cm[i]))
        print(f"\nModelo salvo em: {MODEL_PATH}")
    return meta


_CACHED_PIPE = None
_CACHED_MTIME = 0.0


def _load_pipe():
    """Carrega o modelo do disco (hot-reload se o arquivo foi atualizado)."""
    global _CACHED_PIPE, _CACHED_MTIME
    if not MODEL_PATH.exists():
        return None
    mtime = MODEL_PATH.stat().st_mtime
    if _CACHED_PIPE is None or mtime > _CACHED_MTIME:
        _CACHED_PIPE = joblib.load(MODEL_PATH)
        _CACHED_MTIME = mtime
    return _CACHED_PIPE


def is_trained() -> bool:
    return MODEL_PATH.exists()


def load_meta() -> dict | None:
    if not META_PATH.exists():
        return None
    try:
        return json.loads(META_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None


def predict(features: dict) -> dict:
    """Devolve {'tipo': str|None, 'confidence': float, 'probs': {tipo: p}}."""
    pipe = _load_pipe()
    if pipe is None:
        return {"tipo": None, "confidence": 0.0, "probs": {c: 0.0 for c in CLASSES},
                "reason": "modelo nao treinado"}
    try:
        df = _to_dataframe([features])
        probs = pipe.predict_proba(df)[0]
        classes = list(pipe.classes_)
        # Converte chaves numpy.str_ -> str puro (evita problema em JSON).
        probs_map = {str(c): float(probs[i]) for i, c in enumerate(classes)}
        for c in CLASSES:
            probs_map.setdefault(c, 0.0)
        tipo = max(probs_map.items(), key=lambda kv: kv[1])[0]
        return {"tipo": str(tipo), "confidence": probs_map[tipo], "probs": probs_map}
    except Exception as ex:
        return {"tipo": None, "confidence": 0.0, "probs": {c: 0.0 for c in CLASSES},
                "reason": f"erro: {ex}"}


def predict_many(features_list: list[dict]) -> list[dict]:
    return [predict(f) for f in features_list]


def predict_by_processo(processo: str) -> dict:
    """Procura features mais recentes do processo no JSONL e classifica.

    Retorna {'tipo', 'confidence', 'probs', 'source'} ou
    {'tipo': None, 'reason': 'sem features cacheadas'} se nunca visto.
    """
    processo = (processo or "").strip().upper()
    if not processo:
        return {"tipo": None, "confidence": 0.0, "reason": "processo vazio"}
    feats = _features.load_all()
    candidates = [f for f in feats if (f.get("processo") or "").upper() == processo]
    if not candidates:
        return {"tipo": None, "confidence": 0.0, "reason": "sem features cacheadas",
                "source": "no_cache"}
    candidates.sort(key=lambda f: f.get("ts", 0), reverse=True)
    out = predict(candidates[0])
    out["source"] = "cache"
    out["features_ts"] = candidates[0].get("ts_iso")
    return out
