"""Captura de features de cada processo aberto pelo robô.

Complementa `web/labels.py` (que grava o rótulo Gilson) com os dados que
o classificador vai usar para aprender: lista de peças, presença de
palavras-chave, secretaria, conselheiro etc.

Gravação: append-only JSONL em `web/labels/features.jsonl`. Falha
silenciosa por design — uma falha de telemetria não pode atrapalhar o
disparo do robô.

Chamado de `src/main.py:process_processo_pipeline`, logo após
`_enumerate_piece_names` e a determinação de `_forced_tipo`. Esse é o
ponto em que o processo já está aberto no e-TCM e ainda não houve
mutação.
"""
from __future__ import annotations

import json
import os
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[1]
LABELS_DIR = ROOT / "web" / "labels"
FEATURES_FILE = LABELS_DIR / "features.jsonl"
LABELS_DIR.mkdir(parents=True, exist_ok=True)

_write_lock = threading.Lock()

# Palavras-chave que ajudam a separar as 3 classes. Calculadas em cima do
# texto dos nomes das peças (que é o que o e-TCM exibe na árvore).
_KEYWORDS = {
    "has_manutap_of": ["manutap-of", "manutap of"],
    "has_manutap": ["manutap"],
    "has_ssg": ["ssg"],
    "has_oficio_ssg": ["oficio ssg", "ofício ssg"],
    "has_requerimento_dilacao": ["requerimento dila"],
    "has_dilacao_word": ["dilacao", "dilação"],
    "has_reiteracao_word": ["reiteracao", "reiteração"],
    "has_decisao": ["decisao", "decisão"],
    "has_acordao": ["acordao", "acórdão"],
    "has_despacho": ["despacho"],
    "has_encaminha": ["encaminha"],
    "has_julga": ["julga"],
    "has_aposentadoria": ["aposentadoria"],
    "has_pensao": ["pensao", "pensão"],
    "has_providencias": ["providencia", "providência"],
}


def _norm(s: str) -> str:
    """Lowercase + remove acentos para fazer match robusto."""
    import unicodedata as _u
    return "".join(c for c in _u.normalize("NFD", s or "") if _u.category(c) != "Mn").lower()


def _piece_features(pieces: list[tuple[int, str]] | list[dict] | None) -> dict[str, Any]:
    """Extrai features booleanas + contagens a partir da lista de peças.

    Aceita formato `[(numero, nome), ...]` (igual ao retornado por
    `_enumerate_piece_names` no main.py) ou lista de dicts.
    """
    out: dict[str, Any] = {"pieces_count": 0, "pieces_names": []}
    if not pieces:
        for k in _KEYWORDS:
            out[k] = False
        return out
    names: list[str] = []
    nums: list[int] = []
    for it in pieces:
        if isinstance(it, dict):
            n = it.get("name") or it.get("title") or ""
            num = it.get("num") or it.get("number")
        else:
            try:
                num, n = it[0], it[1]
            except Exception:
                continue
        names.append(str(n))
        if isinstance(num, int):
            nums.append(num)
    out["pieces_count"] = len(names)
    out["pieces_names"] = names
    if nums:
        out["pieces_num_min"] = min(nums)
        out["pieces_num_max"] = max(nums)
    big_blob = _norm(" \n ".join(names))
    for key, terms in _KEYWORDS.items():
        out[key] = any(t in big_blob for t in terms)
    # Heurística: posição da última peça MANUTAP-OF e da última SSG.
    last_manutap_of = -1
    last_ssg = -1
    for i, n in enumerate(names):
        nn = _norm(n)
        if "manutap-of" in nn or "manutap of" in nn:
            last_manutap_of = i
        if "ssg" in nn:
            last_ssg = i
    out["pos_last_manutap_of"] = last_manutap_of
    out["pos_last_ssg"] = last_ssg
    out["ssg_after_manutap"] = last_ssg > last_manutap_of >= 0
    return out


def record(
    processo: str,
    *,
    tipo_forced: str | None = None,
    pieces: list[tuple[int, str]] | list[dict] | None = None,
    piece_chosen_title: str | None = None,
    piece_chosen_number: int | str | None = None,
    secretaria: str | None = None,
    extra: dict | None = None,
    job_id: str | None = None,
) -> None:
    """Grava as features extraídas do processo. Nunca levanta exceção."""
    try:
        feat = _piece_features(pieces or [])
        entry: dict[str, Any] = {
            "ts": time.time(),
            "ts_iso": datetime.now().isoformat(timespec="seconds"),
            "processo": (processo or "").strip().upper(),
            "tipo_forced": (tipo_forced or "").strip().upper() or None,
            "piece_chosen_title": piece_chosen_title,
            "piece_chosen_number": piece_chosen_number,
            "secretaria": secretaria,
            "job_id": job_id or os.getenv("EUCLIDES_JOB_ID"),
            **feat,
        }
        if extra:
            entry["extra"] = extra
        with _write_lock:
            with FEATURES_FILE.open("a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        # Telemetria nunca derruba o robô.
        pass


def load_all() -> list[dict]:
    if not FEATURES_FILE.exists():
        return []
    out: list[dict] = []
    with FEATURES_FILE.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return out


def stats() -> dict:
    rows = load_all()
    if not rows:
        return {"total": 0, "with_pieces": 0, "avg_pieces": 0.0}
    with_pieces = sum(1 for r in rows if r.get("pieces_count", 0) > 0)
    avg = sum(r.get("pieces_count", 0) for r in rows) / len(rows)
    by_tipo: dict[str, int] = {}
    for r in rows:
        t = r.get("tipo_forced") or "UNKNOWN"
        by_tipo[t] = by_tipo.get(t, 0) + 1
    return {
        "total": len(rows),
        "with_pieces": with_pieces,
        "avg_pieces": round(avg, 1),
        "by_tipo": by_tipo,
    }


def join_with_labels() -> list[dict]:
    """JOIN simples por número do processo entre features e labels.

    Útil para treino offline: retorna apenas processos que aparecem nas
    duas tabelas.
    """
    from . import labels as _labels
    feats = {r["processo"]: r for r in load_all() if r.get("processo")}
    out: list[dict] = []
    for lab in _labels.load_all():
        p = lab.get("processo")
        if p and p in feats:
            merged = {**feats[p], **{"label_tipo": lab.get("tipo"),
                                     "label_ts": lab.get("ts"),
                                     "label_source": lab.get("source")}}
            out.append(merged)
    return out
