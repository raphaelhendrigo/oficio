"""Captura supervisionada das classificações que o Gilson faz.

Cada vez que o Gilson confirma o tipo de um processo (UTAP / DILACAO /
REITERACAO) na tela `/confirm/...`, gravamos um registro append-only em
`web/labels/labels.jsonl`. Esse arquivo vai virar o dataset de treino do
classificador automático.

Decisão de design: JSONL puro (sem banco) — append atômico, fácil de
auditar, fácil de importar em pandas/scikit-learn depois. Quando o volume
passar de ~10k linhas, migrar para SQLite.

Esse módulo é *additivo* e não interfere no fluxo de execução do robô.
Mesmo se a gravação falhar, o disparo continua.
"""
from __future__ import annotations

import json
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
LABELS_DIR = ROOT / "web" / "labels"
LABELS_FILE = LABELS_DIR / "labels.jsonl"
LABELS_DIR.mkdir(parents=True, exist_ok=True)

# Lock só para serializar writes — leituras são sempre offline.
_write_lock = threading.Lock()


def record(
    processo: str,
    tipo: str,
    *,
    source: str = "gilson_web_confirm",
    job_id: str | None = None,
    excel_filename: str | None = None,
    user: str | None = None,
) -> None:
    """Grava uma decisão de classificação. Nunca levanta — falha silenciosa.

    Cada linha contém: timestamp, ISO local, processo, tipo, origem e
    metadados úteis para rastreabilidade. As *features* do processo (último
    ato, secretaria etc.) virão num segundo passo, capturadas pelo robô na
    hora que ele abrir o processo no e-TCM — esse módulo só registra o
    *label*.
    """
    try:
        entry = {
            "ts": time.time(),
            "ts_iso": datetime.now().isoformat(timespec="seconds"),
            "processo": (processo or "").strip().upper(),
            "tipo": (tipo or "").strip().upper(),
            "source": source,
            "job_id": job_id,
            "excel_filename": excel_filename,
            "user": user,
        }
        with _write_lock:
            with LABELS_FILE.open("a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        # Falha silenciosa: NÃO deixa um erro de telemetria atrapalhar o
        # disparo do robô. O usuário não pode perder uma janela operacional
        # porque uma gravação opcional falhou.
        pass


def record_many(rows: Iterable[dict], *, source: str = "gilson_web_confirm",
                job_id: str | None = None, excel_filename: str | None = None) -> int:
    """Grava várias decisões num único flush. Retorna nº de linhas escritas."""
    written = 0
    try:
        now = time.time()
        iso = datetime.now().isoformat(timespec="seconds")
        with _write_lock:
            with LABELS_FILE.open("a", encoding="utf-8") as f:
                for row in rows:
                    entry = {
                        "ts": now,
                        "ts_iso": iso,
                        "processo": (row.get("processo") or "").strip().upper(),
                        "tipo": (row.get("tipo") or "").strip().upper(),
                        "source": source,
                        "job_id": job_id,
                        "excel_filename": excel_filename,
                        "user": row.get("user"),
                    }
                    f.write(json.dumps(entry, ensure_ascii=False) + "\n")
                    written += 1
    except Exception:
        pass
    return written


def load_all() -> list[dict]:
    """Lê o JSONL inteiro em memória (para stats e treino offline)."""
    if not LABELS_FILE.exists():
        return []
    out: list[dict] = []
    with LABELS_FILE.open("r", encoding="utf-8") as f:
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
    """Retorna contagens agregadas para a tela de progresso."""
    rows = load_all()
    by_tipo: dict[str, int] = {"UTAP": 0, "DILACAO": 0, "REITERACAO": 0}
    by_day: dict[str, int] = {}
    unique_processos: set[str] = set()
    first_ts: float | None = None
    last_ts: float | None = None
    for r in rows:
        t = r.get("tipo")
        if t in by_tipo:
            by_tipo[t] += 1
        proc = r.get("processo")
        if proc:
            unique_processos.add(proc)
        ts = r.get("ts")
        if isinstance(ts, (int, float)):
            day = datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
            by_day[day] = by_day.get(day, 0) + 1
            if first_ts is None or ts < first_ts:
                first_ts = ts
            if last_ts is None or ts > last_ts:
                last_ts = ts
    total = sum(by_tipo.values())
    days_active = len(by_day)
    avg_per_day = (total / days_active) if days_active else 0
    return {
        "total": total,
        "unique_processos": len(unique_processos),
        "by_tipo": by_tipo,
        "by_day": dict(sorted(by_day.items())),
        "days_active": days_active,
        "avg_per_day": round(avg_per_day, 1),
        "first_ts": first_ts,
        "last_ts": last_ts,
        # Marcos de treino (estimativa em labels totais):
        "milestones": {
            "sinal_minimo": 100,
            "baseline_assistido": 400,
            "auto_confiavel": 1200,
            "substituicao_completa": 2500,
        },
    }
