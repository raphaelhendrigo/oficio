"""Estado em memória dos peeks (pré-classificações) em execução.

Modelo simples, mesmo padrão do web/store.py (jobs do robô principal), mas
mais leve: peek não persiste em disco, só vive enquanto a aba está aberta
ou termina.
"""
from __future__ import annotations

import threading
import time
import uuid
from dataclasses import dataclass, field, asdict

_lock = threading.Lock()
_peeks: dict[str, "PeekJob"] = {}


@dataclass
class PeekJob:
    id: str
    started_at: float
    processos: list[str]
    total: int
    done: int = 0
    status: str = "running"           # running | done | failed
    results: dict = field(default_factory=dict)   # processo -> prediction dict
    log_lines: list[str] = field(default_factory=list)
    finished_at: float | None = None
    error: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)


def new_peek(processos: list[str]) -> PeekJob:
    pj = PeekJob(
        id=uuid.uuid4().hex[:12],
        started_at=time.time(),
        processos=list(processos),
        total=len(processos),
    )
    with _lock:
        _peeks[pj.id] = pj
    return pj


def get(peek_id: str) -> PeekJob | None:
    with _lock:
        return _peeks.get(peek_id)


def append_log(peek_id: str, line: str) -> None:
    with _lock:
        pj = _peeks.get(peek_id)
        if not pj:
            return
        pj.log_lines.append(line)
        if len(pj.log_lines) > 2000:
            pj.log_lines = pj.log_lines[-2000:]


def record_result(peek_id: str, processo: str, prediction: dict) -> None:
    with _lock:
        pj = _peeks.get(peek_id)
        if not pj:
            return
        pj.results[processo] = prediction
        pj.done = len(pj.results)


def mark_finished(peek_id: str, ok: bool, error: str | None = None) -> None:
    with _lock:
        pj = _peeks.get(peek_id)
        if not pj:
            return
        pj.status = "done" if ok else "failed"
        pj.finished_at = time.time()
        if error:
            pj.error = error


def prune_old(max_age_seconds: int = 3600) -> int:
    """Remove peeks finalizados há mais de X. Chamado oportunisticamente."""
    now = time.time()
    removed = 0
    with _lock:
        for pid in list(_peeks.keys()):
            pj = _peeks[pid]
            if pj.finished_at and (now - pj.finished_at) > max_age_seconds:
                del _peeks[pid]
                removed += 1
    return removed
