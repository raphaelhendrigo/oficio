"""Armazena estado dos jobs em memória + persiste cada job em disco para auditoria.

Roda 1 job por vez (lock global) — disparar 2 simultâneos no mesmo robô
geraria conflito no navegador Playwright.
"""
from __future__ import annotations

import json
import threading
import time
import uuid
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Literal

ROOT = Path(__file__).resolve().parents[1]
JOBS_DIR = ROOT / "web" / "jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)

Tipo = Literal["UTAP", "DILACAO", "REITERACAO"]
JobStatus = Literal["pending", "scheduled", "running", "done", "failed", "cancelled"]
ProcStatus = Literal["pending", "running", "ok", "failed", "skipped"]


@dataclass
class ProcessoStatus:
    processo: str
    tipo: Tipo
    status: ProcStatus = "pending"
    attempts: int = 0
    error: str | None = None


@dataclass
class Job:
    id: str
    created_at: float
    excel_filename: str
    processos: list[ProcessoStatus]
    status: JobStatus = "pending"
    current_tipo: Tipo | None = None
    log_lines: list[str] = field(default_factory=list)
    scheduled_at: float | None = None
    started_at: float | None = None
    finished_at: float | None = None
    error: str | None = None

    def to_dict(self) -> dict:
        d = asdict(self)
        return d


_lock = threading.Lock()
_jobs: dict[str, Job] = {}
_running_job_id: str | None = None


def _job_from_dict(data: dict) -> Job:
    processos = [
        ProcessoStatus(
            processo=p.get("processo", ""),
            tipo=p.get("tipo", "UTAP"),
            status=p.get("status", "pending"),
            attempts=int(p.get("attempts") or 0),
            error=p.get("error"),
        )
        for p in data.get("processos", [])
    ]
    return Job(
        id=data.get("id", ""),
        created_at=float(data.get("created_at") or 0),
        excel_filename=data.get("excel_filename", ""),
        processos=processos,
        status=data.get("status", "pending"),
        current_tipo=data.get("current_tipo"),
        log_lines=list(data.get("log_lines") or []),
        scheduled_at=data.get("scheduled_at"),
        started_at=data.get("started_at"),
        finished_at=data.get("finished_at"),
        error=data.get("error"),
    )


def _load_job_from_disk(job_id: str) -> Job | None:
    path = JOBS_DIR / f"{job_id}.json"
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        job = _job_from_dict(data)
        if job.id != job_id:
            return None
        return job
    except Exception:
        return None


def new_job(excel_filename: str, processos: list[dict], scheduled_at: float | None = None) -> Job:
    """processos = [{"processo": "TC/000123/2024", "tipo": "REITERACAO"}, ...]"""
    job_id = uuid.uuid4().hex[:12]
    plist = [
        ProcessoStatus(processo=p["processo"], tipo=p["tipo"])  # type: ignore[arg-type]
        for p in processos
    ]
    initial_status: JobStatus = "scheduled" if scheduled_at and scheduled_at > time.time() else "pending"
    job = Job(
        id=job_id,
        created_at=time.time(),
        excel_filename=excel_filename,
        processos=plist,
        scheduled_at=scheduled_at,
        status=initial_status,
    )
    with _lock:
        _jobs[job_id] = job
    persist(job)
    return job


def mark_scheduled(job_id: str) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        if job.status == "pending":
            job.status = "scheduled"
    persist(job)


def cancel_job(job_id: str) -> bool:
    """Só cancela jobs ainda agendados (não pode interromper um já rodando)."""
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return False
        if job.status not in ("scheduled", "pending"):
            return False
        job.status = "cancelled"
        job.finished_at = time.time()
    persist(job)
    return True


def get_job(job_id: str) -> Job | None:
    with _lock:
        job = _jobs.get(job_id)
        if job:
            return job
    job = _load_job_from_disk(job_id)
    if job:
        with _lock:
            _jobs[job_id] = job
    return job


def mark_running(job_id: str) -> bool:
    """Tenta marcar o job como rodando. Retorna False se já existe outro rodando."""
    global _running_job_id
    with _lock:
        if _running_job_id and _running_job_id != job_id:
            running = _jobs.get(_running_job_id)
            if running and running.status == "running":
                return False
        job = _jobs.get(job_id)
        if not job:
            return False
        job.status = "running"
        job.started_at = time.time()
        _running_job_id = job_id
    persist(job)
    return True


def mark_finished(job_id: str, ok: bool, error: str | None = None) -> None:
    global _running_job_id
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        job.status = "done" if ok else "failed"
        job.finished_at = time.time()
        job.error = error
        if _running_job_id == job_id:
            _running_job_id = None
    persist(job)


def update_processo(job_id: str, processo: str, **kwargs) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        for p in job.processos:
            if p.processo == processo:
                for k, v in kwargs.items():
                    setattr(p, k, v)
                break
    persist(job)


def set_current_tipo(job_id: str, tipo: Tipo | None) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        job.current_tipo = tipo
    persist(job)


def append_log(job_id: str, line: str) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        job.log_lines.append(line)
        if len(job.log_lines) > 5000:
            job.log_lines = job.log_lines[-5000:]
    persist(job)


def persist(job: Job) -> None:
    path = JOBS_DIR / f"{job.id}.json"
    try:
        path.write_text(json.dumps(job.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def list_recent(limit: int = 20) -> list[Job]:
    try:
        for path in JOBS_DIR.glob("*.json"):
            job_id = path.stem
            with _lock:
                already = job_id in _jobs
            if already:
                continue
            job = _load_job_from_disk(job_id)
            if job:
                with _lock:
                    _jobs[job_id] = job
    except Exception:
        pass
    with _lock:
        items = sorted(_jobs.values(), key=lambda j: j.created_at, reverse=True)
        return items[:limit]


def is_busy() -> bool:
    with _lock:
        return _running_job_id is not None
