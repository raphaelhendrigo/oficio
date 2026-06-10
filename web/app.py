"""FastAPI app — interface web para o Gilson disparar o robô APO-PEN.

Fluxo:
    GET  /              página de upload
    POST /upload        recebe o Excel, faz parse, redireciona p/ confirmação
    GET  /confirm/{id}  lista processos com dropdown de tipo
    POST /confirm/{id}  registra tipos e cria job
    GET  /run/{id}      página de acompanhamento
    WS   /run/{id}/ws   stream de log + status
    GET  /api/job/{id}  snapshot JSON do job
"""
from __future__ import annotations

import asyncio
import json
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import excel_parser, features as features_mod, labels, report, runner, store

ROOT = Path(__file__).resolve().parents[1]
WEB_DIR = ROOT / "web"
UPLOADS_DIR = WEB_DIR / "uploads"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Euclides — disparador APO-PEN", version="1.0")
templates = Jinja2Templates(directory=str(WEB_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")

# Cache temporário do upload (antes de virar Job).
_pending_uploads: dict[str, dict] = {}


def _next_business_day_0010() -> datetime:
    """Retorna 00:10 do próximo dia útil (segunda a sexta, sem feriados).

    Default sugerido para o agendamento — bate com o horário usado nos wrappers
    PowerShell existentes.
    """
    now = datetime.now()
    candidate = now.replace(hour=0, minute=10, second=0, microsecond=0) + timedelta(days=1)
    while candidate.weekday() >= 5:  # 5=sábado, 6=domingo
        candidate += timedelta(days=1)
    return candidate


def _parse_scheduled_at(raw: str | None) -> float | None:
    """Converte o valor do input datetime-local em timestamp epoch local."""
    if not raw:
        return None
    raw = raw.strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt).timestamp()
        except ValueError:
            continue
    return None


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    recent = store.list_recent(10)
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "recent": recent, "is_busy": store.is_busy()},
    )


@app.post("/upload")
async def upload(file: Annotated[UploadFile, File()]):
    name = file.filename or "planilha.xlsx"
    if not (name.lower().endswith(".xlsx") or name.lower().endswith(".xls")):
        raise HTTPException(400, "Envie um arquivo .xlsx ou .xls")
    data = await file.read()
    if not data:
        raise HTTPException(400, "Arquivo vazio")
    try:
        processos = excel_parser.parse_processos(name, data)
    except Exception as ex:
        raise HTTPException(400, f"Falha ao ler planilha: {ex}")
    if not processos:
        raise HTTPException(400, "Nenhum processo TC encontrado na planilha")

    ts = time.strftime("%Y%m%d_%H%M%S")
    safe_name = name.replace("/", "_").replace("\\", "_")
    dest = UPLOADS_DIR / f"{ts}_{safe_name}"
    dest.write_bytes(data)

    upload_id = f"up_{ts}"
    _pending_uploads[upload_id] = {
        "filename": name,
        "saved_at": str(dest),
        "processos": processos,
    }
    return RedirectResponse(url=f"/confirm/{upload_id}", status_code=303)


@app.get("/confirm/{upload_id}", response_class=HTMLResponse)
async def confirm_get(upload_id: str, request: Request):
    info = _pending_uploads.get(upload_id)
    if not info:
        raise HTTPException(404, "Upload expirou — envie a planilha de novo.")
    default_dt = _next_business_day_0010()
    return templates.TemplateResponse(
        "confirm.html",
        {
            "request": request,
            "upload_id": upload_id,
            "filename": info["filename"],
            "processos": info["processos"],
            "is_busy": store.is_busy(),
            "default_scheduled_at": default_dt.strftime("%Y-%m-%dT%H:%M"),
            "default_scheduled_label": default_dt.strftime("%d/%m/%Y às %H:%M"),
        },
    )


@app.post("/confirm/{upload_id}")
async def confirm_post(upload_id: str, request: Request):
    info = _pending_uploads.get(upload_id)
    if not info:
        raise HTTPException(404, "Upload expirou — envie a planilha de novo.")
    form = await request.form()

    selected: list[dict] = []
    skipped: list[str] = []
    for processo in info["processos"]:
        key = f"tipo__{processo}"
        tipo = (form.get(key) or "").strip().upper()
        if tipo == "PULAR" or not tipo:
            skipped.append(processo)
            continue
        if tipo not in ("UTAP", "DILACAO", "REITERACAO"):
            raise HTTPException(400, f"Tipo inválido para {processo}: {tipo}")
        selected.append({"processo": processo, "tipo": tipo})

    if not selected:
        raise HTTPException(400, "Nenhum processo selecionado para disparo.")

    when_mode = (form.get("when_mode") or "scheduled").strip()
    scheduled_at: float | None = None
    if when_mode == "scheduled":
        raw = str(form.get("scheduled_at") or "")
        scheduled_at = _parse_scheduled_at(raw)
        if scheduled_at is None:
            scheduled_at = _next_business_day_0010().timestamp()

    # Se for imediato, exige que não haja outro job rodando agora.
    if when_mode != "scheduled" and store.is_busy():
        raise HTTPException(409, "Já existe um job em execução. Aguarde ou agende.")

    job = store.new_job(
        excel_filename=info["filename"],
        processos=selected,
        scheduled_at=scheduled_at,
    )
    # Captura supervisionada (additivo, não interfere no fluxo):
    # cada decisão de tipo do Gilson vira uma linha em web/labels/labels.jsonl
    # para treinar o classificador automático mais à frente.
    labels.record_many(selected, job_id=job.id, excel_filename=info["filename"])
    for p in skipped:
        store.append_log(job.id, f"[skip] {p} marcado como PULAR")

    # Limpa cache de upload — daqui pra frente o job tem identidade própria.
    _pending_uploads.pop(upload_id, None)

    # Dispara em background.
    asyncio.create_task(runner.run_job(job.id))

    return RedirectResponse(url=f"/run/{job.id}", status_code=303)


@app.post("/run/{job_id}/cancel")
async def cancel_run(job_id: str):
    ok = store.cancel_job(job_id)
    if not ok:
        raise HTTPException(409, "Não foi possível cancelar (job já rodando ou finalizado).")
    return RedirectResponse(url=f"/run/{job_id}", status_code=303)


@app.get("/run/{job_id}", response_class=HTMLResponse)
async def run_page(job_id: str, request: Request):
    job = store.get_job(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado")
    return templates.TemplateResponse(
        "run.html",
        {"request": request, "job": job, "job_dict": job.to_dict(), "is_busy": store.is_busy()},
    )


@app.get("/api/job/{job_id}")
async def api_job(job_id: str):
    job = store.get_job(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado")
    return JSONResponse(job.to_dict())


@app.get("/labels", response_class=HTMLResponse)
async def labels_page(request: Request):
    s = labels.stats()
    fs = features_mod.stats()
    joined = features_mod.join_with_labels()
    # Estimativa de quantos dias úteis faltam para cada marco com base no
    # ritmo histórico (avg_per_day). Se ainda não há histórico, mostra "—".
    rate = s["avg_per_day"] or 0
    eta: dict[str, str] = {}
    for name, target in s["milestones"].items():
        if rate <= 0:
            eta[name] = "—"
        elif s["total"] >= target:
            eta[name] = "atingido"
        else:
            remaining = target - s["total"]
            days = remaining / rate
            eta[name] = f"~{int(round(days))} dias úteis ({int(remaining)} labels)"
    return templates.TemplateResponse(
        "labels.html",
        {
            "request": request,
            "stats": s,
            "eta": eta,
            "features_stats": fs,
            "joined_count": len(joined),
            "is_busy": store.is_busy(),
        },
    )


@app.get("/api/labels/stats")
async def api_labels_stats():
    return JSONResponse({
        "labels": labels.stats(),
        "features": features_mod.stats(),
        "joined": len(features_mod.join_with_labels()),
    })


@app.get("/run/{job_id}/report", response_class=HTMLResponse)
async def report_html(job_id: str, request: Request):
    job = store.get_job(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado")
    summary = report.summarize(job)
    return templates.TemplateResponse(
        "report.html",
        {
            "request": request,
            "job": job,
            "summary": summary,
            "is_busy": store.is_busy(),
            "status_label": report.STATUS_LABEL,
        },
    )


@app.get("/run/{job_id}/report.docx")
async def report_docx(job_id: str):
    job = store.get_job(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado")
    blob = report.render_docx(job)
    return Response(
        content=blob,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="euclides_relatorio_{job.id}.docx"'},
    )


@app.get("/run/{job_id}/report.pdf")
async def report_pdf(job_id: str):
    job = store.get_job(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado")
    blob = report.render_pdf(job)
    return Response(
        content=blob,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="euclides_relatorio_{job.id}.pdf"'},
    )


@app.websocket("/run/{job_id}/ws")
async def ws_run(websocket: WebSocket, job_id: str):
    await websocket.accept()
    job = store.get_job(job_id)
    if not job:
        await websocket.send_text(json.dumps({"type": "error", "message": "job não encontrado"}))
        await websocket.close()
        return

    sent_lines = 0
    last_status_snapshot = None
    try:
        while True:
            j = store.get_job(job_id)
            if not j:
                break
            # Envia novas linhas de log.
            new = j.log_lines[sent_lines:]
            if new:
                for ln in new:
                    await websocket.send_text(json.dumps({"type": "log", "line": ln}))
                sent_lines = len(j.log_lines)
            # Envia snapshot de status se mudou.
            snap = {
                "status": j.status,
                "current_tipo": j.current_tipo,
                "processos": [
                    {"processo": p.processo, "tipo": p.tipo, "status": p.status, "error": p.error}
                    for p in j.processos
                ],
            }
            if snap != last_status_snapshot:
                await websocket.send_text(json.dumps({"type": "status", **snap}))
                last_status_snapshot = snap
            if j.status in ("done", "failed"):
                await websocket.send_text(json.dumps({"type": "finished", "status": j.status, "error": j.error}))
                break
            await asyncio.sleep(0.5)
    except WebSocketDisconnect:
        return
    except Exception as ex:
        try:
            await websocket.send_text(json.dumps({"type": "error", "message": str(ex)}))
        except Exception:
            pass
    finally:
        try:
            await websocket.close()
        except Exception:
            pass
