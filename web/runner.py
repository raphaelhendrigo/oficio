"""Orquestrador: dado um Job, dispara src/main.py em sub-batches por tipo.

Para cada tipo (UTAP / DILACAO / REITERACAO) com processos selecionados:
  - monta env: FORCE_TIPO, PROCESSOS_LIST, ONLY_PROCESSOS_AUTHORIZED
  - inicia python subprocess
  - faz streaming de stdout/stderr linha a linha, anexa no log do job
  - parseia eventos do robô para atualizar status por processo
  - aguarda finalizar antes do próximo tipo (subprocess separado evita vazamento
    de env entre tipos, como o wrapper PowerShell já fazia)
"""
from __future__ import annotations

import asyncio
import os
import re
import sys
from datetime import date
from pathlib import Path
from typing import Iterable

from . import signers as _signers
from . import store

ROOT = Path(__file__).resolve().parents[1]
PY = ROOT / ".venv" / "Scripts" / "python.exe"
MAIN_PY = ROOT / "src" / "main.py"

RE_INICIANDO = re.compile(r"Iniciando pipeline do processo (TC/\d+/\d+)")
RE_ASSINATURA_OK = re.compile(r"Assinatura solicitada para .* no processo (TC/\d+/\d+)")
RE_OFICIO_CONCLUIDO = re.compile(r"Of[ií]cio SSG conclu[ií]do.*?(TC/\d+/\d+)")
RE_FALHA = re.compile(r"(?:ERRO|FAIL|Falha|fora de ONLY_PROCESSOS).*?(TC/\d+/\d+)")


def _default_env_for_tipo(tipo: str) -> dict[str, str]:
    """Env vars base copiadas dos runners PowerShell do projeto.

    Comportamento PROD/destrutivo + assinatura sem tramitar.

    Os campos de assinante (`ASSINANTE_NOME`, `SIGNER_MATCH_TOKENS`,
    pastas `OFICIO_TEMPLATES_DIR_<TIPO>`, fallback de secretaria) vem do
    assinante ATIVO em `web/signers.py:active()`. O usuário escolhe na
    home — default = Daniela Shimizu (durante as ferias da Roseli).
    """
    tipo = tipo.strip().upper()
    data_oficio = date.today().strftime("%d/%m/%Y")
    env = {
        "ETCM_URL": "https://etcm.tcm.sp.gov.br/paginas/login.aspx",
        "BASE_URL": "https://etcm.tcm.sp.gov.br/paginas/login.aspx",
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
        "ENVIRONMENT": "producao",
        # Headless puro: nenhum Chrome visivel na VM (versao web 24/7).
        "HEADLESS": "true",
        "SHOW_BROWSER": "false",
        "WATCH_MODE": "false",
        "DEVTOOLS": "false",
        "SLOWMO_MS": "0",
        # Login automatico via env vars ETCM_USERNAME/PASSWORD - nao precisa
        # esperar interacao humana. 15s e suficiente para a tela carregar.
        "LOGIN_MANUAL_WAIT_MS": "15000",
        "PAUSE_AFTER_LOGIN_MS": "1500",
        "USE_CAIXA_CORREIO": "true",
        "SAFE_DELETE_OWN_DRAFTS": "true",
        "RUN_PROD_DESTRUCTIVE_CLEANUP": "true",
        "FORCE_DELETE_OLD_OFICIO_SSG": "true",
        "FORCE_RECREATE_COMUNICACAO": "true",
        # Env var generica que controla estorno de assinatura pendente do
        # assinante configurado em SIGNER_MATCH_TOKENS (Roseli, Daniela ou
        # custom). Mantemos o nome historico p/ nao quebrar runners antigos.
        "FORCE_REVOKE_PENDING_ROSELI_SIGNATURE": "true",
        "FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED": "false",
        "SKIP_COMUNICACAO_CLEANUP": "false",
        "REUSE_EXISTING_OFICIO": "false",
        "STOP_AFTER_OFICIO_CONCLUIDO": "false",
        "SKIP_SIGNATURE": "false",
        "REQUEST_SIGNATURE": "true",
        "SKIP_TRAMITACAO": "true",
        "TRAMITAR_DESTINO": "",
        "DISTRIBUIR_PARA": "",
        "COMUNICACAO_PRAZO_DIAS": "60",
        "COMUNICACAO_REFERENCIA": "gerado automaticamente",
        "STATUS_ENTREGA": "Normal",
        "OFICIO_TEMPLATE_MODE": "auto",
        "OFICIO_PRESERVE_AT_TOKENS": "true",
        "OFICIO_ADD_EUCLIDES_MARKER": "true",
        "OFICIO_EUCLIDES_MARKER": "/euclides",
        "OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA": "true" if tipo in {"UTAP", "REITERACAO"} else "false",
        "OFICIO_ENCAMINHA_TEXT_BOLD": "false",
        "DATA_OFICIO": data_oficio,
        "OFICIO_DATA": data_oficio,
        "MAX_PROCESSOS": "0",
        "FORCE_TIPO": tipo,
    }
    if tipo == "REITERACAO":
        env["REITERACAO_REQUIRE_AUTO_EXTRACT"] = "true"
    # Injeta os campos do assinante ATIVO (escolhido pelo Gilson na home).
    # Sobrescreve qualquer default — se vier vazio o matcher cai no fallback
    # default de oficio_normalize (Roseli/Chaves) sem quebrar.
    env.update(_signers.active().to_env())
    return env


def _filter_log_line(line: str) -> str:
    """Remove caracteres ANSI/controle perigosos para não quebrar o front."""
    line = line.replace("\x1b", "")
    return line.rstrip()


async def _stream_subprocess(job_id: str, cmd: list[str], env: dict[str, str], processos: list[str]) -> int:
    """Roda um subprocesso e streama log linha-a-linha.

    Implementacao Windows-friendly: le bytes em chunks (nao bloqueia em
    readline quando stdout esta full-buffered no child). Quebra em linhas
    com base em \n e \r. Forca PYTHONUNBUFFERED + sys.stdout reconfigure
    via -u no python.
    """
    import subprocess as _sp
    full_env = os.environ.copy()
    full_env.update(env)
    full_env["PYTHONUNBUFFERED"] = "1"

    # Insere -u logo apos o python.exe para forcar unbuffered stdio.
    cmd_unbuffered = list(cmd)
    if cmd_unbuffered and cmd_unbuffered[0].lower().endswith("python.exe"):
        cmd_unbuffered.insert(1, "-u")

    store.append_log(job_id, f"$ {' '.join(cmd_unbuffered)}")
    store.append_log(job_id, f"  FORCE_TIPO={env.get('FORCE_TIPO')}  PROCESSOS={env.get('PROCESSOS_LIST')}")

    creationflags = 0
    if hasattr(_sp, "CREATE_NO_WINDOW"):
        creationflags = _sp.CREATE_NO_WINDOW  # esconde console do child

    proc = await asyncio.create_subprocess_exec(
        *cmd_unbuffered,
        cwd=str(ROOT),
        env=full_env,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
        creationflags=creationflags,
    )

    assert proc.stdout is not None
    buf = bytearray()
    while True:
        chunk = await proc.stdout.read(2048)
        if not chunk:
            break
        buf.extend(chunk)
        # Quebra por \n; tambem trata \r como flush parcial.
        while True:
            nl = buf.find(b"\n")
            if nl < 0:
                # Se buffer ja tem muita coisa sem newline, flusha como linha.
                if len(buf) > 4096:
                    line = bytes(buf).decode("utf-8", errors="replace")
                    buf.clear()
                    line = _filter_log_line(line)
                    if line:
                        store.append_log(job_id, line)
                        _scan_for_events(job_id, line, processos)
                break
            raw_line = bytes(buf[:nl])
            del buf[:nl + 1]
            try:
                line = raw_line.decode("utf-8", errors="replace")
            except Exception:
                line = repr(raw_line)
            line = _filter_log_line(line)
            if not line:
                continue
            store.append_log(job_id, line)
            _scan_for_events(job_id, line, processos)

    # Flush final.
    if buf:
        line = bytes(buf).decode("utf-8", errors="replace").rstrip()
        if line:
            store.append_log(job_id, line)
            _scan_for_events(job_id, line, processos)

    rc = await proc.wait()
    store.append_log(job_id, f"[exit={rc}]")
    return rc


def _scan_for_events(job_id: str, line: str, processos: list[str]) -> None:
    m = RE_INICIANDO.search(line)
    if m and m.group(1) in processos:
        store.update_processo(job_id, m.group(1), status="running")
        return
    m = RE_ASSINATURA_OK.search(line)
    if m and m.group(1) in processos:
        store.update_processo(job_id, m.group(1), status="ok")
        return
    m = RE_OFICIO_CONCLUIDO.search(line)
    if m and m.group(1) in processos:
        store.update_processo(job_id, m.group(1), status="running")
        return
    m = RE_FALHA.search(line)
    if m and m.group(1) in processos:
        store.update_processo(job_id, m.group(1), status="failed", error=line[:300])
        return


async def _wait_until_scheduled(job_id: str) -> bool:
    """Aguarda até job.scheduled_at, anotando countdown no log a cada 60s.

    Retorna False se o job foi cancelado durante a espera, True caso contrário.
    """
    import time as _t
    while True:
        j = store.get_job(job_id)
        if not j:
            return False
        if j.status == "cancelled":
            store.append_log(job_id, "[CANCELADO] Job cancelado antes do horário agendado.")
            return False
        if not j.scheduled_at:
            return True
        remaining = j.scheduled_at - _t.time()
        if remaining <= 0:
            return True
        # Log periódico (no máximo 1 a cada 60s).
        if remaining > 60 and int(remaining) % 60 == 0:
            mins = int(remaining // 60)
            hrs = mins // 60
            store.append_log(job_id, f"[agendado] iniciar em {hrs:02d}h{(mins%60):02d}m...")
        await asyncio.sleep(min(remaining, 10))


async def run_job(job_id: str) -> None:
    """Roda todos os sub-batches em sequência. Marca status final no fim."""
    job = store.get_job(job_id)
    if not job:
        return
    # Se houver scheduled_at no futuro, aguarda primeiro.
    if job.scheduled_at and job.scheduled_at > __import__("time").time():
        store.mark_scheduled(job_id)
        from datetime import datetime as _dt
        when = _dt.fromtimestamp(job.scheduled_at).strftime("%d/%m/%Y %H:%M")
        store.append_log(job_id, f"[agendado] disparo programado para {when}")
        ok = await _wait_until_scheduled(job_id)
        if not ok:
            store.mark_finished(job_id, ok=False, error="cancelado antes da execução")
            return
    if not store.mark_running(job_id):
        store.append_log(job_id, "[ABORTADO] Já existe outro job rodando.")
        store.mark_finished(job_id, ok=False, error="outro job em execução")
        return

    overall_ok = True
    error: str | None = None
    try:
        if not PY.exists():
            raise RuntimeError(f".venv não encontrado em {PY}")
        if not MAIN_PY.exists():
            raise RuntimeError(f"src/main.py não encontrado em {MAIN_PY}")

        by_tipo: dict[str, list[str]] = {"UTAP": [], "DILACAO": [], "REITERACAO": []}
        for p in job.processos:
            by_tipo.setdefault(p.tipo, []).append(p.processo)

        for tipo in ("UTAP", "DILACAO", "REITERACAO"):
            lista = by_tipo.get(tipo, [])
            if not lista:
                continue
            store.set_current_tipo(job_id, tipo)  # type: ignore[arg-type]
            store.append_log(job_id, "")
            store.append_log(job_id, f"================ BATCH {tipo} — {len(lista)} processo(s) ================")
            env = _default_env_for_tipo(tipo)
            env["PROCESSOS_LIST"] = ",".join(lista)
            env["ONLY_PROCESSOS_AUTHORIZED"] = ",".join(lista)
            # Permite que o main.py marque cada feature gravada com o job_id
            # de origem, facilitando o JOIN com a tabela de labels.
            env["EUCLIDES_JOB_ID"] = job_id
            for p in lista:
                store.update_processo(job_id, p, attempts=1)
            rc = await _stream_subprocess(
                job_id,
                [str(PY), str(MAIN_PY)],
                env,
                lista,
            )
            if rc != 0:
                overall_ok = False
                error = f"batch {tipo} terminou com exit={rc}"
                store.append_log(job_id, f"[!] {error}")
            # Marca como falha qualquer processo deste batch que não tenha
            # virado "ok". No fluxo web PROD, "ok" exige a linha de assinatura,
            # não apenas "Ofício SSG concluído".
            for p in lista:
                st = next((x for x in store.get_job(job_id).processos if x.processo == p), None)  # type: ignore[union-attr]
                if st and st.status != "ok":
                    overall_ok = False
                    if error is None:
                        error = f"batch {tipo} terminou sem assinatura confirmada para {p}"
                    store.update_processo(job_id, p, status="failed", error="terminou sem evento de assinatura")

        store.set_current_tipo(job_id, None)
    except Exception as ex:
        overall_ok = False
        error = str(ex)
        store.append_log(job_id, f"[EXCEPTION] {ex}")
    finally:
        store.mark_finished(job_id, ok=overall_ok, error=error)
