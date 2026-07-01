"""Pre-classificação: abre cada processo no e-TCM em headless e enumera
as peças, SEM mudar nenhum estado (não cria comunicação, não deleta nada,
não baixa PDF).

Reutiliza as funções já existentes em `src/main.py` — login_etcm,
filter_and_open_processo, _enumerate_piece_names — para garantir paridade
com o que o robô principal faz no fluxo real.
"""
from __future__ import annotations

import os
import sys
import threading
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

# Lock global para evitar 2 sessões Playwright simultâneas na mesma VM.
_peek_lock = threading.Lock()


def _ensure_src_on_path() -> None:
    src_str = str(SRC)
    if src_str not in sys.path:
        sys.path.insert(0, src_str)


def peek_processos(processos: list[str], log_fn=print, on_processo=None) -> dict[str, dict]:
    """Para cada processo, devolve {processo: {pieces: [(num,name)], error?}}.

    Faz UM login só e percorre todos os processos na mesma sessão.
    Tudo em headless verdadeiro, sem destrutivo.

    Sequencia:
      1. Login
      2. Abre a tela APO-PEN
      3. Para cada processo:
          a. filter_and_open_processo
          b. _enumerate_piece_names
          c. fecha abas extras
      4. Encerra
    """
    _ensure_src_on_path()
    from playwright.sync_api import sync_playwright  # type: ignore
    import main as _main  # type: ignore # noqa: E402

    out: dict[str, dict] = {p: {"pieces": [], "error": None} for p in processos}

    # Configura env para sessão peek — headless, sem destrutivo, sem nada
    # que possa alterar o e-TCM.
    saved_env = {
        k: os.environ.get(k)
        for k in (
            "ETCM_URL", "BASE_URL",
            "HEADLESS", "SHOW_BROWSER", "WATCH_MODE", "DEVTOOLS",
            "SLOWMO_MS", "RUN_PROD_DESTRUCTIVE_CLEANUP",
            "FORCE_DELETE_OLD_OFICIO_SSG", "FORCE_RECREATE_COMUNICACAO",
            "STOP_AFTER_OFICIO_CONCLUIDO", "SKIP_SIGNATURE",
            "REQUEST_SIGNATURE", "SKIP_TRAMITACAO",
            "PROCESSOS_LIST", "ONLY_PROCESSOS_AUTHORIZED",
        )
    }
    # URL base usada pelo _go_to_mesa_trabalho e amigos (precisa estar no env,
    # nao so como variavel local) — o web service so promove credenciais.
    _etcm_url_for_main = (
        os.environ.get("ETCM_URL")
        or os.environ.get("BASE_URL")
        or "https://etcm.tcm.sp.gov.br/paginas/login.aspx"
    )
    os.environ["ETCM_URL"] = _etcm_url_for_main
    os.environ["BASE_URL"] = _etcm_url_for_main
    os.environ["HEADLESS"] = "true"
    os.environ["SHOW_BROWSER"] = "false"
    os.environ["WATCH_MODE"] = "false"
    os.environ["DEVTOOLS"] = "false"
    os.environ["SLOWMO_MS"] = "0"
    # Garante que NADA destrutivo rode mesmo se algum guard vazar:
    os.environ["RUN_PROD_DESTRUCTIVE_CLEANUP"] = "false"
    os.environ["FORCE_DELETE_OLD_OFICIO_SSG"] = "false"
    os.environ["FORCE_RECREATE_COMUNICACAO"] = "false"
    os.environ["SKIP_SIGNATURE"] = "true"
    os.environ["REQUEST_SIGNATURE"] = "false"
    os.environ["SKIP_TRAMITACAO"] = "true"
    os.environ["STOP_AFTER_OFICIO_CONCLUIDO"] = "true"
    # ONLY_PROCESSOS_AUTHORIZED nao se aplica a peek — limpamos para evitar guards.
    os.environ.pop("ONLY_PROCESSOS_AUTHORIZED", None)

    # Credenciais do mesmo jeito que src/main.py:main() faz. A URL é
    # configurada logo abaixo via env vars (precisa estar no environ para
    # que funcoes internas do main.py achem).
    username = os.getenv("ETCM_USERNAME") or os.getenv("ETCM_LOGIN") or os.getenv("ETCM_USER") or ""
    password = os.getenv("ETCM_PASSWORD") or os.getenv("ETCM_SENHA") or os.getenv("ETCM_PASS") or ""

    if not username or not password:
        # Sem credenciais nao da pra logar headless. Marca todos como erro.
        for p in processos:
            out[p]["error"] = "credenciais ETCM_USERNAME/PASSWORD ausentes"
        log_fn("[peek] credenciais ETCM_USERNAME/PASSWORD ausentes — abortando peek")
        return out

    with _peek_lock:
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=True,
                    channel="chrome",
                    args=["--headless=new"],
                )
                context = browser.new_context(
                    viewport={"width": 1600, "height": 900},
                    accept_downloads=False,
                    ignore_https_errors=True,
                )
                page = context.new_page()
                try:
                    log_fn(f"[peek] fazendo login como {username}...")
                    _main.login_etcm(
                        page, _etcm_url_for_main, username, password,
                        pause_after_login_ms=1000,
                        login_manual_wait_ms=0,
                        headless=True,
                    )
                    # Abre a Mesa de Trabalho (ponto de partida das buscas).
                    from urllib.parse import urljoin as _urljoin
                    try:
                        page.goto(_urljoin(_etcm_url_for_main, "/paginas/mesatrabalho.aspx"),
                                  wait_until="domcontentloaded", timeout=60000)
                    except Exception:
                        pass
                    main_page = page
                    for proc in processos:
                        if on_processo:
                            try: on_processo(proc, "running", None)
                            except Exception: pass
                        try:
                            log_fn(f"[peek] {proc} ...")
                            grid_page = (
                                _main.open_process_action_context_anywhere(context, main_page, proc)
                                or main_page
                            )
                            active_page = _main.filter_and_open_processo(context, grid_page, proc) or grid_page
                            pieces = _main._enumerate_piece_names(active_page)
                            out[proc]["pieces"] = pieces
                            log_fn(f"[peek] {proc} OK ({len(pieces)} pecas)")
                            if on_processo:
                                try: on_processo(proc, "ok", pieces)
                                except Exception: pass
                            # Fecha abas extras (best-effort).
                            try:
                                _main._close_extra_pages(context, [main_page, grid_page])
                            except Exception:
                                pass
                        except Exception as ex:
                            out[proc]["error"] = str(ex)[:300]
                            log_fn(f"[peek] {proc} FALHOU: {ex}")
                            if on_processo:
                                try: on_processo(proc, "failed", str(ex)[:200])
                                except Exception: pass
                finally:
                    try:
                        context.close()
                    except Exception:
                        pass
                    try:
                        browser.close()
                    except Exception:
                        pass
        finally:
            # Restaura env vars que mexemos.
            for k, v in saved_env.items():
                if v is None:
                    os.environ.pop(k, None)
                else:
                    os.environ[k] = v

    return out


def _classify_from_pieces(proc: str, pieces) -> dict:
    """Persiste features + classifica um processo. Helper para peek_id flow."""
    from . import classifier as _clf
    from . import features as _features
    if not pieces:
        return {"tipo": None, "confidence": 0.0, "label_class": "none",
                "error": "sem pecas"}
    try:
        _features.record(
            processo=proc, tipo_forced=None, pieces=pieces,
            piece_chosen_title=None, piece_chosen_number=None,
            extra={"source": "peek_prefetch"},
        )
    except Exception:
        pass
    pred = _clf.predict_by_processo(proc) if _clf.is_trained() else {
        "tipo": None, "confidence": 0.0, "probs": {}, "reason": "sem modelo"
    }
    conf = pred.get("confidence", 0.0)
    if pred.get("tipo") is None:
        label_class = "none"
    elif conf >= 0.85:
        label_class = "high"
    elif conf >= 0.60:
        label_class = "med"
    else:
        label_class = "low"
    pred["label_class"] = label_class
    pred["source"] = "peek"
    return pred


def peek_and_persist_features(processos: list[str], log_fn=print, peek_id: str | None = None) -> dict[str, dict]:
    """Faz o peek E grava as features. Se peek_id for fornecido, reporta
    progresso processo-a-processo via peek_store.

    Devolve {processo: prediction_dict} ao final (compatibilidade com a
    versão antiga sem peek_id).
    """
    from . import peek_store as _ps
    predictions: dict[str, dict] = {}

    def _on_processo(proc, status, payload):
        if peek_id:
            _ps.append_log(peek_id, f"[peek] {proc} {status}")
        if status == "running":
            return
        if status == "ok":
            pred = _classify_from_pieces(proc, payload)  # payload = pieces
        else:
            pred = {"tipo": None, "confidence": 0.0, "label_class": "none",
                    "error": str(payload) if payload else "falhou",
                    "source": "peek"}
        predictions[proc] = pred
        if peek_id:
            _ps.record_result(peek_id, proc, pred)

    peek_processos(processos, log_fn=log_fn, on_processo=_on_processo)
    return predictions
