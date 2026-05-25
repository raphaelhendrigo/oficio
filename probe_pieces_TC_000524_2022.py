"""
Sonda READ-ONLY: lista as peças do TC/000524/2022 e tenta extrair Referência
de cada peça que pareça ser REQUERIMENTO/OFÍCIO/SOLICITAÇÃO/DILAÇÃO.

NÃO faz cleanup, NÃO cria comunicação, NÃO anexa nada. Apenas:
 1) Login (reusa storage_state.json se válido).
 2) Abre o processo.
 3) Enumera peças (índice, nome).
 4) Para peças cujo título sugira ser o documento-fonte da Referência,
    abre o PDF e procura "Ofício nº NNN/AAAA - ..." (mesmo regex de
    _extract_referencia_from_requerimento).

Como rodar:
    .\.venv\Scripts\python probe_pieces_TC_000524_2022.py
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

# Permite importar main.py
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from playwright.sync_api import sync_playwright  # noqa: E402

from main import (  # noqa: E402
    _is_login_page,
    login_etcm,
    open_process_action_context_anywhere,
    filter_and_open_processo,
    search_processo_and_open_viewer,
    _enumerate_piece_names,
    click_last_piece_and_open_pdf,
    extract_text_from_pdf,
    find_frame_with_selector,
)

PROCESSO = "TC/000524/2022"
URL = os.environ.get("ETCM_URL", "https://etcm.tcm.sp.gov.br/paginas/login.aspx")

# Padrões de título que valem a pena inspecionar como possíveis fontes da Referência.
# (REQUERIMENTO é o padrão da Educação. Para Saúde, palpitamos com SOLICITACAO, OFICIO,
# DILACAO no título, ou qualquer peça antes do MANUTAP-OF.)
TITLE_KEYWORDS = ["REQUERIMENTO", "SOLICITAC", "SOLICITAÇ", "OFICIO", "OFÍCIO", "DILAC", "DILAÇ", "SMS", "3403/2026", "897/2026"]

# Mesmo regex de _extract_referencia_from_requerimento.
REF_RE = re.compile(r"Of[ií]cio\s+n[º°\.º°]+\s*\d+\s*/\s*\d{4}[^\n\r]*", re.I)


def main() -> int:
    # Credenciais
    for k in ("ETCM_USERNAME", "ETCM_USER", "ETCM_LOGIN"):
        v = os.environ.get(k)
        if v:
            username = v
            break
    else:
        username = ""
    for k in ("ETCM_PASSWORD", "ETCM_PASS", "ETCM_SENHA"):
        v = os.environ.get(k)
        if v:
            password = v
            break
    else:
        password = ""
    if not username or not password:
        print("ERRO: ETCM_USERNAME / ETCM_PASSWORD nao encontradas no ambiente.")
        return 2

    output_dir = ROOT / "output"
    output_dir.mkdir(exist_ok=True)
    storage_state_file = ROOT / "storage_state.json"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=200)
        ctx_kwargs = {
            "viewport": {"width": 1600, "height": 900},
            "accept_downloads": True,
            "ignore_https_errors": True,
        }
        if storage_state_file.exists():
            ctx_kwargs["storage_state"] = str(storage_state_file)
            print(f"[probe] Reusando sessao: {storage_state_file}")
        context = browser.new_context(**ctx_kwargs)
        page = context.new_page()

        # 1) Login (reusa sessao se possivel)
        mesa_url = urljoin(URL, "/paginas/mesatrabalho.aspx")
        try:
            page.goto(mesa_url, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        if _is_login_page(page):
            print("[probe] Logando...")
            login_etcm(page, username, password, url=URL, manual_wait_ms=60000)
        else:
            print("[probe] Sessao anterior reutilizada.")

        # 2) Navegar para o processo
        print(f"[probe] Abrindo {PROCESSO} ...")
        grid_page = open_process_action_context_anywhere(context, page, PROCESSO) or page
        active_page = filter_and_open_processo(context, grid_page, PROCESSO) or grid_page
        try:
            find_frame_with_selector(active_page, "#splLeitorDocumentos_pgcPecas_trePecas", timeout_ms=15000)
        except Exception:
            try:
                active_page.locator(f"#cod_processo[value*='{PROCESSO}']").first.wait_for(state="attached", timeout=8000)
            except Exception:
                active_page = search_processo_and_open_viewer(context, grid_page, PROCESSO)
        print("[probe] Visualizador aberto.")

        # 3) Enumerar pecas
        pieces = _enumerate_piece_names(active_page)
        print(f"\n[probe] {len(pieces)} pecas encontradas:")
        for idx, name in pieces:
            print(f"  {idx:>3}. {name}")

        # 4) Para pecas com palavras-chave no titulo, baixar e tentar extrair Referencia
        candidates = []
        for idx, name in pieces:
            nu = name.upper()
            if any(k in nu for k in TITLE_KEYWORDS):
                candidates.append((idx, name))

        if not candidates:
            print("\n[probe] Nenhuma peca com palavra-chave de Referencia (REQUERIMENTO/SOLICITACAO/OFICIO/...).")
            print("[probe] Mostre essa lista ao operador para decidir manualmente.")
        else:
            print(f"\n[probe] {len(candidates)} candidato(s) a fonte da Referencia:")
            for idx, name in candidates:
                print(f"\n--- peca {idx}: {name} ---")
                try:
                    # Reusa o helper do main.py para baixar o PDF da peca cujo titulo casa.
                    pdf_path, _t, _n = click_last_piece_and_open_pdf(
                        context, active_page, output_dir, PROCESSO,
                        position=f"match:{name[:40]}",
                        return_piece_number=True,
                    )
                except Exception as e:
                    print(f"  [probe] erro baixando: {e}")
                    continue
                if not pdf_path:
                    print("  [probe] PDF nao baixado.")
                    continue
                txt = extract_text_from_pdf(pdf_path) or ""
                # Mostra as primeiras linhas para inspecao
                preview = "\n".join(l for l in txt.splitlines()[:35] if l.strip())
                print(f"  [probe] preview (primeiras linhas):\n{preview}")
                m = REF_RE.search(txt)
                if m:
                    ref = re.sub(r"\s+", " ", m.group(0)).strip()
                    print(f"  [probe] >>>>>> REFERENCIA CANDIDATA: '{ref}'")
                else:
                    print("  [probe] (regex 'Oficio n. NNN/AAAA' nao casou nessa peca)")

        print("\n[probe] Concluido. Nada foi alterado no e-TCM.")
        try:
            context.storage_state(path=str(storage_state_file))
        except Exception:
            pass
        browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
