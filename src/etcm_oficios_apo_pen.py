import os
import re
import time
import unicodedata
from pathlib import Path
from typing import List, Optional, Union
from urllib.parse import urljoin

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import Page, sync_playwright

# Configuracoes padrao (ajuste facilmente aqui)
ETCM_URL = "https://homologacao-etcm.tcm.sp.gov.br/paginas/login.aspx"
DESTINATARIO_PADRAO = "Secretaria Municipal de Educacao (*)"
RELATOR_PADRAO = "DOMINGOS DISSEI"
DESCRICAO_PADRAO = "dilacao"
REFERENCIA_PADRAO = "dilacao"
STATUS_PADRAO = "Urgente"
PRAZO_PADRAO = "60"


def log(msg: str) -> None:
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def env_bool(name: str, default: bool = False) -> bool:
    return str(os.getenv(name, str(default))).strip().lower() in {"1", "true", "yes", "y", "on"}


def _strip_quotes(v: Optional[str]) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if len(s) >= 2 and ((s[0] == s[-1] == '"') or (s[0] == s[-1] == "'")):
        s = s[1:-1]
    return s.strip()


def _is_login_page(page: Page) -> bool:
    """Detecta se ainda estamos na tela de login do e‑TCM."""
    try:
        user_loc = page.locator("#ctl00_cphMain_txtUsuario_I").first
        pass_loc = page.locator("#ctl00_cphMain_txtSenha_I").first
        if user_loc.count() > 0 and pass_loc.count() > 0:
            if user_loc.is_visible() or pass_loc.is_visible():
                return True
    except Exception:
        pass
    try:
        url = page.url or ""
        if "login.aspx" in url.lower():
            return True
    except Exception:
        pass
    return False


def _normalize(text: str) -> str:
    base = unicodedata.normalize("NFKD", text or "")
    base = "".join(ch for ch in base if not unicodedata.combining(ch))
    base = re.sub(r"\s+", " ", base)
    return base.strip().lower()


def wait_devexpress_idle(page: Page, timeout: int = 15000) -> None:
    try:
        page.wait_for_timeout(200)
        page.wait_for_function(
            "() => !document.querySelector('.dxgvLoadingPanel, .dxlpLoadingPanel, .dx-loading-panel, div.dx-loading')",
            timeout=timeout,
        )
    except Exception:
        pass


def wait_visible(page: Page, selector: str, timeout: int = 15000):
    return page.locator(selector).first.wait_for(state="visible", timeout=timeout)


def _first_container_with_selector(page: Page, selector: str, timeout_ms: int = 0) -> Optional[Union[Page, "Frame"]]:
    """Retorna a primeira pagina/frame que contenha o seletor."""
    deadline = time.time() + (timeout_ms / 1000.0) if timeout_ms > 0 else None
    while True:
        containers = [page] + list(page.frames)
        for c in containers:
            try:
                loc = c.locator(selector).first
                if loc.count() > 0:
                    try:
                        loc.wait_for(state="visible", timeout=2000)
                    except Exception:
                        pass
                    return c
            except Exception:
                continue
        if deadline is None or time.time() >= deadline:
            break
        time.sleep(0.3)
    return None


def login_etcm(
    page: Page,
    url: str,
    username: str,
    password: str,
    pause_after_login_ms: int = 0,
    login_manual_wait_ms: int = 45000,
    headless: bool = False,
) -> None:
    """Realiza login no e‑TCM de forma robusta (DevExpress + captchas)."""
    username = _strip_quotes(username)
    password = _strip_quotes(password)

    log("Acessando tela de login...")
    page.goto(url, wait_until="domcontentloaded", timeout=60000)

    try:
        page.locator("#ctl00_cphMain_txtUsuario_I, input[name='ctl00$cphMain$txtUsuario']").first.wait_for(
            state="visible", timeout=30000
        )
    except Exception:
        pass

    # Usuario / senha
    page.fill("#ctl00_cphMain_txtUsuario_I", username)
    page.fill("#ctl00_cphMain_txtSenha_I", password)

    btn_selectors = [
        "#ctl00_cphMain_btnLogin_I",
        "#ctl00_cphMain_btnLogin",
        "input[name='ctl00$cphMain$btnLogin']",
        "input[type='submit'][value*='Entrar' i]",
        "button:has-text('Entrar')",
    ]
    clicked = False
    for sel in btn_selectors:
        try:
            page.locator(sel).first.click(timeout=5000)
        except Exception:
            continue
        try:
            page.wait_for_function(
                "() => document.querySelector('#ctl00_cphMain_txtUsuario_I') === null && "
                "document.querySelector('#ctl00_cphMain_txtSenha_I') === null",
                timeout=15000,
            )
            clicked = True
            break
        except Exception:
            continue

    if not clicked:
        try:
            page.evaluate(
                "(() => { try { __doPostBack('ctl00$cphMain$btnLogin',''); } catch(e) { "
                "var f=document.forms['aspnetForm']; if(f){f.__EVENTTARGET.value='ctl00$cphMain$btnLogin'; "
                "f.__EVENTARGUMENT.value=''; f.submit();} } })();"
            )
        except Exception:
            pass

    wait_devexpress_idle(page)

    if pause_after_login_ms > 0:
        log(f"Pausa apos tentativa de login: {pause_after_login_ms} ms")
        time.sleep(pause_after_login_ms / 1000.0)

    if _is_login_page(page):
        if login_manual_wait_ms != 0:
            log(
                "Tela de login ainda visivel. Resolva captcha/erro manualmente e clique em Entrar. "
                f"Aguardando ate {login_manual_wait_ms} ms..."
            )
            deadline = time.time() + (login_manual_wait_ms / 1000.0)
            while time.time() < deadline:
                if not _is_login_page(page):
                    break
                time.sleep(1.5)
        else:
            log("Aguardando login manual sem timeout (Ctrl+C para abortar)...")
            while _is_login_page(page):
                time.sleep(1.5)
        if _is_login_page(page):
            extra = ""
            try:
                if page.locator("#gRecaptchaToken, iframe[src*='recaptcha' i], div.g-recaptcha").count() > 0:
                    extra = " Captcha/reCAPTCHA detectado."
            except Exception:
                pass
            if headless and "captcha" in extra.lower():
                extra += " (modo headless nao permite resolver; use HEADLESS=false)."
            raise RuntimeError("Login nao concluido: formulario ainda visivel." + extra)

    # Garante Mesa de Trabalho
    try:
        if "mesatrabalho.aspx" not in (page.url or "").lower():
            mesa_url = urljoin(url, "/paginas/mesatrabalho.aspx")
            page.goto(mesa_url, wait_until="domcontentloaded", timeout=60000)
    except Exception:
        pass
    log("Login concluido.")


def _abrir_fila_apopen(page: Page, timeout: int = 20000) -> Union[Page, "Frame"]:
    """Abre Processos > Oficios > Em confeccao APO-PEN e retorna o container que tem a grid."""
    log("Abrindo menu Processos > Oficios > Em confeccao APO-PEN...")
    selectors_grid = "#gvProcesso, table#gvProcesso, #sptMesaTrabalho_gvProcesso"
    labels = [
        re.compile(r"Processos", re.I),
        re.compile(r"Of[ií]cios", re.I),
        re.compile(r"Em\s*confec", re.I),
        re.compile(r"APO-?PEN", re.I),
    ]

    deadline = time.time() + (timeout / 1000.0)
    # Tenta acionar via JavaScript (mais robusto que clicar na árvore renderizada)
    try:
        page.evaluate(
            """
            () => {
              try {
                if (typeof AtualizarGrid === 'function') {
                  AtualizarGrid('confappen_16','PROCESSO');
                  return true;
                }
              } catch (e) {}
              try {
                if (typeof AtualizarGrid === 'function') {
                  AtualizarGrid('conf_16','PROCESSO');
                  return true;
                }
              } catch (e) {}
              return false;
            }
            """
        )
    except Exception:
        pass

    while time.time() < deadline:
        containers = [page] + list(page.frames)
        for container in containers:
            for label in labels:
                try:
                    container.get_by_text(label, exact=False).first.click()
                    container.wait_for_timeout(200)
                except Exception:
                    continue
            # tenta alguns seletores diretos por id
            for sel in [
                "a#016_PROCESSO, a[id*='UNIDADE']",
                "a#confappen_16_PROCESSO, a[id*='confappen']",
                "a:has-text('Em confeccao APO-PEN')",
            ]:
                try:
                    container.locator(sel).first.click()
                except Exception:
                    pass
        wait_devexpress_idle(page)
        found = _first_container_with_selector(page, selectors_grid, timeout_ms=2000)
        if found:
            try:
                page.wait_for_function(
                    "sel => { const g = document.querySelector(sel); return g && getComputedStyle(g).display !== 'none'; }",
                    arg=selectors_grid,
                    timeout=5000,
                )
            except Exception:
                pass
            log("Grid APO-PEN visivel.")
            return found
        time.sleep(0.5)

    found = _first_container_with_selector(page, selectors_grid, timeout_ms=2000)
    if found:
        log("Grid APO-PEN visivel.")
        return found
    raise RuntimeError("Grid da fila Em confeccao APO-PEN nao ficou visivel.")


def exportar_planilha_apopen(page: Page) -> Path:
    container = _abrir_fila_apopen(page)
    tmp_dir = Path("tmp")
    tmp_dir.mkdir(parents=True, exist_ok=True)

    export_selectors = [
        "#gvProcesso_Title_btnExport_I",
        "#gvProcesso_Title_btnExport",
        "input[id*='btnExport' i]",
        "button:has-text('Exportar')",
        "a:has-text('Exportar')",
    ]
    download_path: Optional[Path] = None
    log("Disparando exportacao da planilha APO-PEN...")
    for sel in export_selectors:
        loc = container.locator(sel).first
        if loc.count() == 0:
            continue
        try:
            with page.expect_download(timeout=60000) as dl_info:
                loc.click()
            download = dl_info.value
            suggested = download.suggested_filename or "apopen.xlsx"
            download_path = tmp_dir / suggested
            download.save_as(str(download_path))
            break
        except Exception:
            continue
    if not download_path:
        raise RuntimeError("Nao foi possivel disparar o download da planilha (botao Exportar).")
    log(f"Planilha exportada: {download_path.resolve()}")
    return download_path


def _coluna_processo(df: pd.DataFrame) -> Optional[str]:
    for col in df.columns:
        norm = _normalize(str(col))
        if "processo" in norm:
            return col
        if "n" in norm and "proc" in norm:
            return col
    return df.columns[0] if len(df.columns) else None


def carregar_processos(caminho_arquivo_excel: Path) -> List[str]:
    if not caminho_arquivo_excel.exists():
        raise FileNotFoundError(f"Planilha nao encontrada: {caminho_arquivo_excel}")
    df = pd.read_excel(caminho_arquivo_excel)
    if df.empty:
        return []
    coluna = _coluna_processo(df)
    if not coluna:
        return []
    series = df[coluna].dropna().astype(str).str.strip()
    series = series[series != ""]
    processos = list(dict.fromkeys(series.tolist()))
    log(f"Processos carregados da planilha: {len(processos)} encontrado(s).")
    return processos


def _filtrar_processo_na_grid(page: Page, numero_processo: str) -> Union[Page, "Frame", None]:
    container = _first_container_with_selector(
        page,
        "#gvProcesso_DXMainTable, #sptMesaTrabalho_gvProcesso_DXMainTable, #gvProcesso, #sptMesaTrabalho_gvProcesso",
        timeout_ms=3000,
    ) or page
    try:
        filtro = container.locator("input[id$='_DXFREditorcol17_I'], input[name$='$DXFREditorcol17']").first
        if filtro.count() > 0:
            filtro.fill("")
            filtro.type(numero_processo, delay=50)
            filtro.press("Enter")
            wait_devexpress_idle(page)
            return
    except Exception:
        pass
    try:
        container.get_by_placeholder(re.compile("Processo", re.I)).fill(numero_processo)
    except Exception:
        pass
    wait_devexpress_idle(page)
    return container


def abrir_comunicacoes_processo(page: Page, numero_processo: str) -> Page:
    container = _abrir_fila_apopen(page)
    container = _filtrar_processo_na_grid(page, numero_processo) or container

    row_selector = "#gvProcesso_DXMainTable tr[id*='DXDataRow'], #sptMesaTrabalho_gvProcesso_DXMainTable tr[id*='DXDataRow']"
    row = container.locator(row_selector).filter(has_text=re.compile(re.escape(numero_processo), re.I)).first
    row.wait_for(state="visible", timeout=15000)

    target_page = page
    icon = row.locator("img[src*='img_notificacao' i], a:has(img[src*='img_notificacao' i])").first
    try:
        with page.expect_popup(timeout=8000) as pop_info:
            icon.click()
        target_page = pop_info.value
        target_page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        try:
            icon.click()
        except Exception:
            pass
        target_page = page

    try:
        wait_visible(target_page, "#btnAdicionarNotificacao_I", timeout=12000)
    except Exception:
        pass
    return target_page


def _selecionar_combo(page: Page, base_id: str, valor: str) -> None:
    if not valor:
        return
    input_sel = f"#{base_id}_I, #{base_id} input[id$='_I']"
    try:
        inp = page.locator(input_sel).first
        if inp.count() > 0:
            inp.click()
            inp.fill(valor)
            try:
                inp.press("Enter")
            except Exception:
                pass
            return
    except Exception:
        pass
    try:
        btn = page.locator(f"#{base_id}_B-1, #{base_id}_B0").first
        if btn.count() > 0:
            btn.click()
            page.get_by_text(re.compile(re.escape(valor), re.I)).first.click()
    except Exception:
        pass


def criar_comunicacao(page: Page, dados_processo: dict) -> None:
    descricao = dados_processo.get("descricao", DESCRICAO_PADRAO)
    referencia = dados_processo.get("referencia", REFERENCIA_PADRAO)
    destinatario = dados_processo.get("destinatario", DESTINATARIO_PADRAO)
    relator = dados_processo.get("relator", RELATOR_PADRAO)
    status = dados_processo.get("status", STATUS_PADRAO)
    prazo = str(dados_processo.get("prazo", PRAZO_PADRAO))

    wait_visible(page, "#btnAdicionarNotificacao_I", timeout=15000)
    page.click("#btnAdicionarNotificacao_I")
    wait_visible(page, "#ppcNoificacao_txtDescricao_I", timeout=15000)

    _selecionar_combo(page, "ppcNoificacao_cbbUsuarios", destinatario)
    _selecionar_combo(page, "ppcNoificacao_cbbPessoa", relator)

    page.fill("#ppcNoificacao_txtDescricao_I", descricao)
    page.fill("#ppcNoificacao_txtReferencia_I", referencia)

    try:
        page.locator("#ppcNoificacao_cbbStatusProvidencia").select_option(label=re.compile(status, re.I))
    except Exception:
        try:
            _selecionar_combo(page, "ppcNoificacao_cbbStatusProvidencia", status)
        except Exception:
            pass

    try:
        page.fill("#ppcNoificacao_txtPrazo_I", prazo)
    except Exception:
        pass

    wait_devexpress_idle(page)
    page.click("#ppcNoificacao_btnPopSalvar_I")
    try:
        page.wait_for_selector("#ppcNoificacao_btnPopSalvar_I", state="hidden", timeout=15000)
    except Exception:
        pass
    wait_devexpress_idle(page)
    wait_visible(page, "#gvNotificacao", timeout=20000)
    log(f"Comunicacao cadastrada para o processo {dados_processo.get('processo')}.")


def anexar_atos(page: Page, dados_processo: dict) -> None:
    try:
        grid = page.locator("#gvNotificacao_DXMainTable, #gvNotificacao").first
        row = grid.locator("tr[id*='DXDataRow']").filter(
            has_text=re.compile(re.escape(dados_processo.get("descricao", DESCRICAO_PADRAO)), re.I)
        ).first
        clip = row.locator("img[src*='img_clip' i], img[src*='clip' i], a:has(img[src*='clip' i])").first
        clip.click()
        wait_devexpress_idle(page)
        page.keyboard.press("Escape")
    except Exception:
        pass


def main() -> None:
    load_dotenv()
    url = os.getenv("ETCM_URL", ETCM_URL)
    username = os.getenv("ETCM_LOGIN") or os.getenv("ETCM_USERNAME")
    password = os.getenv("ETCM_SENHA") or os.getenv("ETCM_PASSWORD")
    headless = env_bool("HEADLESS", False)
    slow_mo = int(os.getenv("SLOWMO_MS", "0") or 0)
    pause_after_login_ms = int(os.getenv("PAUSE_AFTER_LOGIN_MS", "0") or 0)
    login_manual_wait_ms = int(os.getenv("LOGIN_MANUAL_WAIT_MS", "45000") or 45000)
    use_storage_state = env_bool("USE_STORAGE_STATE", True)
    storage_state_path = os.getenv("STORAGE_STATE_PATH", "storage_state.json").strip()
    storage_state_file = Path(storage_state_path) if storage_state_path else None

    if not username or not password:
        raise RuntimeError("Defina ETCM_LOGIN/ETCM_SENHA (ou ETCM_USERNAME/ETCM_PASSWORD).")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=slow_mo)
        context_kwargs = {"accept_downloads": True, "viewport": {"width": 1400, "height": 900}}
        if use_storage_state and storage_state_file and storage_state_file.exists():
            context_kwargs["storage_state"] = str(storage_state_file)
            log(f"Carregando sessao anterior: {storage_state_file}")
        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        try:
            need_login = True
            if use_storage_state and storage_state_file and storage_state_file.exists():
                try:
                    mesa_url = urljoin(url, "/paginas/mesatrabalho.aspx")
                    page.goto(mesa_url, wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass
                need_login = _is_login_page(page)
                if not need_login:
                    log("Sessao anterior reutilizada com sucesso.")
            if need_login:
                login_etcm(
                    page,
                    url,
                    username,
                    password,
                    pause_after_login_ms=pause_after_login_ms,
                    login_manual_wait_ms=login_manual_wait_ms,
                    headless=headless,
                )
                if use_storage_state and storage_state_file:
                    try:
                        context.storage_state(path=str(storage_state_file))
                        log(f"Sessao salva em: {storage_state_file}")
                    except Exception:
                        pass
            _abrir_fila_apopen(page)

            planilha = exportar_planilha_apopen(page)
            processos = carregar_processos(planilha)

            for numero in processos:
                log(f"Processando {numero} ...")
                target_page = abrir_comunicacoes_processo(page, numero)
                dados = {
                    "processo": numero,
                    "destinatario": DESTINATARIO_PADRAO,
                    "relator": RELATOR_PADRAO,
                    "descricao": DESCRICAO_PADRAO,
                    "referencia": REFERENCIA_PADRAO,
                    "status": STATUS_PADRAO,
                    "prazo": PRAZO_PADRAO,
                }
                criar_comunicacao(target_page, dados)
                anexar_atos(target_page, dados)  # stub opcional
                if target_page != page:
                    try:
                        target_page.close()
                    except Exception:
                        pass
                    page.bring_to_front()
                    _abrir_fila_apopen(page)
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
