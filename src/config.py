import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


def env_bool(name: str, default: bool = False) -> bool:
    val = os.getenv(name)
    if val is None:
        return default
    return str(val).strip().lower() in {"1", "true", "yes", "y", "on"}


def env_int(name: str, default: int) -> int:
    val = os.getenv(name)
    if val is None or val == "":
        return default
    try:
        return int(val)
    except ValueError:
        return default


def env_str(name: str, default: str = "") -> str:
    val = os.getenv(name)
    if val is None:
        return default
    return str(val).strip()


def parse_list(value: str) -> list[str]:
    if not value:
        return []
    sep = ";" if ";" in value and "," not in value else ","
    return [item.strip() for item in value.split(sep) if item.strip()]


@dataclass
class Config:
    base_url: str
    etcm_user: str
    etcm_pass: str
    headless: bool
    slowmo_ms: int
    timeout_ms: int
    mode: str
    steps_path: Path
    log_dir: Path
    artifacts_dir: Path
    download_dir: Path
    evidence_dir: Path
    html_dir: Path
    use_storage_state: bool
    storage_state_path: Path
    process_list: list[str]
    process_all: bool
    max_processes: int
    single_processo: str
    destinatario: str
    relator: str
    descricao: str
    status_entrega: str
    prazo_dias: str
    anexo_docx: str
    visoes: str
    distribuido_para: str


def load_config() -> Config:
    load_dotenv()

    base_url = env_str("BASE_URL") or env_str("ETCM_URL", "https://homologacao-etcm.tcm.sp.gov.br/paginas/login.aspx")
    etcm_user = env_str("ETCM_USER") or env_str("ETCM_USERNAME") or env_str("ETCM_LOGIN")
    etcm_pass = env_str("ETCM_PASS") or env_str("ETCM_PASSWORD") or env_str("ETCM_SENHA")

    headless = env_bool("HEADLESS", False)
    slowmo_ms = env_int("SLOWMO_MS", 150)
    timeout_ms = env_int("TIMEOUT_MS", 30000)
    mode = env_str("MODE", "run")

    steps_path = Path(env_str("STEPS_PATH", "docs/steps.yaml"))
    artifacts_dir = Path(env_str("ARTIFACTS_DIR", "artifacts"))
    download_dir = Path(env_str("DOWNLOAD_DIR", str(artifacts_dir / "downloads")))
    evidence_dir = Path(env_str("EVIDENCE_DIR", str(artifacts_dir / "evidence")))
    html_dir = Path(env_str("HTML_DIR", str(artifacts_dir / "html")))
    log_dir = Path(env_str("LOG_DIR", "logs"))

    use_storage_state = env_bool("USE_STORAGE_STATE", True)
    storage_state_path = Path(env_str("STORAGE_STATE_PATH", "storage_state.json"))

    process_list = parse_list(env_str("PROCESSOS_LIST") or env_str("PROCESS_LIST"))
    process_all = env_bool("PROCESS_ALL", False)
    max_processes = env_int("MAX_PROCESSOS", 0)
    single_processo = env_str("PROCESSO") or env_str("PROCESSO_LABEL")

    destinatario = env_str("DESTINATARIO", "Secretaria Municipal de Educacao")
    relator = env_str("RELATOR", "")
    descricao = env_str("DESCRICAO", "Oficio gerado automaticamente")
    status_entrega = env_str("STATUS_ENTREGA", "Urgente")
    prazo_dias = env_str("PRAZO_DIAS", "60")
    anexo_docx = env_str("ANEXO_DOCX", "")
    visoes = env_str("VISOES", "Aposentadoria")
    distribuido_para = env_str("DISTRIBUIDO_PARA", "")

    return Config(
        base_url=base_url,
        etcm_user=etcm_user,
        etcm_pass=etcm_pass,
        headless=headless,
        slowmo_ms=slowmo_ms,
        timeout_ms=timeout_ms,
        mode=mode,
        steps_path=steps_path,
        log_dir=log_dir,
        artifacts_dir=artifacts_dir,
        download_dir=download_dir,
        evidence_dir=evidence_dir,
        html_dir=html_dir,
        use_storage_state=use_storage_state,
        storage_state_path=storage_state_path,
        process_list=process_list,
        process_all=process_all,
        max_processes=max_processes,
        single_processo=single_processo,
        destinatario=destinatario,
        relator=relator,
        descricao=descricao,
        status_entrega=status_entrega,
        prazo_dias=prazo_dias,
        anexo_docx=anexo_docx,
        visoes=visoes,
        distribuido_para=distribuido_para,
    )
