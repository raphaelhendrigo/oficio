"""Cadastro de diretores por secretaria — usado no ENCAMINHAMENTO.

Cada processo APO-PEN, após o ofício SSG ser anexado, precisa gerar um
"encaminhamento" cuja tabela tem 3 células:
   1) Nº do ofício SSG gerado (preenchido automaticamente)
   2) Nome do diretor(a) responsável pela secretaria destino
   3) Cargo do diretor(a) — ex.: "Diretora/SME" ou "Diretor/SMS"

Esse módulo persiste a lista {secretaria -> (nome, cargo)} em
`web/labels/directors.json`. Interface pode editar via /api/directors.
"""
from __future__ import annotations

import json
import threading
from dataclasses import dataclass, asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
STORE_PATH = ROOT / "web" / "labels" / "directors.json"
STORE_PATH.parent.mkdir(parents=True, exist_ok=True)

# Chaves canônicas de secretaria — batem com o que o robô já usa
# (educacao/saude/geral em `_detect_secretaria_from_text` e no
# fallback de `_select_template_for`).
KEYS = ("educacao", "saude", "geral")


@dataclass
class Director:
    key: str
    name: str
    cargo: str
    secretaria_label: str  # rótulo humano — só para exibição

    def is_configured(self) -> bool:
        return bool(self.name.strip())


DEFAULTS: dict[str, Director] = {
    "educacao": Director(
        key="educacao",
        name="Vandréia Cristian de Oliveira",
        cargo="Diretora/SME",
        secretaria_label="Secretaria Municipal de Educação (SME)",
    ),
    "saude": Director(
        key="saude",
        name="Cassio Cavalcante Farias",
        cargo="Diretor/SMS",
        secretaria_label="Secretaria Municipal da Saúde (SMS)",
    ),
    "geral": Director(
        key="geral",
        name="",
        cargo="",
        secretaria_label="Outras Secretarias (Geral)",
    ),
}

_lock = threading.Lock()


# ----- Persistência ----------------------------------------------------------

def _load_raw() -> dict:
    if not STORE_PATH.exists():
        return {}
    try:
        data = json.loads(STORE_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_raw(data: dict) -> None:
    STORE_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _merge_defaults(raw: dict) -> dict[str, Director]:
    """Aplica overrides do JSON sobre os defaults hardcoded."""
    out: dict[str, Director] = {}
    for k, d in DEFAULTS.items():
        override = raw.get(k) or {}
        out[k] = Director(
            key=k,
            name=(override.get("name") or d.name).strip(),
            cargo=(override.get("cargo") or d.cargo).strip(),
            secretaria_label=d.secretaria_label,
        )
    return out


# ----- API pública -----------------------------------------------------------

def list_directors() -> list[Director]:
    with _lock:
        raw = _load_raw()
    return list(_merge_defaults(raw).values())


def get(key: str) -> Director | None:
    """Case-insensitive por chave curta ou por rótulo humano (best-effort)."""
    if not key:
        return None
    kk = key.strip().lower()
    with _lock:
        raw = _load_raw()
    merged = _merge_defaults(raw)
    if kk in merged:
        return merged[kk]
    # Fallback: procura por substring no label ou no nome — útil quando o
    # main.py passa a secretaria detectada como "Educação" ou "Saúde".
    from unicodedata import normalize as _un
    def _norm(s: str) -> str:
        return "".join(
            c for c in _un("NFD", s or "") if not (0x0300 <= ord(c) <= 0x036f)
        ).lower()
    kn = _norm(kk)
    for d in merged.values():
        if kn in _norm(d.key) or kn in _norm(d.secretaria_label):
            return d
    return None


def update(key: str, name: str | None = None, cargo: str | None = None) -> Director:
    """Atualiza nome e/ou cargo de uma secretaria conhecida (educacao/saude/geral)."""
    kk = (key or "").strip().lower()
    if kk not in DEFAULTS:
        raise ValueError(f"secretaria desconhecida: {key}")
    with _lock:
        raw = _load_raw()
        entry = dict(raw.get(kk) or {})
        if name is not None:
            entry["name"] = name.strip()
        if cargo is not None:
            entry["cargo"] = cargo.strip()
        raw[kk] = entry
        _save_raw(raw)
        merged = _merge_defaults(raw)
    return merged[kk]


def to_dict(d: Director) -> dict:
    return asdict(d)
