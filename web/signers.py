"""Cadastro de assinantes para o disparador Euclides.

Dois presets fixos no código (não dá pra apagar):
  - Roseli Chaves (titular, modelos historicos em modelos_utap/...)
  - Daniela Shimizu (substituta, modelos em modelos daniela/...)

Mais: lista de "outros" criados pelo Gilson na interface, persistidos em
`web/labels/signers.json`. Cada custom pode ser apagado a qualquer hora.

O JSON também guarda o `active_id` — o assinante ativo é o que o
`web/runner.py` propaga para o `src/main.py` em cada disparo.
"""
from __future__ import annotations

import json
import re
import threading
import uuid
from dataclasses import dataclass, field, asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
STORE_PATH = ROOT / "web" / "labels" / "signers.json"
STORE_PATH.parent.mkdir(parents=True, exist_ok=True)


@dataclass
class Signer:
    id: str
    name: str
    tokens: list[str]
    templates_dirs: dict[str, str] = field(default_factory=dict)
    secretaria_fallback: str | None = None
    is_preset: bool = False
    description: str = ""

    def to_env(self) -> dict[str, str]:
        """Vars de ambiente que o subprocess do main.py vai consumir."""
        env: dict[str, str] = {
            "ASSINANTE_NOME": self.name,
            "SIGNER_NAME": self.name,
            "SIGNER_MATCH_TOKENS": ",".join(self.tokens),
        }
        for tipo, folder in self.templates_dirs.items():
            env[f"OFICIO_TEMPLATES_DIR_{tipo.upper()}"] = folder
        if self.secretaria_fallback:
            env["OFICIO_SECRETARIA_FALLBACK"] = self.secretaria_fallback
        return env


# ----- Presets imutáveis -----------------------------------------------------

PRESETS: dict[str, Signer] = {
    "roseli": Signer(
        id="roseli",
        name="Roseli Chaves",
        tokens=["roseli", "chaves"],
        templates_dirs={
            "UTAP": "modelos_utap",
            "DILACAO": "modelos_dilacao",
            "REITERACAO": "modelos_reiteracao",
        },
        secretaria_fallback=None,
        is_preset=True,
        description="Titular (Subsecretária-Geral)",
    ),
    "daniela": Signer(
        id="daniela",
        name="Daniela Shimizu",
        tokens=["daniela", "shimizu"],
        templates_dirs={
            "UTAP": "modelos daniela/manutap",
            "DILACAO": "modelos daniela/dilação",
            "REITERACAO": "modelos daniela/reiteração",
        },
        secretaria_fallback="educacao",
        is_preset=True,
        description="Substituta da Roseli durante férias",
    ),
}

DEFAULT_ACTIVE = "daniela"

_lock = threading.Lock()


# ----- Persistência ----------------------------------------------------------

def _load_raw() -> dict:
    if not STORE_PATH.exists():
        return {"active_id": DEFAULT_ACTIVE, "customs": []}
    try:
        data = json.loads(STORE_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("formato invalido")
        data.setdefault("active_id", DEFAULT_ACTIVE)
        data.setdefault("customs", [])
        return data
    except Exception:
        return {"active_id": DEFAULT_ACTIVE, "customs": []}


def _save_raw(data: dict) -> None:
    STORE_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _custom_to_signer(c: dict) -> Signer:
    """Hidrata um custom do JSON: campos faltantes herdam de Roseli (templates)."""
    base = PRESETS["roseli"]
    return Signer(
        id=c["id"],
        name=c["name"],
        tokens=list(c.get("tokens") or []),
        templates_dirs=dict(c.get("templates_dirs") or base.templates_dirs),
        secretaria_fallback=c.get("secretaria_fallback"),
        is_preset=False,
        description=c.get("description") or "Personalizado",
    )


# ----- API pública -----------------------------------------------------------

def list_signers() -> list[Signer]:
    """Devolve presets (na ordem fixa) + customs (na ordem de criação)."""
    with _lock:
        data = _load_raw()
    out: list[Signer] = [PRESETS["roseli"], PRESETS["daniela"]]
    for c in data.get("customs", []):
        try:
            out.append(_custom_to_signer(c))
        except Exception:
            continue
    return out


def active() -> Signer:
    """Assinante atualmente selecionado. Cai no DEFAULT_ACTIVE se inconsistente."""
    with _lock:
        data = _load_raw()
        aid = data.get("active_id") or DEFAULT_ACTIVE
    for s in list_signers():
        if s.id == aid:
            return s
    return PRESETS[DEFAULT_ACTIVE]


def set_active(signer_id: str) -> bool:
    """Marca um assinante como ativo. False se id não existe."""
    signer_id = (signer_id or "").strip()
    if not signer_id:
        return False
    with _lock:
        data = _load_raw()
        all_ids = set(PRESETS.keys()) | {c["id"] for c in data.get("customs", [])}
        if signer_id not in all_ids:
            return False
        data["active_id"] = signer_id
        _save_raw(data)
    return True


def _auto_tokens(name: str) -> list[str]:
    """Gera tokens a partir do nome, ignorando partículas comuns."""
    parts = [p for p in re.split(r"\s+", (name or "").strip()) if p]
    stop = {"de", "da", "do", "das", "dos", "moraes", "morais", "e"}
    keep = [p for p in parts if len(p) > 2 and p.lower() not in stop]
    if len(keep) >= 2:
        return [keep[0].lower(), keep[-1].lower()]
    if keep:
        return [keep[0].lower()]
    return [name.strip().lower()] if name and name.strip() else []


def add_custom(name: str, tokens: list[str] | None = None,
               templates_dirs: dict[str, str] | None = None,
               secretaria_fallback: str | None = None) -> Signer:
    """Adiciona um assinante personalizado. tokens auto se omitido.

    Templates default = pastas da Roseli (modelos_utap etc.) — o operador
    pode mudar depois editando o JSON se precisar.
    """
    name = (name or "").strip()
    if not name:
        raise ValueError("nome vazio")
    if not tokens:
        tokens = _auto_tokens(name)
    if not tokens:
        raise ValueError("não consegui derivar tokens do nome — informe tokens manualmente")
    new_id = "c_" + uuid.uuid4().hex[:10]
    new_custom = {
        "id": new_id,
        "name": name,
        "tokens": [t.lower().strip() for t in tokens if t.strip()],
        "templates_dirs": templates_dirs or dict(PRESETS["roseli"].templates_dirs),
        "secretaria_fallback": secretaria_fallback,
        "description": "Personalizado",
    }
    with _lock:
        data = _load_raw()
        data.setdefault("customs", []).append(new_custom)
        _save_raw(data)
    return _custom_to_signer(new_custom)


def delete_custom(signer_id: str) -> bool:
    """Remove um assinante personalizado. Presets não podem ser removidos.

    Se o removido era o ativo, volta a Daniela (DEFAULT_ACTIVE).
    """
    if signer_id in PRESETS:
        return False
    with _lock:
        data = _load_raw()
        before = len(data.get("customs", []))
        data["customs"] = [c for c in data.get("customs", []) if c.get("id") != signer_id]
        if len(data["customs"]) == before:
            return False
        if data.get("active_id") == signer_id:
            data["active_id"] = DEFAULT_ACTIVE
        _save_raw(data)
    return True


def get(signer_id: str) -> Signer | None:
    for s in list_signers():
        if s.id == signer_id:
            return s
    return None


def to_dict(s: Signer) -> dict:
    """Serializa para envio na API."""
    d = asdict(s)
    return d
