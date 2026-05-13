"""
Regras de normalizacao usadas no fluxo de Oficios SSG / APO-PEN.

Foco: funcoes puras, sem dependencia de I/O ou Playwright, faceis de testar.

Cobre:
- normalize_descricao_comunicacao: producao da string padrao da comunicacao
  processual (conhecimento/providencias, dilacao, reiteracao, decisao de juizo
  singular) — sempre minusculas, com acento, com barra quando aplicavel.
- normalize_secretaria_label: aceita variacoes (educacao/saude/geral) e
  retorna a forma usada nos diretorios e logs.
- signer_name_matches_roseli_chaves: aceita variacoes de "Roseli Chaves",
  "Roseli de Morais Chaves", "Roseli Moraes Chaves", baseado em tokens
  normalizados (sem acentos, em minusculas).
"""

from __future__ import annotations

import re
import unicodedata
from typing import Iterable


# ----------------------------- Helpers internos -----------------------------

def _strip_accents(s: str) -> str:
    """Remove acentos preservando o casing original."""
    if not s:
        return ""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _norm_compare(s: str) -> str:
    """Forma comparavel: sem acentos, minusculas, sem pontuacao redundante."""
    if not s:
        return ""
    out = _strip_accents(s).lower()
    out = re.sub(r"[^a-z0-9]+", " ", out)
    return out.strip()


def _tokens(s: str) -> set[str]:
    """Conjunto de tokens alfanumericos normalizados (sem acentos)."""
    return set(_norm_compare(s).split())


# --------------------------- Descricao da comunicacao -----------------------

# Saidas oficiais aceitas pelo e-TCM e exigidas pelo brief.
DESCRICAO_CONHECIMENTO_PROVIDENCIAS = "conhecimento/providências"
DESCRICAO_DILACAO = "dilação"
DESCRICAO_REITERACAO = "reiteração"
DESCRICAO_JUIZO_SINGULAR = "decisão de juízo singular"

# Tipos internos reconhecidos pelo classificador (existente em main.py).
TIPO_CANONICO_UTAP = "UTAP"
TIPO_CANONICO_DILACAO = "DILACAO"
TIPO_CANONICO_REITERACAO = "REITERACAO"
TIPO_CANONICO_JUIZO = "JUIZO"


def normalize_descricao_comunicacao(tipo_ou_label: str) -> str:
    """Converte qualquer rotulo de tipo (canonico ou variacao livre) na
    descricao oficial da comunicacao processual.

    Exemplos aceitos -> retorno:
      "UTAP", "conhecimento e providencias", "CONHECIMENTO E PROVIDENCIAS",
      "providencias", "conhecimento/providencias" -> "conhecimento/providências"
      "DILACAO", "dilacao", "DILAÇÃO" -> "dilação"
      "REITERACAO", "reiteracao", "REITERAÇÃO" -> "reiteração"
      "JUIZO", "juizo singular", "decisao de juizo singular" -> "decisão de juízo singular"

    Levanta ValueError quando nao reconhece. Mantemos rigido por seguranca:
    a descricao errada e um dos bugs reportados no brief.
    """
    if not tipo_ou_label or not str(tipo_ou_label).strip():
        raise ValueError("Tipo da comunicacao nao informado")

    key = _norm_compare(tipo_ou_label)

    # Juizo singular tem prioridade quando aparece junto com 'decisao'.
    if "juizo singular" in key or "decisao de juizo" in key or key == "juizo":
        return DESCRICAO_JUIZO_SINGULAR

    if key == "dilacao" or "dilacao" in key.split():
        return DESCRICAO_DILACAO

    if key == "reiteracao" or "reiteracao" in key.split():
        return DESCRICAO_REITERACAO

    # UTAP cobre: utap, providencias, conhecimento e providencias,
    # conhecimento providencias, conhecimento/providencias
    if (
        key == "utap"
        or "providencias" in key
        or "conhecimento" in key
    ):
        return DESCRICAO_CONHECIMENTO_PROVIDENCIAS

    raise ValueError(f"Tipo de comunicacao nao reconhecido: {tipo_ou_label!r}")


# ----------------------------- Secretaria -----------------------------------

SECRETARIA_EDUCACAO = "Educação"
SECRETARIA_SAUDE = "Saúde"
SECRETARIA_GERAL = "Geral"


def normalize_secretaria_label(text: str) -> str:
    """Detecta a secretaria a partir de um texto livre.

    Heuristica conservadora: primeiro educacao, depois saude, depois geral.
    A regra de prioridade segue a usada em main._detect_secretaria_from_text
    (educacao > saude > geral). Mantida aqui para permitir testes isolados e
    eventual reuso fora de main.py.
    """
    if not text:
        return SECRETARIA_GERAL
    t = _norm_compare(text)
    educ_keys = (
        "secretaria municipal de educacao",
        "secretaria de educacao",
        "sme",
        "educacao",
    )
    if any(k in t for k in educ_keys):
        return SECRETARIA_EDUCACAO

    saude_keys = (
        "secretaria municipal da saude",
        "secretaria municipal de saude",
        "secretaria de saude",
        "sms",
        "saude",
    )
    if any(k in t for k in saude_keys):
        return SECRETARIA_SAUDE

    return SECRETARIA_GERAL


# ----------------------------- Assinante ------------------------------------

# Token raiz que identifica o assinante alvo, independente de "DE MORAIS",
# "MORAIS", "MORAES" ou de o nome estar abreviado.
SIGNER_REQUIRED_TOKENS: tuple[str, ...] = ("roseli", "chaves")


def signer_name_matches_roseli_chaves(candidate: str,
                                       required: Iterable[str] = SIGNER_REQUIRED_TOKENS) -> bool:
    """True se o nome candidato contiver TODOS os tokens raiz (normalizados).

    Aceita ROSELI DE MORAIS CHAVES, ROSELI MORAES CHAVES, Roseli Chaves,
    Roseli M. Chaves, "Sra. Roseli de Moraes Chaves", etc. Nao aceita apenas
    "Roseli" nem apenas "Chaves" — exige a combinacao.

    O parametro `required` permite reaproveitar a funcao para outros
    assinantes (ex.: tokens 'fulano','tal').
    """
    if not candidate:
        return False
    candidate_tokens = _tokens(candidate)
    required_tokens = {_norm_compare(t) for t in required}
    required_tokens.discard("")
    if not required_tokens:
        return False
    return required_tokens.issubset(candidate_tokens)


def find_signer_in_list(names: Iterable[str],
                        required: Iterable[str] = SIGNER_REQUIRED_TOKENS) -> str | None:
    """Retorna o PRIMEIRO nome em `names` que combine com `required`.

    Quando varios candidatos combinam, e retornado o primeiro encontrado.
    Util para iteracoes sobre listas exibidas em telas (autocomplete,
    grids de selecao de assinante).
    """
    for n in names:
        if signer_name_matches_roseli_chaves(n, required=required):
            return n
    return None
