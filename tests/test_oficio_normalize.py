"""Testes para src/oficio_normalize.py."""
from __future__ import annotations

import pytest

from oficio_normalize import (  # type: ignore  # noqa: E402
    DESCRICAO_CONHECIMENTO_PROVIDENCIAS,
    DESCRICAO_DILACAO,
    DESCRICAO_JUIZO_SINGULAR,
    DESCRICAO_REITERACAO,
    SECRETARIA_EDUCACAO,
    SECRETARIA_GERAL,
    SECRETARIA_SAUDE,
    SIGNER_REQUIRED_TOKENS,
    decode_zip_unicode_escape_name,
    find_signer_in_list,
    normalize_descricao_comunicacao,
    normalize_secretaria_label,
    signer_name_matches_roseli_chaves,
)


# ---------------------- normalize_descricao_comunicacao ---------------------

@pytest.mark.parametrize(
    "entrada,saida",
    [
        # UTAP (conhecimento/providencias) — variacoes aceitas
        ("UTAP", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("utap", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("CONHECIMENTO E PROVIDENCIAS", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("conhecimento e providencias", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("conhecimento/providencias", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("conhecimento/providências", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("Conhecimento e Providências", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),
        ("providencias", DESCRICAO_CONHECIMENTO_PROVIDENCIAS),

        # Dilacao
        ("DILACAO", DESCRICAO_DILACAO),
        ("dilacao", DESCRICAO_DILACAO),
        ("dilação", DESCRICAO_DILACAO),
        ("DILAÇÃO", DESCRICAO_DILACAO),

        # Reiteracao
        ("REITERACAO", DESCRICAO_REITERACAO),
        ("reiteracao", DESCRICAO_REITERACAO),
        ("reiteração", DESCRICAO_REITERACAO),
        ("REITERAÇÃO", DESCRICAO_REITERACAO),

        # Juizo singular
        ("JUIZO", DESCRICAO_JUIZO_SINGULAR),
        ("juizo", DESCRICAO_JUIZO_SINGULAR),
        ("juizo singular", DESCRICAO_JUIZO_SINGULAR),
        ("decisao de juizo singular", DESCRICAO_JUIZO_SINGULAR),
        ("Decisão de Juízo Singular", DESCRICAO_JUIZO_SINGULAR),
    ],
)
def test_normalize_descricao_known_variations(entrada: str, saida: str) -> None:
    assert normalize_descricao_comunicacao(entrada) == saida


def test_normalize_descricao_rejects_unknown() -> None:
    with pytest.raises(ValueError):
        normalize_descricao_comunicacao("apenas um comentario solto")


def test_normalize_descricao_rejects_empty() -> None:
    with pytest.raises(ValueError):
        normalize_descricao_comunicacao("")
    with pytest.raises(ValueError):
        normalize_descricao_comunicacao("   ")


@pytest.mark.parametrize(
    "entrada,saida",
    [
        ("Sa#U00fade", "Saúde"),
        ("Educa#U00e7#U00e3o", "Educação"),
        ("Provid#U00eancias", "Providências"),
        ("Dila#U00e7#U00e3o", "Dilação"),
        ("Reitera#U00e7#U00e3o", "Reiteração"),
    ],
)
def test_decode_zip_unicode_escape_name(entrada: str, saida: str) -> None:
    assert decode_zip_unicode_escape_name(entrada) == saida


def test_normalize_descricao_outputs_have_proper_diacritics() -> None:
    # As 4 saidas oficiais devem conter acentos (encoding correto).
    assert "ê" in DESCRICAO_CONHECIMENTO_PROVIDENCIAS  # "providências"
    assert "ã" in DESCRICAO_DILACAO                     # "dilação"
    assert "ã" in DESCRICAO_REITERACAO                  # "reiteração"
    assert "ã" in DESCRICAO_JUIZO_SINGULAR              # "decisão"
    assert "í" in DESCRICAO_JUIZO_SINGULAR              # "juízo"
    assert "/" in DESCRICAO_CONHECIMENTO_PROVIDENCIAS


# ---------------------- normalize_secretaria_label --------------------------

@pytest.mark.parametrize(
    "entrada,saida",
    [
        ("Secretaria Municipal de Educação", SECRETARIA_EDUCACAO),
        ("Secretaria de Educação", SECRETARIA_EDUCACAO),
        ("SME-SP gestão escolar", SECRETARIA_EDUCACAO),
        ("Educação pública municipal", SECRETARIA_EDUCACAO),
        ("Secretaria Municipal da Saúde", SECRETARIA_SAUDE),
        ("Secretaria Municipal de Saúde", SECRETARIA_SAUDE),
        ("Programa SMS imunização", SECRETARIA_SAUDE),
        ("texto generico sem ente publico", SECRETARIA_GERAL),
        ("", SECRETARIA_GERAL),
    ],
)
def test_normalize_secretaria_label(entrada: str, saida: str) -> None:
    assert normalize_secretaria_label(entrada) == saida


# ---------------------- signer_name_matches_roseli_chaves -------------------

@pytest.mark.parametrize(
    "candidato",
    [
        "ROSELI DE MORAIS CHAVES",
        "ROSELI MORAES CHAVES",
        "Roseli Chaves",
        "roseli chaves",
        "Roseli M. Chaves",
        "Sra. Roseli de Moraes Chaves",
        "CHAVES, ROSELI",
        "  Roseli   de Morais   Chaves  ",
    ],
)
def test_signer_match_accepts_roseli_chaves_variations(candidato: str) -> None:
    assert signer_name_matches_roseli_chaves(candidato)


@pytest.mark.parametrize(
    "candidato",
    [
        "Roseli",
        "Chaves",
        "Roseli de Morais",
        "Jose das Chaves",
        "",
        "outro nome qualquer",
    ],
)
def test_signer_match_rejects_partials(candidato: str) -> None:
    assert not signer_name_matches_roseli_chaves(candidato)


def test_required_tokens_default_is_roseli_chaves() -> None:
    assert set(SIGNER_REQUIRED_TOKENS) == {"roseli", "chaves"}


def test_find_signer_in_list_returns_first_match() -> None:
    names = [
        "Jose Silva",
        "Maria Souza",
        "ROSELI DE MORAIS CHAVES",
        "Roseli Chaves",
    ]
    assert find_signer_in_list(names) == "ROSELI DE MORAIS CHAVES"


def test_find_signer_in_list_returns_none_when_no_match() -> None:
    names = ["Jose Silva", "Maria Souza"]
    assert find_signer_in_list(names) is None


# ---------------- SIGNER_MATCH_TOKENS / current_signer_tokens ---------------
# Caso real: Roseli Chaves entrou de ferias em 24/06/2026 e foi substituida
# pela Daniela Shimizu (Subsecretaria-Geral Substituta). O matcher precisa
# trocar de assinante so com env var, sem mexer em codigo.

from oficio_normalize import (  # type: ignore  # noqa: E402
    current_signer_tokens,
    signer_name_matches,
)


def test_current_signer_tokens_default_eh_roseli_chaves(monkeypatch) -> None:
    monkeypatch.delenv("SIGNER_MATCH_TOKENS", raising=False)
    assert current_signer_tokens() == ("roseli", "chaves")


def test_current_signer_tokens_le_env_var_daniela(monkeypatch) -> None:
    monkeypatch.setenv("SIGNER_MATCH_TOKENS", "daniela,shimizu")
    assert current_signer_tokens() == ("daniela", "shimizu")


def test_current_signer_tokens_normaliza_espacos_e_acentos(monkeypatch) -> None:
    monkeypatch.setenv("SIGNER_MATCH_TOKENS", "  DANIELA ; Shimizú  ")
    assert current_signer_tokens() == ("daniela", "shimizu")


def test_current_signer_tokens_envar_vazia_volta_para_default(monkeypatch) -> None:
    monkeypatch.setenv("SIGNER_MATCH_TOKENS", "")
    assert current_signer_tokens() == ("roseli", "chaves")


def test_signer_name_matches_default_continua_aceitando_roseli(monkeypatch) -> None:
    monkeypatch.delenv("SIGNER_MATCH_TOKENS", raising=False)
    assert signer_name_matches("Roseli de Morais Chaves")
    assert signer_name_matches("ROSELI MORAES CHAVES")
    assert not signer_name_matches("Daniela Shimizu")


def test_signer_name_matches_com_env_daniela_aceita_daniela(monkeypatch) -> None:
    monkeypatch.setenv("SIGNER_MATCH_TOKENS", "daniela,shimizu")
    assert signer_name_matches("Daniela Shimizu")
    assert signer_name_matches("DANIELA K. SHIMIZU")
    assert signer_name_matches("Sra. Daniela Shimizu - Subsecretaria-Geral Substituta")
    # nao confunde com Roseli
    assert not signer_name_matches("Roseli de Morais Chaves")
    # exige AMBOS tokens — so o primeiro nome nao basta
    assert not signer_name_matches("Daniela Souza")


def test_signer_name_matches_legado_intacto() -> None:
    """A funcao legada signer_name_matches_roseli_chaves NAO le env var
    e continua presa a Roseli/Chaves (50+ testes existentes garantem
    estabilidade)."""
    import os
    os.environ["SIGNER_MATCH_TOKENS"] = "daniela,shimizu"
    try:
        assert signer_name_matches_roseli_chaves("Roseli Chaves")
        assert not signer_name_matches_roseli_chaves("Daniela Shimizu")
    finally:
        os.environ.pop("SIGNER_MATCH_TOKENS", None)
