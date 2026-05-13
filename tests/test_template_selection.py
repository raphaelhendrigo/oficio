"""Testes para a selecao automatica de modelo em main.py.

Verifica que classify_and_select_template_path produz o arquivo correto
de acordo com:
  - tipo (UTAP / DILACAO / REITERACAO / JUIZO) — derivado de palavras-chave
    no texto e/ou no nome da ultima peca;
  - secretaria (Educacao / Saude / Geral) — derivada do texto do PDF.
"""
from __future__ import annotations

from pathlib import Path

import pytest

import main as fluxo_atos  # type: ignore  # noqa: E402


# --------------------------- Helpers do main --------------------------------

def _classify(text: str, last_piece: str | None = None, cover: str | None = None) -> str:
    return fluxo_atos._classify_tipo_from_text_and_piece(text, last_piece, cover_text=cover)


def _select(pdf_text: str, last_piece: str | None = None, cover: str | None = None):
    return fluxo_atos.classify_and_select_template_path(pdf_text, last_piece, cover_text=cover)


# --------------------------- Classificacao por tipo -------------------------

def test_classify_juizo_by_piece_name() -> None:
    assert _classify("texto qualquer", "Decisao Juizo Singular - publicada") == "JUIZO"


def test_classify_utap_by_piece_name_manutap() -> None:
    assert _classify("texto qualquer", "MANUTAP-OF 1234") == "UTAP"


def test_classify_dilacao_by_keywords() -> None:
    txt = "O Relator autorizo a dilacao de prazo solicitada pelo orgao"
    assert _classify(txt, None) == "DILACAO"


def test_classify_reiteracao_by_keywords() -> None:
    txt = "Considerando o tempo decorrido, reitere-se o oficio anterior"
    assert _classify(txt, None) == "REITERACAO"


def test_classify_utap_fallback() -> None:
    # Sem palavras-chave especificas; o fallback conservador e UTAP.
    assert _classify("texto sem nada", None) == "UTAP"


# --------------------------- Selecao por secretaria -------------------------

def test_select_educacao_template(project_root: Path) -> None:
    txt = "Secretaria Municipal de Educacao - SME - acompanhamento"
    path = _select(txt)
    assert path is not None
    assert "Educa" in path.name, f"Esperava modelo Educacao, veio: {path.name}"


def test_select_saude_template(project_root: Path) -> None:
    txt = "Secretaria Municipal da Saude - SMS - imunizacao"
    path = _select(txt)
    assert path is not None
    assert "Sa" in path.name, f"Esperava modelo Saude, veio: {path.name}"


def test_select_geral_template_when_no_secretaria(project_root: Path) -> None:
    txt = "Texto generico sem indicacao de orgao"
    path = _select(txt)
    assert path is not None
    assert "Geral" in path.name, f"Esperava modelo Geral, veio: {path.name}"


def test_select_juizo_template_routes_to_modelos_juizo(project_root: Path) -> None:
    txt = "Texto com decisao de juizo singular do conselheiro"
    path = _select(txt, last_piece=None)
    assert path is not None
    # Deve cair em modelos_juizo (pasta dedicada)
    assert "modelos_juizo" in str(path), f"Esperava modelos_juizo, veio: {path}"


def test_select_dilacao_template_routes_to_modelos_dilacao(project_root: Path) -> None:
    txt = "Conselheiro autorizo a dilacao de prazo do processo"
    path = _select(txt, last_piece=None)
    assert path is not None
    assert "modelos_dilacao" in str(path), f"Esperava modelos_dilacao, veio: {path}"


def test_select_reiteracao_template_routes_to_modelos_reiteracao(project_root: Path) -> None:
    txt = "reitere-se o oficio considerando o tempo decorrido"
    path = _select(txt, last_piece=None)
    assert path is not None
    assert "modelos_reiteracao" in str(path), f"Esperava modelos_reiteracao, veio: {path}"
