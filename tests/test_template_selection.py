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
    assert "Educa" not in path.name, f"Modelo Saúde nunca deve retornar Educação: {path.name}"


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



# --------------------------- Overrides via env vars ----------------------------
# Caso real: ferias da Roseli Chaves (24/06/2026) - as pastas de modelos
# 'modelos daniela/{manutap,dilação,reiteração}' substituem as historicas
# via env vars OFICIO_TEMPLATES_DIR_<TIPO>. Pasta da Daniela so tem Educacao
# e Saude, entao OFICIO_SECRETARIA_FALLBACK=educacao cobre o caso 'Geral'.

def test_select_template_usa_override_de_env_var_para_utap(project_root: Path, monkeypatch, tmp_path) -> None:
    # Cria uma pasta temporaria com modelos da Daniela
    d = tmp_path / "daniela_manutap"
    d.mkdir()
    (d / "SSG - Aposentadoria MANUTAP - Educacao - Daniela.docx").write_bytes(b"fake")
    (d / "SSG - Aposentadoria MANUTAP - Saude - Daniela.docx").write_bytes(b"fake")
    monkeypatch.setenv("OFICIO_TEMPLATES_DIR_UTAP", str(d))
    txt = "Secretaria Municipal de Educacao"
    path = _select(txt, last_piece="MANUTAP-OF 123/2026")
    assert path is not None
    assert "daniela_manutap" in str(path).replace("\\", "/")
    assert "Educa" in path.name


def test_select_template_fallback_educacao_quando_pasta_so_tem_educacao_saude(
    project_root: Path, monkeypatch, tmp_path
) -> None:
    """Pasta da Daniela so tem Educacao e Saude. Veio um processo Geral.
    Sem fallback, antes caia em qualquer arquivo. Com OFICIO_SECRETARIA_FALLBACK=educacao,
    cai explicitamente em Educacao."""
    d = tmp_path / "daniela_dilacao"
    d.mkdir()
    (d / "SSG - Aposentadoria Dilacao - Educacao - Daniela.dotx").write_bytes(b"fake")
    (d / "SSG - Aposentadoria Dilacao - Saude - Daniela.dotx").write_bytes(b"fake")
    monkeypatch.setenv("OFICIO_TEMPLATES_DIR_DILACAO", str(d))
    monkeypatch.setenv("OFICIO_SECRETARIA_FALLBACK", "educacao")
    txt = "Conselheiro autorizo a dilacao de prazo do processo (orgao generico)"
    path = _select(txt, last_piece=None)
    assert path is not None
    assert "Educa" in path.name, f"Fallback deveria ser Educacao, veio {path.name}"


def test_select_template_override_para_reiteracao(
    project_root: Path, monkeypatch, tmp_path
) -> None:
    d = tmp_path / "daniela_reit"
    d.mkdir()
    (d / "SSG - Aposentadoria Reiteracao - Saude - Daniela.dotx").write_bytes(b"fake")
    monkeypatch.setenv("OFICIO_TEMPLATES_DIR_REITERACAO", str(d))
    txt = "reitere-se o oficio - secretaria municipal da saude"
    path = _select(txt, last_piece=None)
    assert path is not None
    assert "daniela_reit" in str(path).replace("\\", "/")
