"""Testes end-to-end para main.generate_oficio_from_template.

Foco: confirmar que, contra os modelos reais do projeto (modelos_utap,
modelos_dilacao, modelos_reiteracao, modelos_juizo), a geracao preserva
todos os tokens @@... mesmo quando o chamador tenta forcar uma chave @@
no `extra`.
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest

# main.py depende de pacotes do projeto (.dotenv, playwright) — importamos
# tardio para que conftest.py ja tenha inserido src/ em sys.path.
import main as fluxo_atos  # type: ignore  # noqa: E402
from docx_utils import (  # type: ignore  # noqa: E402
    assert_encaminha_has_piece_number,
    assert_encaminha_text_not_bold,
    assert_euclides_marker_present_once,
    extract_at_tokens_from_docx,
    extract_visible_text_from_docx,
)


def _gen_into(tmp_path: Path, template: Path, extra: dict | None = None) -> Path | None:
    return fluxo_atos.generate_oficio_from_template(
        processo="TC/000999/2026",
        output_dir=tmp_path,
        extra=extra,
        template_path=template,
    )


@pytest.fixture(autouse=True)
def _isolate_production_docx_flags(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA", raising=False)
    monkeypatch.delenv("OFICIO_ADD_EUCLIDES_MARKER", raising=False)
    monkeypatch.delenv("OFICIO_EUCLIDES_MARKER", raising=False)


def test_generate_from_utap_geral_preserves_at_tokens(modelo_utap_geral: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "true")
    out = _gen_into(tmp_path, modelo_utap_geral, extra=None)
    assert out is not None, "Gerador deveria produzir DOCX a partir do modelo Geral"
    assert out.exists()
    tokens_modelo = extract_at_tokens_from_docx(modelo_utap_geral)
    tokens_gerado = extract_at_tokens_from_docx(out)
    missing = tokens_modelo - tokens_gerado
    assert not missing, f"Tokens @@ ausentes apos geracao: {missing}"


def test_generate_from_utap_educacao_preserves_at_tokens(modelo_utap_educacao: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "true")
    out = _gen_into(tmp_path, modelo_utap_educacao, extra=None)
    assert out is not None
    tokens_modelo = extract_at_tokens_from_docx(modelo_utap_educacao)
    tokens_gerado = extract_at_tokens_from_docx(out)
    assert (tokens_modelo - tokens_gerado) == set()


def test_generate_from_utap_saude_preserves_at_tokens(modelo_utap_saude: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "true")
    out = _gen_into(tmp_path, modelo_utap_saude, extra=None)
    assert out is not None
    tokens_modelo = extract_at_tokens_from_docx(modelo_utap_saude)
    tokens_gerado = extract_at_tokens_from_docx(out)
    assert (tokens_modelo - tokens_gerado) == set()


def test_generate_filters_at_keys_from_extra(modelo_utap_geral: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Mesmo que o chamador insira chaves @@... no `extra`, elas devem ser
    filtradas antes do replace (preservando os tokens no DOCX final)."""
    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "true")
    extra = {
        "@@processo": "TC/000000/2099",   # NUNCA deve sobrescrever no DOCX
        "@@numero_oficio": "1234",        # idem
        "@@nome_relator": "Fulano",       # idem
        "_meta_nome_relator": "Conselheiro Teste",
        "{{ASSUNTO}}": "Assunto de teste",  # placeholder permitido
    }
    out = _gen_into(tmp_path, modelo_utap_geral, extra=extra)
    assert out is not None
    tokens_modelo = extract_at_tokens_from_docx(modelo_utap_geral)
    tokens_gerado = extract_at_tokens_from_docx(out)
    # Todos os @@ do modelo devem ter sobrevivido
    assert (tokens_modelo - tokens_gerado) == set(), (
        f"@@ ausentes no gerado: {sorted(tokens_modelo - tokens_gerado)}"
    )


def test_generate_blocks_when_violation_detected(modelo_utap_geral: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Quando OFICIO_PRESERVE_AT_TOKENS=false, mudancas que violam @@ apenas
    avisam. Quando true (default), o helper retorna None. Aqui forcamos uma
    violacao via monkeypatch de assert_at_tokens_preserved.
    """
    from docx_utils import AtTokenViolation

    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "true")

    def fake_assert(template, generated):
        raise AtTokenViolation("simulada para teste")

    monkeypatch.setattr(fluxo_atos, "assert_at_tokens_preserved", fake_assert)
    out = _gen_into(tmp_path, modelo_utap_geral, extra=None)
    assert out is None, "Gerador deveria retornar None quando preservacao falha"

    # Modo permissivo: devolve mesmo com violacao simulada
    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "false")
    out2 = _gen_into(tmp_path, modelo_utap_geral, extra=None)
    assert out2 is not None


def test_generate_final_docx_replaces_encaminha_and_adds_euclides(
    modelo_utap_geral: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    template_text = "\n".join(extract_visible_text_from_docx(modelo_utap_geral))
    if "Cópia da(s) peça(s) dos autos" not in template_text:
        pytest.skip("Modelo de teste não contém a linha genérica de Encaminha")

    monkeypatch.setenv("OFICIO_PRESERVE_AT_TOKENS", "true")
    monkeypatch.setenv("OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA", "true")
    monkeypatch.setenv("OFICIO_ADD_EUCLIDES_MARKER", "true")
    # Brief 2026-05-14: barra correta é "/", nunca "\".
    monkeypatch.setenv("OFICIO_EUCLIDES_MARKER", "/euclides")
    out = _gen_into(
        tmp_path,
        modelo_utap_geral,
        extra={
            "Cópia da(s) peça(s) dos autos.": "Cópia da peça 03 dos autos.",
            "{{ENCAMINHA}}": "Cópia da peça 03 dos autos.",
        },
    )
    assert out is not None
    assert_encaminha_has_piece_number(out)
    assert_encaminha_text_not_bold(out)
    assert_euclides_marker_present_once(out)
    final_text = "\n".join(extract_visible_text_from_docx(out))
    assert "Cópia da(s) peça(s) dos autos" not in final_text
