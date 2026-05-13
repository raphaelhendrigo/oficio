"""Testes para src/docx_utils.py.

Foco: garantir que tokens @@... dos modelos oficiais cheguem intactos ao
DOCX gerado, porque o preenchimento e responsabilidade do e-TCM apos o upload.
"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from docx_utils import (  # type: ignore  # noqa: E402  (importado via conftest)
    AT_TOKEN_RE,
    AtTokenReport,
    AtTokenViolation,
    assert_at_tokens_preserved,
    convert_dotx_to_docx_preserving_layout,
    extract_at_tokens_from_docx,
    filter_out_at_tokens,
    is_at_token_key,
    safe_replace_non_at_placeholders,
)


# --------------------------- AT_TOKEN_RE ------------------------------------

def test_at_token_re_matches_expected_shapes() -> None:
    """Regex deve casar exatamente os formatos pedidos no brief."""
    sample = (
        "@@processo @@data_extenso @@numero_oficio @@Nome_interessado "
        "@@processoexterno @@nome_relator @@Tipo_Processo @@natureza_processo "
        "@@instancia @@assinador_role @@dia_extenso_pt_BR"
    )
    tokens = set(AT_TOKEN_RE.findall(sample))
    expected = {
        "@@processo", "@@data_extenso", "@@numero_oficio", "@@Nome_interessado",
        "@@processoexterno", "@@nome_relator", "@@Tipo_Processo", "@@natureza_processo",
        "@@instancia", "@@assinador_role", "@@dia_extenso_pt_BR",
    }
    assert tokens == expected


def test_at_token_re_does_not_match_at_alone_or_single_at() -> None:
    assert AT_TOKEN_RE.findall("@processo") == []
    assert AT_TOKEN_RE.findall("@@ ") == []
    assert AT_TOKEN_RE.findall("@@") == []


def test_at_token_re_accepts_latin_accents() -> None:
    assert AT_TOKEN_RE.findall("@@orgaopublico @@instância @@órgão") == [
        "@@orgaopublico", "@@instância", "@@órgão",
    ]


# ----------------------- extract_at_tokens_from_docx ------------------------

def test_extract_at_tokens_from_modelo_utap_educacao(modelo_utap_educacao: Path) -> None:
    tokens = extract_at_tokens_from_docx(modelo_utap_educacao)
    # Os modelos oficiais usam alguns desses tokens — verificamos um nucleo
    # comum (todos terao ao menos esses).
    must_have = {
        "@@numero_oficio",
        "@@data_extenso",
        "@@processo",
        "@@Nome_interessado",
        "@@processoexterno",
        "@@nome_relator",
    }
    missing = must_have - tokens
    assert not missing, f"Tokens esperados ausentes no modelo Educacao: {missing} (presentes={sorted(tokens)})"


def test_extract_at_tokens_handles_dotx(modelo_dilacao_geral: Path) -> None:
    tokens = extract_at_tokens_from_docx(modelo_dilacao_geral)
    # DOTX deve produzir o mesmo tipo de set
    assert isinstance(tokens, set)
    assert all(t.startswith("@@") for t in tokens)
    # Esperamos pelo menos um token relevante
    assert tokens, "Esperava tokens @@ no modelo DOTX de dilacao, mas nenhum foi extraido"


def test_extract_at_tokens_raises_on_missing_file(tmp_path: Path) -> None:
    inexistente = tmp_path / "nao_existe.docx"
    with pytest.raises(FileNotFoundError):
        extract_at_tokens_from_docx(inexistente)


# ----------------------- assert_at_tokens_preserved -------------------------

def test_assert_at_tokens_preserved_when_template_copied(modelo_utap_geral: Path, tmp_path: Path) -> None:
    """Copiar o modelo byte-a-byte deve manter todos os @@ intactos."""
    out = tmp_path / "copia.docx"
    shutil.copyfile(modelo_utap_geral, out)
    report = assert_at_tokens_preserved(modelo_utap_geral, out)
    assert isinstance(report, AtTokenReport)
    assert report.ok
    assert report.missing_in_generated == []
    assert report.template_tokens == report.generated_tokens


def test_assert_at_tokens_preserved_fails_when_token_replaced(modelo_utap_geral: Path, tmp_path: Path) -> None:
    """Se algum @@token sumir do gerado, deve levantar AtTokenViolation.

    Estrategia: escolhe um token que aparece intacto no XML cru (ou seja, NAO
    esta em split-run) e o substitui em todas as partes XML do pacote, depois
    valida que extract_at_tokens_from_docx confirma a remocao antes de
    chamar assert_at_tokens_preserved.
    """
    import zipfile

    tokens_orig = extract_at_tokens_from_docx(modelo_utap_geral)
    assert tokens_orig, "Modelo de teste nao possui tokens @@"

    # Procura no XML cru por um token que esteja INTACTO (sem split-run),
    # para que str.replace consiga removê-lo.
    with zipfile.ZipFile(modelo_utap_geral, "r") as zin:
        intactos: list[str] = []
        for name in zin.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            xml = zin.read(name).decode("utf-8", errors="ignore")
            for tok in tokens_orig:
                if tok in xml and tok not in intactos:
                    intactos.append(tok)
    if not intactos:
        pytest.skip("Modelo nao possui @@token intacto no XML cru para mutilacao simples")
    alvo = intactos[0]

    out = tmp_path / "mutilado.docx"
    with zipfile.ZipFile(modelo_utap_geral, "r") as zin, zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                xml = data.decode("utf-8", errors="ignore").replace(alvo, "VALOR_REAL")
                data = xml.encode("utf-8")
            zout.writestr(info, data)

    # Confirma que o token efetivamente sumiu antes de testar a validacao.
    tokens_apos = extract_at_tokens_from_docx(out)
    assert alvo not in tokens_apos, (
        f"Pre-condicao falhou: token {alvo} ainda presente apos mutilacao (presentes={sorted(tokens_apos)})"
    )

    with pytest.raises(AtTokenViolation) as exc:
        assert_at_tokens_preserved(modelo_utap_geral, out)
    assert alvo in str(exc.value)


# ------------------ safe_replace_non_at_placeholders ------------------------

def test_safe_replace_rejects_at_tokens(modelo_utap_geral: Path, tmp_path: Path) -> None:
    out = tmp_path / "rejeicao.docx"
    shutil.copyfile(modelo_utap_geral, out)
    bad = {"@@processo": "TC/000000/2026", "{CARGO}": "Secretario"}
    with pytest.raises(AtTokenViolation) as exc:
        safe_replace_non_at_placeholders(out, bad, template_path=modelo_utap_geral)
    assert "@@processo" in str(exc.value)


def test_safe_replace_no_at_keys_keeps_tokens(modelo_utap_geral: Path, tmp_path: Path) -> None:
    """Aplicar replacements nao-@@ deve preservar 100% dos @@ do template."""
    out = tmp_path / "geracao.docx"
    shutil.copyfile(modelo_utap_geral, out)
    # Esses placeholders podem nao existir no modelo (e tudo bem). O teste
    # garante que a chamada nao toca em @@ e que o relatorio final passa.
    replacements = {
        "{CARGO}": "Secretario Municipal",
        "{NOME}": "Fulano de Tal",
        "{ORGAO}": "Prefeitura de Sao Paulo",
        "{ENDERECO}": "Rua X, 123",
    }
    report = safe_replace_non_at_placeholders(out, replacements, template_path=modelo_utap_geral)
    assert report.ok
    assert report.missing_in_generated == []


def test_safe_replace_empty_replacements_is_valid(modelo_utap_geral: Path, tmp_path: Path) -> None:
    out = tmp_path / "copia_pura.docx"
    shutil.copyfile(modelo_utap_geral, out)
    report = safe_replace_non_at_placeholders(out, {}, template_path=modelo_utap_geral)
    assert report.ok


# ------------------ convert_dotx_to_docx_preserving_layout ------------------

def test_convert_dotx_preserves_tokens(modelo_dilacao_geral: Path, tmp_path: Path) -> None:
    out = tmp_path / "convertido.docx"
    convert_dotx_to_docx_preserving_layout(modelo_dilacao_geral, out)
    assert out.exists()
    tokens_in = extract_at_tokens_from_docx(modelo_dilacao_geral)
    tokens_out = extract_at_tokens_from_docx(out)
    assert tokens_in == tokens_out, (
        f"Conversao DOTX->DOCX deveria preservar todos os @@. "
        f"Faltando={sorted(tokens_in - tokens_out)}; Extras={sorted(tokens_out - tokens_in)}"
    )


def test_convert_docx_to_docx_is_pure_copy(modelo_utap_geral: Path, tmp_path: Path) -> None:
    out = tmp_path / "copia_docx.docx"
    convert_dotx_to_docx_preserving_layout(modelo_utap_geral, out)
    assert out.exists()
    tokens_in = extract_at_tokens_from_docx(modelo_utap_geral)
    tokens_out = extract_at_tokens_from_docx(out)
    assert tokens_in == tokens_out


# ------------------------ Helpers de filtro/sanidade ------------------------

def test_filter_out_at_tokens_removes_only_at_keys() -> None:
    raw = {
        "@@processo": "TC/1",
        "@@data_extenso": "x",
        "{CARGO}": "Secret.",
        "{{NOME}}": "Fulano",
        "destinatario": "Ente publico",
    }
    cleaned = filter_out_at_tokens(raw)
    assert "@@processo" not in cleaned
    assert "@@data_extenso" not in cleaned
    assert cleaned["{CARGO}"] == "Secret."
    assert cleaned["{{NOME}}"] == "Fulano"
    assert cleaned["destinatario"] == "Ente publico"


def test_is_at_token_key() -> None:
    assert is_at_token_key("@@processo")
    assert not is_at_token_key("processo")
    assert not is_at_token_key("@processo")
    assert not is_at_token_key("")
