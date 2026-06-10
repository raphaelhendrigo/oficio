from __future__ import annotations

from pathlib import Path

import pytest

import main  # type: ignore

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "modelos_dilacao" / "SSG - Aposentadoria Dilação - Educação.dotx"


PIECES = [
    (2, "2. PLAAN - 13164/2025 - 21/10/2025 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
    (3, "3. MANUTAP-OF - 3624/2025 - 21/10/2025 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
    (4, "4. OF SSG - 19547/2025 - 30/10/2025 - UNIDADE TÉCNICA DE OFÍCIOS"),
    (9, "9. OF SSG - 13808/2026 - 02/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
    (13, "13. REQUERIMENTO - DILAÇÃO DE PRAZO - 006266/2026"),
]


def test_ssg_ref_comes_from_piece_after_last_manutap():
    # Deve pegar a peça 4 (logo após MANUTAP-OF), NÃO o OF SSG posterior (peça 9).
    assert main._extract_ssg_ref_after_manutap(PIECES) == "19547/2025"


def test_ssg_ref_empty_without_manutap():
    assert main._extract_ssg_ref_after_manutap([(1, "1. TÍTULO")]) == ""


def test_referencia_regex_picks_header_oficio_not_body():
    import re

    text = (
        "Ofício nº 818/2026 - SME / COGEP / DITEM\n"
        "Tribunal de Contas do Município de São Paulo\n"
        "Senhor Presidente,\n"
        "...estabelecido através do ofício SSG nº 13808/2026 – Processo TC/006237/2023..."
    )
    m = re.search(r"Of[ií]cio\s+n[º°\.º°]+\s*\d+\s*/\s*\d{4}[^\n\r]*", text, re.I)
    assert m is not None
    assert re.sub(r"\s+", " ", m.group(0)).strip() == "Ofício nº 818/2026 - SME / COGEP / DITEM"


def test_reiteracao_uses_last_piece_when_des_present():
    """O Encaminha usa a ultima MANUTAP e a ultima peca anexada."""
    pieces = [
        (4, "4. MANUTAP-OF - 886/2026 - 05/03/2026 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
        (5, "5. OF SSG - 14354/2026 - 17/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (8, "8. DES - 764/2026 - 26/05/2026 - GABINETE CONSELHEIRO JOAO ANTONIO"),
        (9, "9. ENC - 828/2026 - 26/05/2026 - GABINETE DO CONSELHEIRO"),
    ]
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "4"
    assert r["last_seq"] == "9"
    assert r["des_seq"] == "9", "des_seq e alias legado da ultima peca"
    assert r["ssg_ref"] == "14354/2026"


def test_reiteracao_uses_last_piece_without_des():
    """Sem DES, a regra continua sendo a ultima peca anexada."""
    pieces = [
        (4, "4. MANUTAP-OF - 886/2026 - 05/03/2026 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
        (5, "5. OF SSG - 14354/2026 - 17/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (7, "7. ENC - 2200/2026 - 23/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (9, "9. ENC - 828/2026 - 26/05/2026 - GABINETE DO CONSELHEIRO"),
    ]
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "4"
    assert r["last_seq"] == "9"
    assert r["des_seq"] == "9"
    assert r["ssg_ref"] == "14354/2026"


def test_reiteracao_uses_last_piece_even_from_other_units():
    """A ultima peca pode ser de qualquer tipo/unidade."""
    pieces = [
        (4, "4. MANUTAP-OF - 886/2026 - 05/03/2026 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
        (5, "5. OF SSG - 14354/2026 - 17/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (7, "7. ENC - 2200/2026 - 23/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
    ]
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "4"
    assert r["last_seq"] == "7"
    assert r["des_seq"] == "7"


def test_reiteracao_last_piece_uses_display_number_not_index_order():
    """O e-TCM pode devolver index_ato fora da ordem visual da arvore."""
    pieces = [
        (20, "7. ENC - 2200/2026 - 23/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (4, "4. MANUTAP-OF - 886/2026 - 05/03/2026 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
        (5, "5. OF SSG - 14354/2026 - 17/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (3, "9. DES - 764/2026 - 26/05/2026 - GABINETE CONSELHEIRO JOAO ANTONIO"),
    ]
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "4"
    assert r["last_seq"] == "9"
    assert r["des_seq"] == "9"


@pytest.mark.skipif(not TEMPLATE.exists(), reason="modelo de dilação ausente")
def test_set_oficio_ssg_ref_preserves_prazo_bold(tmp_path: Path):
    """O run-replace do SSG NÃO pode destruir o negrito de
    '60 (sessenta) dias corridos' que vive no mesmo parágrafo."""
    from docx import Document  # type: ignore
    from docx_utils import convert_dotx_to_docx_preserving_layout  # type: ignore

    out = tmp_path / "oficio.docx"
    convert_dotx_to_docx_preserving_layout(TEMPLATE, out)

    assert main.set_oficio_ssg_ref(out, "19547/2025") is True

    doc = Document(str(out))
    body = next(p for p in doc.paragraphs if "cumprimento ao despacho" in (p.text or ""))
    assert "Ofício SSG 19547/2025" in body.text
    assert "XXXX" not in body.text
    bold_runs = [r.text for r in body.runs if r.bold and r.text.strip()]
    assert any("60 (sessenta) dias corridos" in t for t in bold_runs)
