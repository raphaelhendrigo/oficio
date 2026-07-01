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


def test_extract_ssg_ref_from_ato_text_ignores_tc_process_number():
    text = "TC/008940/2023\tOficio SSG\t15777/2026\tConcluido"
    assert main._extract_ssg_ref_from_ato_text(text) == "15777/2026"


def test_extract_ssg_ref_from_ato_text_returns_empty_without_ssg_number():
    text = "TC/008940/2023\tOficio SSG\tConcluido"
    assert main._extract_ssg_ref_from_ato_text(text) == ""


def test_extract_ssg_ref_from_ato_text_accepts_number_before_label():
    text = "15777/2026\tOficio SSG\tConcluido\tTC/008940/2023"
    assert main._extract_ssg_ref_from_ato_text(text) == "15777/2026"


def test_upload_ato_tipo_inferido_por_nome_do_docx():
    assert main._infer_upload_ato_tipo_label(Path("encaminhamento_TC_008940_2023.docx")) == "Encaminhamento"
    assert main._infer_upload_ato_tipo_label(Path("oficio_TC_008940_2023.docx")) == "Ofício SSG"


def test_upload_ato_tipo_override_explicito():
    assert main._infer_upload_ato_tipo_label(Path("qualquer.docx"), ato_tipo="Encaminhamento") == "Encaminhamento"


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


# pieces[] vinda de _enumerate_piece_names() ja vem ordenada pela árvore
# visível do processo. Os testes abaixo passam pieces como se fossem o
# espelho fiel da arvore: o numero a esquerda do . no name reflete a
# posicao na lista — porque eh isso que o operador vê e usa como
# "Cópia das peças X e Y dos autos.". O iv (1o membro da tupla) eh o
# atributo HTML index_ato, que pode estar esparso por exclusoes anteriores
# e NAO deve mais ser usado como numero de display.

def _build_pieces(*titles_after_manutap) -> list[tuple[int, str]]:
    """Constroi pieces[] simulando uma árvore tipica:
    0..3 = CAPA/TITULO/PLAAN/MANUTAP-OF
    4 = OF SSG da reiteracao (e os parametros sao as peças seguintes).
    Os iv pulam 100 a cada peça apos a MANUTAP, simulando o cenario real
    de exclusoes (display position != iv)."""
    base = [
        (0, "0. CAPA"),
        (1, "1. TÍTULO DE APOSENTADORIA"),
        (2, "2. PLAAN - 2495/2026 - 05/03/2026 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
        (3, "3. MANUTAP-OF - 886/2026 - 05/03/2026 - UNIDADE TÉCNICA DE APOSENTADORIA E PENSÕES"),
        (4, "4. OF SSG - 14354/2026 - 17/03/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
    ]
    out = list(base)
    iv = 100  # iv esparso de proposito para validar que nao usamos mais iv
    for t in titles_after_manutap:
        out.append((iv, t))
        iv += 100
    return out


def test_reiteracao_usa_posicao_quando_ultima_peca_e_des():
    """Última peça é um DES (despacho do conselheiro). des_seq deve ser a
    POSIÇÃO na árvore (display), não o iv."""
    pieces = _build_pieces(
        "5. PTC - 100/2026 - 10/04/2026 - UNIDADE TÉCNICA DE OFÍCIOS",
        "6. ENC - 200/2026 - 10/04/2026 - UNIDADE TÉCNICA DE OFÍCIOS",
        "7. INF - 300/2026 - 15/05/2026 - UNIDADE TÉCNICA DE CARTÓRIO",
        "8. DES - 764/2026 - 26/05/2026 - GABINETE CONSELHEIRO JOAO ANTONIO",
    )
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "3"
    assert r["ssg_seq"] == "4"
    assert r["des_seq"] == "8", "DES esta na posicao 8 da arvore (display), nao no iv"
    assert r["ssg_ref"] == "14354/2026"


def test_reiteracao_usa_posicao_quando_ultima_peca_e_enc():
    """Última peça é um ENC (qualquer unidade — sem inspeção de tipo)."""
    pieces = _build_pieces(
        "5. PTC - 100/2026 - 10/04/2026 - UNIDADE TÉCNICA DE OFÍCIOS",
        "6. ENC - 200/2026 - 10/04/2026 - UNIDADE TÉCNICA DE OFÍCIOS",
        "7. ENC - 828/2026 - 26/05/2026 - GABINETE DO CONSELHEIRO",
    )
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "3"
    assert r["des_seq"] == "7"


def test_reiteracao_usa_posicao_quando_ultima_peca_e_qualquer_outro_tipo():
    """A nova regra ignora o tipo — INF, PTC ou qualquer outro pode ser a
    "2a peça do Encaminha" se for a ultima da arvore."""
    pieces = _build_pieces(
        "5. INF - 3922/2026 - 25/05/2026 - UNIDADE TÉCNICA DE CARTÓRIO",
    )
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "3"
    assert r["des_seq"] == "5", "ultima peça (INF) vai pro Encaminha de qualquer jeito"


def test_reiteracao_simula_arvore_com_iv_esparso_de_tc_016628_2024():
    """Caso real observado em 29/05/2026 com TC/016628/2024:
       - arvore visivel mostra 15 peças (0..14)
       - mas iv da ultima peça eh 16 (atos antigos excluidos)
       - operador ve "14. ENC - 451/2026 - GABINETE RICARDO TORRES"
       - oficio deve dizer "Cópia das peças 03 e 14 dos autos."
    Antes do fix dava "03 e 16". Apos o fix da pos-based: "03 e 14"."""
    pieces = [
        (0, "0. CAPA"),
        (1, "1. TÍTULO DE APOSENTADORIA"),
        (2, "2. PLAAN - 14553/2025 - 28/11/2025"),
        (3, "3. MANUTAP-OF - 5027/2025 - 28/11/2025"),
        (4, "4. OF SSG - 12046/2026 - 08/01/2026 - UNIDADE TÉCNICA DE OFÍCIOS"),
        (5, "5. PTC - 134/2026 - 14/01/2026"),
        (6, "6. ENC - 45/2026 - 14/01/2026"),
        (7, "7. REQUERIMENTO - DILAÇÃO DE PRAZO - 002659/2026"),
        (8, "8. INF - 1642/2026 - 25/03/2026"),
        (9, "9. DES - 552/2026 - 26/03/2026 - GABINETE DO CONSELHEIRO RICARDO TORRES"),
        (10, "10. OF SSG - 15088/2026 - 06/04/2026"),
        (11, "11. PTC - 3577/2026 - 10/04/2026"),
        (12, "12. ENC - 2921/2026 - 10/04/2026"),
        (13, "13. INF - 3520/2026 - 14/05/2026"),
        (16, "14. ENC - 451/2026 - 29/05/2026 - GABINETE DO CONSELHEIRO RICARDO TORRES"),
    ]
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "3"
    assert r["ssg_seq"] == "4"
    assert r["ssg_ref"] == "12046/2026"
    assert r["des_seq"] == "14", (
        "ultima peça da arvore esta na posicao 14, mesmo que seu iv seja 16"
    )


def test_reiteracao_usa_display_do_NOME_quando_enumeracao_perde_uma_peca():
    """Caso real TC/005666/2022 (22/06/2026): a árvore exibe 38 peças
    (0..37), mas _enumerate_piece_names devolve só 37 — algum item (ex.:
    "1. TITULO DE APOSENTADORIA") não tem link clicável e é filtrado. Com
    isso, todas as posições de array depois do item perdido ficam -1 em
    relação ao número exibido na árvore.

    Robô gerou "Cópia das peças 21 e 36 dos autos" quando o correto era
    "22 e 37". O fix passa a parsear o número de display direto do NOME
    da peça (sempre presente em "<num>. <TIPO> - ..."), eliminando a
    dependência de posição do array.
    """
    # Simula 37 peças capturadas (uma a menos que as 38 da árvore):
    # display=1 (TÍTULO) foi filtrado mas o resto mantém o prefixo
    # numérico correto.
    pieces = [
        (0, "0. CAPA"),
        # (1, "1. TITULO ...") ausente: nao foi capturado
        (2, "2. CERTIDÃO"),
        (3, "3. PREXDIG - 5254/2023"),
        (4, "4. TERDIG - 5105/2023"),
        (5, "5. ATRIB - 141/2025"),
        (6, "6. DEMREMUN - 158/2025"),
        (7, "7. HOLE - 270/2025"),
        (8, "8. HOLE - 268/2025"),
        (9, "9. HOLE - 269/2025"),
        (10, "10. PLAAN - 1504/2025"),
        (11, "11. MANUTAP-OF - 214/2025 - 17/02/2025"),
        (12, "12. OF SSG - 12843/2025"),
        (13, "13. E-MAIL - 189/2025"),
        (14, "14. ENC - 805/2025"),
        (15, "15. PTC - 171/2025"),
        (16, "16. INF - 1443/2025"),
        (17, "17. ENC - 6319/2025"),
        (18, "18. OF SSG - 14629/2025"),
        (19, "19. PTC - 3079/2025"),
        (20, "20. RESPOSTA DE COMUNICAÇÃO PROCESSUAL"),
        (21, "21. ENC - 2442/2025"),
        (22, "22. MANUTAP-OF - 1907/2025 - 24/09/2025"),  # LAST MANUTAP — display=22
        (23, "23. OF SSG - 18775/2025 - 03/10/2025"),     # SSG after manutap — display=23
        (24, "24. PTC - 8137/2025"),
        (25, "25. ENC - 6319/2025"),
        (26, "26. INF - 6168/2025"),
        (27, "27. ENC - 548/2025"),
        (28, "28. OF SSG - 21343/2025"),
        (29, "29. PTC - 10963/2025"),
        (30, "30. ENC - 8803/2025"),
        (31, "31. INF - 2050/2026"),
        (32, "32. DES - 536/2026"),
        (33, "33. OF SSG - 14922/2026"),
        (34, "34. PTC - 3397/2026"),
        (35, "35. ENC - 2749/2026"),
        (36, "36. INF - 4432/2026"),
        (37, "37. ENC - 495/2026 - 18/06/2026"),  # LAST piece — display=37
    ]
    assert len(pieces) == 37  # 1 piece short of the displayed 38
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "22", "MANUTAP-OF tem display=22 no nome"
    assert r["des_seq"] == "37", "Ultima peca tem display=37 no nome"
    assert r["ssg_seq"] == "23"
    assert r["ssg_ref"] == "18775/2025"


def test_reiteracao_usa_display_do_NOME_quando_so_a_ultima_pula_um_numero():
    """Caso TC/011446/2022 (22/06/2026): tree mostra "28." como ultima,
    mas a enumeracao retorna 28 itens (0..27). Antes do fix:
        manutap_seq=17 (correto, parse da posicao 17 do array)
        des_seq=27 (errado, n-1=27, deveria ser 28)
    O fix usa o display number do nome ('28. ...') -> 28.
    """
    pieces = []
    for i in range(18):  # 0..17
        if i == 17:
            pieces.append((i, "17. MANUTAP-OF - 999/2025"))
        else:
            pieces.append((i, f"{i}. PEÇA - {i}/2025"))
    # Depois da MANUTAP-OF, a árvore mostra mais 11 peças numeradas 18..28,
    # mas a enumeracao só captura 10 delas (pulou uma). A última piece
    # capturada tem display=28 mas array_pos=27.
    iv = 100
    captured_displays = [18, 19, 20, 21, 22, 23, 24, 25, 26, 28]  # display 27 perdido
    for d in captured_displays:
        pieces.append((iv, f"{d}. PEÇA - {d}/2026"))
        iv += 1
    assert len(pieces) == 28  # uma a menos que as 29 da arvore
    r = main._extract_reiteracao_data_from_pieces(pieces)
    assert r["manutap_seq"] == "17"
    assert r["des_seq"] == "28", "ultima peça tem display=28 no nome (array_pos=27)"


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
