"""
Utilidades OOXML para Oficios SSG.

Objetivo desta camada (isolada e testavel):
- garantir que os tokens "@@" presentes nos modelos oficiais (Educacao/Saude/Geral
  para UTAP/Dilacao/Reiteracao/Juizo) cheguem intactos ao DOCX gerado, porque o
  proprio e-TCM e quem preenche esses campos automaticamente apos o upload;
- permitir que apenas placeholders NAO iniciados por "@@" sejam substituidos
  com regra de negocio (ex.: {CARGO}, {NOME}, {ORGAO}, {ENDERECO}, {{...}});
- expor primitivas para conversao DOTX->DOCX preservando layout (sem reserializar
  conteudo quando nao for necessario).

Toda funcao publica e auto-contida: nao depende de main.py para nao ter risco
de ciclo. main.py importa daqui, nunca o contrario.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple
from xml.etree import ElementTree as ET


# ------------------------------ Constantes ----------------------------------

# Regex pedida pelo brief: @@ seguido de [A-Za-z, latinos acentuados, 0-9, _].
# Marcamos UNICODE explicitamente para que [À-ÿ] funcione com encoding utf-8.
AT_TOKEN_RE: re.Pattern[str] = re.compile(r"@@[A-Za-zÀ-ÿ0-9_]+", flags=re.UNICODE)

# Arquivos do pacote OOXML onde o texto visivel pode aparecer.
# Limita a busca para evitar varrer recursos binarios e medias.
OOXML_TEXT_PARTS: Tuple[str, ...] = (
    "word/document.xml",
    "word/header1.xml", "word/header2.xml", "word/header3.xml",
    "word/footer1.xml", "word/footer2.xml", "word/footer3.xml",
    "word/footnotes.xml", "word/endnotes.xml", "word/comments.xml",
)

GENERIC_ENCAMINHA_TEXT = "Cópia da(s) peça(s) dos autos."
# A barra correta no marcador final é "/" (nunca "\"). Histórico: a primeira
# versão usou r"\euclides" e gerou ofícios com a barra invertida em PROD;
# o brief exige "/euclides".
DEFAULT_EUCLIDES_MARKER = "/euclides"


# ------------------------------ Excecoes ------------------------------------

class AtTokenViolation(RuntimeError):
    """Erro levantado quando alguma regra de preservacao de @@ e violada.

    Casos cobertos:
      - chave proibida (@@...) passada para safe_replace_non_at_placeholders;
      - @@token presente no modelo desaparece no DOCX gerado;
      - @@token recebe substituicao para valor real antes do upload.
    """


class DocxValidationError(RuntimeError):
    """Erro de validação do DOCX final antes do upload."""


# ------------------------------ Estrutura -----------------------------------

@dataclass
class AtTokenReport:
    """Relatorio retornado por assert_at_tokens_preserved / safe_replace_*."""

    template_path: str
    generated_path: str
    template_tokens: List[str] = field(default_factory=list)
    generated_tokens: List[str] = field(default_factory=list)
    missing_in_generated: List[str] = field(default_factory=list)
    extra_in_generated: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.missing_in_generated

    def as_log_line(self) -> str:
        return (
            f"[@@ validacao] template={Path(self.template_path).name} "
            f"gerado={Path(self.generated_path).name} "
            f"tokens_modelo={len(self.template_tokens)} "
            f"tokens_gerado={len(self.generated_tokens)} "
            f"ausentes={self.missing_in_generated or '[]'} "
            f"extras={self.extra_in_generated or '[]'}"
        )


# ------------------------------ Helpers OOXML -------------------------------

def _read_part_text(zf: zipfile.ZipFile, name: str) -> str:
    try:
        data = zf.read(name)
    except KeyError:
        return ""
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("utf-8", errors="ignore")


def _iter_ooxml_text_payload(path: Path) -> Iterable[str]:
    """Itera o conteudo XML cru das partes textuais do pacote OOXML.

    Trabalhar sobre o XML cru e essencial para detectar tokens mesmo quando o
    Word divide um placeholder em multiplos elementos <w:t>. Por isso fazemos
    a varredura em duas frentes (XML cru + texto visivel apos strip de tags).
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Pacote OOXML nao encontrado: {p}")
    try:
        with zipfile.ZipFile(p, "r") as zf:
            for name in OOXML_TEXT_PARTS:
                xml = _read_part_text(zf, name)
                if xml:
                    yield xml
    except zipfile.BadZipFile as e:
        raise ValueError(f"Arquivo {p.name} nao e um pacote OOXML valido: {e}") from e


_TAG_STRIP_RE = re.compile(r"<[^>]+>")

# Namespace WordprocessingML
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_P = f"{{{_W_NS}}}p"
_W_T = f"{{{_W_NS}}}t"


def _iter_paragraph_texts(xml: str) -> Iterable[str]:
    """Parseia o XML OOXML e produz o texto concatenado por paragrafo.

    Isso e a forma correta de lidar com o "split-run": o Word frequentemente
    separa um token (ex: `@@numero_oficio`) em dois `<w:t>` adjacentes dentro
    do mesmo `<w:p>` por causa de mudancas de formatacao. Concatenar todos os
    `<w:t>` filhos de um mesmo `<w:p>` recompoe o token original SEM correr o
    risco de unir textos de paragrafos diferentes (que produziria tokens
    espurios como `@@Nome_interessadoProc`).
    """
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return
    for p in root.iter(_W_P):
        parts: List[str] = []
        for t in p.iter(_W_T):
            if t.text:
                parts.append(t.text)
        joined = "".join(parts)
        if joined:
            yield joined


def _strip_accents_for_compare(s: str) -> str:
    import unicodedata

    nfkd = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _format_piece_number(value: int | str) -> str:
    raw = str(value).strip()
    m = re.search(r"\d+", raw)
    if not m:
        raise ValueError(f"Número de peça inválido: {value!r}")
    n = int(m.group(0))
    if n <= 0:
        raise ValueError(f"Número de peça inválido: {value!r}")
    return f"{n:02d}"


def _marker_count_in_docx(docx_path: Path | str, marker: str) -> int:
    text = "\n".join(extract_visible_text_from_docx(docx_path))
    return text.count(marker)


# ------------------------------ API publica ---------------------------------

def extract_at_tokens_from_docx(path: Path | str) -> Set[str]:
    """Retorna o conjunto de tokens @@... presentes no DOCX/DOTX.

    Lida corretamente com:
      - tokens intactos (`@@processo` em um unico <w:t>);
      - tokens split-run (`@@` + `processo` em dois <w:t> do MESMO <w:p>);
      - tokens repetidos em multiplas partes do pacote (document/header/footer).

    NAO confunde texto de paragrafos diferentes (cada `<w:p>` e tratado como
    uma unidade de concatenacao), evitando criar tokens espurios como
    `@@Nome_interessadoProc`.

    Tambem aplica uma passada de fallback sobre o XML cru por seguranca, para
    cobrir partes onde a estrutura `<w:p>/<w:t>` esteja ausente ou exotica.
    """
    tokens: Set[str] = set()
    for xml in _iter_ooxml_text_payload(Path(path)):
        # 1) Por paragrafo (lida com split-run corretamente)
        for paragraph_text in _iter_paragraph_texts(xml):
            for m in AT_TOKEN_RE.findall(paragraph_text):
                tokens.add(m)
        # 2) Fallback no XML cru (tokens que estejam fora de <w:p>/<w:t>,
        # por exemplo dentro de comentarios ou notas com estrutura propria).
        for m in AT_TOKEN_RE.findall(xml):
            tokens.add(m)
    return tokens


def extract_visible_text_from_docx(path: Path | str) -> List[str]:
    """Retorna textos visíveis por parágrafo nas partes principais do OOXML."""
    paragraphs: List[str] = []
    for xml in _iter_ooxml_text_payload(Path(path)):
        paragraphs.extend(_iter_paragraph_texts(xml))
    return paragraphs


def format_encaminha_from_piece_numbers(piece_numbers: list[int | str]) -> str:
    """Formata a linha Encaminha com números ordinais das peças do processo.

    Exemplos:
      [3] -> "Cópia da peça 03 dos autos."
      [3, 5] -> "Cópia das peças 03 e 05 dos autos."
      [3, 5, 7] -> "Cópia das peças 03, 05 e 07 dos autos."
    """
    if not piece_numbers:
        raise ValueError("Nenhum número de peça informado para Encaminha.")
    nums: list[str] = []
    for value in piece_numbers:
        formatted = _format_piece_number(value)
        if formatted not in nums:
            nums.append(formatted)
    if len(nums) == 1:
        return f"Cópia da peça {nums[0]} dos autos."
    if len(nums) == 2:
        joined = " e ".join(nums)
    else:
        joined = ", ".join(nums[:-1]) + f" e {nums[-1]}"
    return f"Cópia das peças {joined} dos autos."


def assert_encaminha_has_piece_number(docx_path: Path | str) -> None:
    """Falha se o DOCX mantiver Encaminha genérico ou sem peça numerada."""
    text = "\n".join(extract_visible_text_from_docx(docx_path))
    comparable = _strip_accents_for_compare(text).lower()
    generic = _strip_accents_for_compare(GENERIC_ENCAMINHA_TEXT).lower()
    if generic in comparable:
        raise DocxValidationError(
            f"Encaminha genérico encontrado em {Path(docx_path).name}: {GENERIC_ENCAMINHA_TEXT}"
        )
    if not re.search(r"c[oó]pia\s+d(?:a|as)\s+pe[cç]a(?:s)?\s+\d{2}\b", text, flags=re.I):
        raise DocxValidationError(
            f"Linha Encaminha sem referência a peça numerada em {Path(docx_path).name}."
        )


def set_encaminha_text_without_bold(docx_path: Path | str, encaminha_text: str) -> None:
    """Atualiza o conteúdo de Encaminha e força esse conteúdo sem negrito.

    O rótulo "Encaminha" pode manter a formatação original. O texto depois do
    rótulo, por exemplo "Cópia da peça 03 dos autos.", recebe `bold=False`
    explicitamente para não herdar negrito do rótulo no Word.

    Brief 2026-05-14: o conteúdo (após "Encaminha") DEVE ser Times New Roman
    12pt sem negrito. O template tinha o run com fonte/tamanho diferente do
    corpo do ofício, e a cópia 1:1 de `rPr` preservava esse erro. Agora
    aplicamos explicitamente font name/size APÓS o `_copy_run_format`, para
    sobrescrever só essas duas propriedades sem perder cor/estilo de
    sublinhado/itálico que possam vir do template.
    """
    import os as _os

    encaminha_text = (encaminha_text or "").strip()
    if not encaminha_text:
        raise ValueError("encaminha_text não pode ser vazio")

    try:
        from docx import Document  # type: ignore
        from docx.shared import Pt  # type: ignore
    except Exception as e:  # pragma: no cover - dependencia obrigatoria
        raise RuntimeError(f"python-docx ausente; não foi possível editar Encaminha: {e}") from e

    forced_font_name = (_os.getenv("OFICIO_ENCAMINHA_FONT_NAME") or "Times New Roman").strip() or "Times New Roman"
    try:
        forced_font_size_pt = int((_os.getenv("OFICIO_ENCAMINHA_FONT_SIZE_PT") or "12").strip() or "12")
    except ValueError:
        forced_font_size_pt = 12

    path = Path(docx_path)
    doc = Document(str(path))

    def _copy_run_format(src, dst) -> None:
        """Copia atributos visiveis (fonte/tamanho/estilo/cor) de src para dst."""
        if src is None or dst is None:
            return
        try:
            if src.style is not None:
                dst.style = src.style
        except Exception:
            pass
        for attr in ("name", "size", "italic", "underline", "strike",
                     "subscript", "superscript", "all_caps", "small_caps"):
            try:
                value = getattr(src.font, attr)
                if value is not None:
                    setattr(dst.font, attr, value)
            except Exception:
                pass
        # Cor RGB (pode falhar com tema; tolerante a None)
        try:
            if src.font.color is not None and src.font.color.rgb is not None:
                dst.font.color.rgb = src.font.color.rgb
        except Exception:
            pass
        # Propaga rPr (XML) como ultimo recurso para preservar mais detalhes.
        try:
            from copy import deepcopy
            src_rpr = src._element.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
            )
            if src_rpr is not None:
                existing = dst._element.find(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
                )
                if existing is not None:
                    dst._element.remove(existing)
                dst._element.insert(0, deepcopy(src_rpr))
        except Exception:
            pass

    def _find_content_run_style(paragraph, label_end_in_full: int):
        """Retorna um run que represente o estilo do CONTEUDO (depois de Encaminha)."""
        if not paragraph.runs:
            return None
        # Conta o offset acumulado dos runs e devolve o primeiro run cujo
        # range cobre o pos `label_end_in_full + 1` (ou seja, o primeiro
        # caractere apos o rotulo). Se nao achar, devolve o ultimo run
        # do paragrafo (fallback razoavel — geralmente o conteudo padrao).
        offset = 0
        for run in paragraph.runs:
            run_len = len(run.text or "")
            if offset + run_len > label_end_in_full and (run.text or "").strip():
                # Excluir run que contenha 'Encaminha' (e' o rotulo, nao conteudo).
                if "encaminha" not in _strip_accents_for_compare(run.text or "").lower():
                    return run
            offset += run_len
        # Fallback: ultimo run com texto nao vazio que nao seja o label.
        for run in reversed(paragraph.runs):
            if run.text and "encaminha" not in _strip_accents_for_compare(run.text).lower():
                return run
        return None

    def rewrite_same_paragraph(paragraph) -> bool:
        full_text = "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
        if not full_text:
            return False
        m = re.search(r"\bEncaminha\b", full_text, flags=re.I)
        if not m:
            return False
        # Só mexe em parágrafos que contenham a linha genérica ou texto de peça.
        tail = full_text[m.end():]
        tail_cmp = _strip_accents_for_compare(tail).lower()
        if (
            "copia da(s) peca(s) dos autos" not in tail_cmp
            and not re.search(r"copia\s+d(?:a|as)\s+peca(?:s)?\s+\d{2}", tail_cmp)
            and encaminha_text not in full_text
        ):
            return False

        # Captura o separador (tab/espaços) que originalmente existia entre o
        # rótulo "Encaminha" e o conteúdo. Em PROD o modelo usa um TAB para
        # alinhar a coluna; rstrip() apagava esse \t e substituía por " ",
        # quebrando a tabulação visual.
        sep_match = re.match(r"[ \t\xa0]*", tail)
        original_separator = sep_match.group(0) if sep_match else ""
        if not original_separator:
            # Sem espaços no original (raro): força um TAB para preservar o
            # layout tabulado do template SSG-Aposentadoria.
            original_separator = "\t"
        label_text = full_text[:m.end()]
        label_bold = None
        label_run_ref = None
        if paragraph.runs:
            for run in paragraph.runs:
                if run.text and "encaminha" in _strip_accents_for_compare(run.text).lower():
                    label_bold = run.bold
                    label_run_ref = run
                    break
        # Captura o estilo do CONTEUDO original (antes de zerar os runs).
        content_style_ref = _find_content_run_style(paragraph, m.end())
        if not paragraph.runs:
            paragraph.text = ""
        else:
            for run in paragraph.runs:
                run.text = ""
        label_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        label_run.text = f"{label_text}{original_separator}"
        if label_bold is not None:
            label_run.bold = label_bold
        if label_run_ref is not None:
            _copy_run_format(label_run_ref, label_run)
        content_run = paragraph.add_run(encaminha_text)
        content_run.bold = False
        # IMPORTANTE: copia fonte/tamanho/cor do run de conteudo original
        # para nao cair no default do python-docx (Calibri 11pt). Logo apos,
        # sobrescrevemos font name/size explicitamente para o padrao do brief
        # (Times New Roman 12pt) — o template tinha um run com font/size que
        # nao batia com o corpo do oficio, e a copia 1:1 propagava o erro.
        if content_style_ref is not None:
            _copy_run_format(content_style_ref, content_run)
            content_run.bold = False
        try:
            content_run.font.name = forced_font_name
            content_run.font.size = Pt(forced_font_size_pt)
            # Tambem ajusta os atributos eastAsia/cs no rPr para garantir
            # que o Word renderize Times New Roman mesmo em paragrafos com
            # script complex / CJK; senao alguns motoradores caem no default.
            try:
                from docx.oxml.ns import qn  # type: ignore
                rPr = content_run._element.get_or_add_rPr()
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    from docx.oxml import OxmlElement  # type: ignore
                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                    rFonts.set(qn(f"w:{attr}"), forced_font_name)
            except Exception:
                pass
        except Exception:
            pass
        return True

    def rewrite_content_paragraph(paragraph) -> bool:
        full_text = "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
        cmp = _strip_accents_for_compare(full_text or "").lower()
        if (
            "copia da(s) peca(s) dos autos" not in cmp
            and not re.search(r"copia\s+d(?:a|as)\s+peca(?:s)?\s+\d{2}", cmp)
        ):
            return False
        if not paragraph.runs:
            paragraph.text = encaminha_text
            for run in paragraph.runs:
                run.bold = False
                try:
                    run.font.name = forced_font_name
                    run.font.size = Pt(forced_font_size_pt)
                except Exception:
                    pass
            return True
        # Preserva o estilo do primeiro run com texto util como referencia
        # e edita apenas o texto. Forca TNR 12pt sem negrito conforme brief.
        first_run = paragraph.runs[0]
        first_run.text = encaminha_text
        first_run.bold = False
        try:
            first_run.font.name = forced_font_name
            first_run.font.size = Pt(forced_font_size_pt)
            try:
                from docx.oxml.ns import qn  # type: ignore
                rPr = first_run._element.get_or_add_rPr()
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    from docx.oxml import OxmlElement  # type: ignore
                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                    rFonts.set(qn(f"w:{attr}"), forced_font_name)
            except Exception:
                pass
        except Exception:
            pass
        for run in paragraph.runs[1:]:
            run.text = ""
            run.bold = False
        return True

    def process_paragraphs(paragraphs) -> bool:
        changed = False
        for idx, paragraph in enumerate(paragraphs):
            if rewrite_same_paragraph(paragraph):
                changed = True
                continue
            text_cmp = _strip_accents_for_compare(paragraph.text or "").strip().lower()
            if text_cmp == "encaminha" and idx + 1 < len(paragraphs):
                if rewrite_content_paragraph(paragraphs[idx + 1]):
                    changed = True
        return changed

    def process_tables(tables) -> bool:
        changed = False
        for table in tables or []:
            for row in table.rows:
                for cell in row.cells:
                    changed = process_paragraphs(cell.paragraphs) or changed
                    changed = process_tables(getattr(cell, "tables", []) or []) or changed
        return changed

    changed = process_paragraphs(getattr(doc, "paragraphs", []))
    changed = process_tables(getattr(doc, "tables", []) or []) or changed
    for section in getattr(doc, "sections", []) or []:
        try:
            changed = process_paragraphs(section.header.paragraphs) or changed
            changed = process_tables(getattr(section.header, "tables", []) or []) or changed
            changed = process_paragraphs(section.footer.paragraphs) or changed
            changed = process_tables(getattr(section.footer, "tables", []) or []) or changed
        except Exception:
            continue
    if changed:
        doc.save(str(path))


def assert_encaminha_text_not_bold(docx_path: Path | str) -> None:
    """Falha se o conteúdo numerado de Encaminha estiver em negrito direto."""
    try:
        from docx import Document  # type: ignore
    except Exception as e:  # pragma: no cover
        raise RuntimeError(f"python-docx ausente; não foi possível validar Encaminha: {e}") from e

    doc = Document(str(docx_path))
    found = False

    def check_paragraph(paragraph) -> None:
        nonlocal found
        full = "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
        cmp = _strip_accents_for_compare(full or "").lower()
        if not re.search(r"copia\s+d(?:a|as)\s+peca(?:s)?\s+\d{2}", cmp):
            return
        found = True
        after_label = cmp
        if "encaminha" in after_label:
            after_label = after_label.split("encaminha", 1)[1]
        for run in paragraph.runs:
            run_cmp = _strip_accents_for_compare(run.text or "").lower()
            if not run_cmp.strip():
                continue
            # O rótulo pode ser negrito; o conteúdo da cópia não pode.
            if "copia" in run_cmp or "peca" in run_cmp or re.search(r"\b\d{2}\b", run_cmp):
                if run.bold is True:
                    raise DocxValidationError(
                        f"Conteúdo de Encaminha está em negrito em {Path(docx_path).name}."
                    )

    def walk(paragraphs, tables) -> None:
        for paragraph in paragraphs or []:
            check_paragraph(paragraph)
        for table in tables or []:
            for row in table.rows:
                for cell in row.cells:
                    walk(cell.paragraphs, getattr(cell, "tables", []) or [])

    walk(getattr(doc, "paragraphs", []), getattr(doc, "tables", []) or [])
    for section in getattr(doc, "sections", []) or []:
        try:
            walk(section.header.paragraphs, getattr(section.header, "tables", []) or [])
            walk(section.footer.paragraphs, getattr(section.footer, "tables", []) or [])
        except Exception:
            continue
    if not found:
        raise DocxValidationError(
            f"Conteúdo numerado de Encaminha não encontrado em {Path(docx_path).name}."
        )


def add_euclides_marker_to_docx(docx_path: Path | str, marker: str = DEFAULT_EUCLIDES_MARKER) -> None:
    """Insere o marcador literal no fim do corpo do DOCX, sem duplicar."""
    marker = marker or DEFAULT_EUCLIDES_MARKER
    path = Path(docx_path)
    count = _marker_count_in_docx(path, marker)
    if count == 1:
        return
    if count > 1:
        raise DocxValidationError(
            f"Marcador {marker!r} duplicado em {path.name}: {count} ocorrências."
        )

    try:
        from docx import Document  # type: ignore
    except Exception as e:  # pragma: no cover - dependencia obrigatoria
        raise RuntimeError(f"python-docx ausente; não foi possível inserir {marker!r}: {e}") from e

    doc = Document(str(path))
    paragraphs = list(getattr(doc, "paragraphs", []) or [])
    last_visible = None
    for paragraph in reversed(paragraphs):
        if (paragraph.text or "").strip():
            last_visible = paragraph
            break

    if last_visible is not None and (last_visible.text or "").strip() == "/":
        if last_visible.runs:
            last_visible.runs[0].text = marker
            for run in last_visible.runs[1:]:
                run.text = ""
        else:
            last_visible.text = marker
    else:
        doc.add_paragraph(marker)

    doc.save(str(path))


def assert_euclides_marker_present_once(
    docx_path: Path | str,
    marker: str = DEFAULT_EUCLIDES_MARKER,
) -> None:
    """Falha se o marcador estiver ausente ou duplicado."""
    marker = marker or DEFAULT_EUCLIDES_MARKER
    count = _marker_count_in_docx(docx_path, marker)
    if count != 1:
        raise DocxValidationError(
            f"Marcador {marker!r} deve aparecer exatamente uma vez em "
            f"{Path(docx_path).name}; encontrado={count}."
        )


def assert_at_tokens_preserved(template_path: Path | str,
                                generated_path: Path | str) -> AtTokenReport:
    """Verifica que todo @@token do modelo continua presente no DOCX gerado.

    Regras:
      - se algum @@token do modelo SUMIU no gerado -> AtTokenViolation;
      - tokens "extras" no gerado nao quebram (apenas aparecem no relatorio).
        Isso pode acontecer quando o gerador insere tokens novos por engano,
        e a operacao subsequente em main.py decide o que fazer.
    """
    t_path = Path(template_path)
    g_path = Path(generated_path)
    t_tokens = extract_at_tokens_from_docx(t_path)
    g_tokens = extract_at_tokens_from_docx(g_path)

    report = AtTokenReport(
        template_path=str(t_path),
        generated_path=str(g_path),
        template_tokens=sorted(t_tokens),
        generated_tokens=sorted(g_tokens),
        missing_in_generated=sorted(t_tokens - g_tokens),
        extra_in_generated=sorted(g_tokens - t_tokens),
    )
    if report.missing_in_generated:
        raise AtTokenViolation(
            "Tokens @@ ausentes no DOCX gerado: "
            f"{report.missing_in_generated} (template={t_path.name}, gerado={g_path.name})"
        )
    return report


def safe_replace_non_at_placeholders(
    docx_path: Path | str,
    replacements: Dict[str, str],
    template_path: Optional[Path | str] = None,
) -> AtTokenReport:
    """Substitui APENAS placeholders que NAO comecem por '@@'.

    - Rejeita qualquer chave que comece com '@@' (levanta AtTokenViolation).
    - Aplica substituicoes em paragrafos, tabelas, cabecalhos, rodapes.
    - Mantem o formato do primeiro run da paragrafo apos a substituicao
      (corrige o problema de split-run para placeholders {...} sem destruir
      a formatacao do paragrafo).
    - Quando template_path e fornecido, valida os @@ ao final via
      assert_at_tokens_preserved.

    Retorna o AtTokenReport produzido pela validacao final (ou um relatorio
    sem comparacao quando template_path = None).
    """
    forbidden = sorted(k for k in replacements.keys() if str(k).startswith("@@"))
    if forbidden:
        raise AtTokenViolation(
            f"safe_replace_non_at_placeholders rejeitou chaves @@ proibidas: {forbidden}"
        )

    g_path = Path(docx_path)
    if not g_path.exists():
        raise FileNotFoundError(f"DOCX nao existe: {g_path}")

    if replacements:
        try:
            from docx import Document  # type: ignore
        except Exception as e:  # pragma: no cover - dependencia obrigatoria
            raise RuntimeError(
                f"python-docx ausente; instale 'python-docx' para usar safe_replace_non_at_placeholders ({e})"
            ) from e

        doc = Document(str(g_path))
        _apply_paragraph_replacements_doc(doc, replacements)
        doc.save(str(g_path))

    if template_path is not None:
        return assert_at_tokens_preserved(template_path, g_path)

    # Sem template: gera relatorio simples (so com tokens do gerado).
    g_tokens = extract_at_tokens_from_docx(g_path)
    return AtTokenReport(
        template_path="(nao fornecido)",
        generated_path=str(g_path),
        template_tokens=[],
        generated_tokens=sorted(g_tokens),
        missing_in_generated=[],
        extra_in_generated=[],
    )


def convert_dotx_to_docx_preserving_layout(template_path: Path | str,
                                            out_path: Path | str) -> Path:
    """Converte um .dotx em .docx SEM reserializar o conteudo.

    Estrategia: copia o pacote OOXML byte-a-byte trocando apenas o
    content-type principal de template para document. Isso preserva
    100% do layout (estilos, runs, splits, tabelas) — ao contrario de
    abrir e re-salvar via python-docx, que pode alterar a estrutura.
    """
    src = Path(template_path)
    dst = Path(out_path)
    if not src.exists():
        raise FileNotFoundError(f"Modelo nao encontrado: {src}")
    if src.suffix.lower() == ".docx":
        # Ja e DOCX: apenas copia
        shutil.copyfile(src, dst)
        return dst

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                try:
                    xml = data.decode("utf-8")
                except UnicodeDecodeError:
                    xml = data.decode("utf-8", errors="ignore")
                xml = xml.replace(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                )
                data = xml.encode("utf-8")
            zout.writestr(info, data)
    return dst


# ------------------------- python-docx helpers ------------------------------

def _apply_paragraph_replacements_doc(doc, replacements: Dict[str, str]) -> None:
    """Aplica replacements em um objeto Document do python-docx.

    Corrige split-run: concatena todos os runs do paragrafo, substitui no
    texto inteiro, e re-injeta no primeiro run, esvaziando os demais. Isso
    preserva o estilo do primeiro run (que geralmente e o run "padrao" do
    placeholder no modelo do Word) sem destruir a paragrafo inteiro.
    """

    def replace_in_paragraph(para) -> None:
        if not para.runs:
            # Paragrafo sem runs: edita o texto direto.
            txt = para.text or ""
            new = txt
            for k, v in replacements.items():
                if k in new:
                    new = new.replace(k, str(v))
            if new != txt:
                para.text = new
            return

        original = "".join(r.text for r in para.runs)
        new_text = original
        for k, v in replacements.items():
            if k in new_text:
                new_text = new_text.replace(k, str(v))
        if new_text == original:
            return
        # Mantem o primeiro run com o texto completo e esvazia os demais.
        para.runs[0].text = new_text
        for r in para.runs[1:]:
            r.text = ""

    def walk_paragraphs(paragraphs) -> None:
        for p in paragraphs:
            replace_in_paragraph(p)

    def walk_tables(tables) -> None:
        for table in tables or []:
            for row in table.rows:
                for cell in row.cells:
                    walk_paragraphs(cell.paragraphs)
                    walk_tables(getattr(cell, "tables", None) or [])

    walk_paragraphs(getattr(doc, "paragraphs", []))
    walk_tables(getattr(doc, "tables", []) or [])

    for section in getattr(doc, "sections", []) or []:
        try:
            walk_paragraphs(section.header.paragraphs)
            walk_tables(getattr(section.header, "tables", []) or [])
            walk_paragraphs(section.footer.paragraphs)
            walk_tables(getattr(section.footer, "tables", []) or [])
        except Exception:
            # Cabecalho/rodape ausente — sem problema.
            continue


# ----------------------------- Sanity helpers -------------------------------

def filter_out_at_tokens(replacements: Dict[str, str]) -> Dict[str, str]:
    """Retorna um novo dict sem chaves @@... (uso defensivo).

    Util quando o codigo legado monta um mapping misto e queremos passar
    apenas as chaves nao-@@ adiante (ex.: para safe_replace_non_at_placeholders).
    """
    return {k: v for k, v in replacements.items() if not str(k).startswith("@@")}


def is_at_token_key(key: str) -> bool:
    """Verdadeiro quando a chave inicia exatamente por '@@'."""
    return bool(key) and str(key).startswith("@@")
