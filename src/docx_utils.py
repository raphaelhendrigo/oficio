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


# ------------------------------ Excecoes ------------------------------------

class AtTokenViolation(RuntimeError):
    """Erro levantado quando alguma regra de preservacao de @@ e violada.

    Casos cobertos:
      - chave proibida (@@...) passada para safe_replace_non_at_placeholders;
      - @@token presente no modelo desaparece no DOCX gerado;
      - @@token recebe substituicao para valor real antes do upload.
    """


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
