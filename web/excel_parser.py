"""Lê uma planilha enviada pelo Gilson e extrai a lista de processos TC/xxxxx/aaaa.

Regras:
- Aceita .xlsx (openpyxl) e .xls (xlrd 1.2).
- Faz varredura em todas as abas e todas as células.
- Captura padrões TC/000000/0000 ou TC 000000/0000 (com variações de barra/hífen).
- Normaliza para o formato canônico "TC/000000/0000".
- Preserva ordem de aparição e remove duplicatas.
"""
from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Iterable

import openpyxl

TC_PATTERN = re.compile(
    r"TC[\s/.\-]*0*(\d{1,7})[\s/.\-]+(\d{4})",
    re.IGNORECASE,
)


def _normalize(num: str, ano: str) -> str:
    num_pad = num.zfill(6)
    return f"TC/{num_pad}/{ano}"


def _iter_cells_xlsx(data: bytes) -> Iterable[str]:
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue
                    yield str(cell)
    finally:
        wb.close()


def _iter_cells_xls(data: bytes) -> Iterable[str]:
    import xlrd

    book = xlrd.open_workbook(file_contents=data)
    for sheet in book.sheets():
        for ri in range(sheet.nrows):
            for ci in range(sheet.ncols):
                v = sheet.cell_value(ri, ci)
                if v is None or v == "":
                    continue
                yield str(v)


def parse_processos(filename: str, data: bytes) -> list[str]:
    """Retorna a lista de processos detectados, na ordem de aparição, sem duplicatas."""
    name = (filename or "").lower()
    if name.endswith(".xls"):
        cells = _iter_cells_xls(data)
    else:
        cells = _iter_cells_xlsx(data)

    seen: set[str] = set()
    ordered: list[str] = []
    for text in cells:
        for m in TC_PATTERN.finditer(text):
            num, ano = m.group(1), m.group(2)
            canonical = _normalize(num, ano)
            if canonical in seen:
                continue
            seen.add(canonical)
            ordered.append(canonical)
    return ordered


def parse_processos_from_path(path: str | Path) -> list[str]:
    p = Path(path)
    return parse_processos(p.name, p.read_bytes())
