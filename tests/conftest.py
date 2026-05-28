"""Fixtures comuns para os testes do projeto oficio_automation.

Garante que `src/` esteja no sys.path para que os testes possam importar
modulos do projeto diretamente (ex.: `import docx_utils`), independente
do diretorio de invocacao do pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


@pytest.fixture(autouse=True)
def _clear_operator_overrides(monkeypatch: pytest.MonkeyPatch) -> None:
    """Garante que overrides de operador (definidos no ambiente para execucoes
    de producao) NAO vazem para os testes, que verificam a logica padrao.

    Ex.: FORCE_TIPO=DILACAO setado num shell faria o classificador retornar
    sempre DILACAO e quebraria os testes de classificacao/selecao de modelo.
    """
    for var in ("FORCE_TIPO", "OFICIO_REFERENCIA_TEXT", "OFICIO_SSG_REF", "OFICIO_ENCAMINHA_TEXT"):
        monkeypatch.delenv(var, raising=False)


@pytest.fixture(scope="session")
def project_root() -> Path:
    return ROOT


@pytest.fixture(scope="session")
def modelo_utap_educacao(project_root: Path) -> Path:
    p = project_root / "modelos_utap" / "SSG - Aposentadoria Providências - Educação.docx"
    if not p.exists():
        pytest.skip(f"Modelo nao encontrado: {p}")
    return p


@pytest.fixture(scope="session")
def modelo_utap_geral(project_root: Path) -> Path:
    p = project_root / "modelos_utap" / "SSG - Aposentadoria Providências - Geral.docx"
    if not p.exists():
        pytest.skip(f"Modelo nao encontrado: {p}")
    return p


@pytest.fixture(scope="session")
def modelo_utap_saude(project_root: Path) -> Path:
    p = project_root / "modelos_utap" / "SSG - Aposentadoria Providências - Saúde.docx"
    if not p.exists():
        pytest.skip(f"Modelo nao encontrado: {p}")
    return p


@pytest.fixture(scope="session")
def modelo_dilacao_geral(project_root: Path) -> Path:
    p = project_root / "modelos_dilacao" / "SSG - Aposentadoria Dilação - Geral.dotx"
    if not p.exists():
        pytest.skip(f"Modelo nao encontrado: {p}")
    return p


@pytest.fixture(scope="session")
def modelo_reiteracao_geral(project_root: Path) -> Path:
    p = project_root / "modelos_reiteracao" / "SSG - Aposentadoria Reiteração - Geral.dotx"
    if not p.exists():
        pytest.skip(f"Modelo nao encontrado: {p}")
    return p


@pytest.fixture(scope="session")
def modelo_juizo_geral(project_root: Path) -> Path:
    p = project_root / "modelos_juizo" / "Modelo Único - APO-PEN julgada retorna SEI - Geral.dotx"
    if not p.exists():
        pytest.skip(f"Modelo nao encontrado: {p}")
    return p
