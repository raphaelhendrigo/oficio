"""Treina o classificador automatico Euclides.

Uso: .venv/Scripts/python.exe scripts/train_classifier.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from web import classifier  # noqa: E402

if __name__ == "__main__":
    meta = classifier.train_and_save(verbose=True)
    print()
    print(f"meta salvo em {classifier.META_PATH}")
