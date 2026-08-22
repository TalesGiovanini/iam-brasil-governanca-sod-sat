from __future__ import annotations

import sys
from pathlib import Path


def project_root() -> Path:
    """Retorna a pasta gravável do projeto ou a pasta ao lado do executável."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resource_root() -> Path:
    """Retorna os recursos embarcados no executável ou a raiz do código-fonte."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return project_root()

