from __future__ import annotations

import json
from pathlib import Path


class ConfigurationError(ValueError):
    """Indica configuração ausente ou incompatível com a fonte."""


def load_config(path: Path) -> dict:
    if not path.is_file():
        raise ConfigurationError(f"Arquivo de configuração não encontrado: {path}")
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigurationError(f"JSON de configuração inválido: {path}: {exc}") from exc
    for field in ("base_funcionalidades", "base_usuarios", "atividades_conflitantes", "template_oficial"):
        if field not in config:
            raise ConfigurationError(f"Configuração obrigatória ausente: {field}")
    return config


def resolve_from_config(config_path: Path, value: str) -> Path:
    return (config_path.parent / value).resolve()

