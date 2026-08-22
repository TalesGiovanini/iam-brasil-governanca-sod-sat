from __future__ import annotations

import hashlib
from datetime import datetime, timezone
from pathlib import Path

from .configuration import load_config
from .processor import _read_excel, _resolve_columns


def _hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _describe(path: Path, source_config: dict) -> dict:
    if not path.is_file():
        return {"arquivo": str(path), "erro": "Arquivo não encontrado."}
    try:
        sheet, frame = _read_excel(path, source_config.get("aba"))
        mapping, _ = _resolve_columns(frame, source_config, path.name)
        empty = {str(column): int(frame[column].isna().sum()) for column in frame.columns if frame[column].isna().any()}
        return {
            "arquivo": str(path), "sha256": _hash(path), "abas": [{
                "aba": sheet, "linhas": int(len(frame)), "colunas": [str(column) for column in frame.columns],
                "tipos_aparentes": {str(column): str(frame[column].dtype) for column in frame.columns},
                "vazios": empty, "duplicidades_exatas": int(frame.duplicated().sum()), "mapeamento_proposto": mapping,
            }]
        }
    except Exception as exc:
        return {"arquivo": str(path), "erro": str(exc)}


def discover_workbooks(funcionalidades_path: Path, usuarios_path: Path, conflitos_path: Path | None, config_path: Path) -> dict:
    config = load_config(config_path)
    result = {
        "gerado_em_utc": datetime.now(timezone.utc).isoformat(),
        "base_funcionalidades": _describe(funcionalidades_path, config["base_funcionalidades"]),
        "base_usuarios": _describe(usuarios_path, config["base_usuarios"]),
        "proposta_de_mapeamento": {
            "base_funcionalidades": config["base_funcionalidades"].get("colunas", {}),
            "base_usuarios": config["base_usuarios"].get("colunas", {}),
        },
        "pendencias": [
            "O mapeamento é uma proposta técnica até ser validado pelo responsável da fonte.",
            "A reconciliação usuários ↔ funcionalidades será realizada nos dois sentidos antes da Matriz Funcional.",
            "Conflitos e SAT não são inferidos sem fonte ou critério explicitamente aprovado.",
        ],
    }
    if conflitos_path:
        result["atividades_conflitantes"] = _describe(conflitos_path, config["atividades_conflitantes"])
        result["proposta_de_mapeamento"]["atividades_conflitantes"] = config["atividades_conflitantes"].get("colunas", {})
    else:
        result["atividades_conflitantes"] = {"arquivo": None, "pendencia": "Fonte não selecionada."}
    return result


def write_discovery_report(report: dict, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    lines = ["# Descoberta das planilhas", "", f"Gerado em UTC: {report['gerado_em_utc']}", ""]
    for title, key in (("Base de funcionalidades", "base_funcionalidades"), ("Base de usuários", "base_usuarios"), ("Regras de conflitos", "atividades_conflitantes")):
        source = report[key]
        lines.extend([f"## {title}", "", f"Arquivo: `{source.get('arquivo')}`"])
        if source.get("erro"):
            lines.extend([f"Pendência: {source['erro']}", ""])
            continue
        if source.get("pendencia"):
            lines.extend([f"Pendência: {source['pendencia']}", ""])
            continue
        lines.append(f"SHA-256: `{source['sha256']}`")
        for sheet in source["abas"]:
            lines.extend(["", f"### Aba: {sheet['aba']}", f"Linhas preenchidas: {sheet['linhas']}", f"Duplicidades exatas: {sheet['duplicidades_exatas']}", "", "Colunas identificadas:"])
            lines.extend(f"- `{column}`" for column in sheet["colunas"])
            lines.extend(["", "Mapeamento proposto:"])
            lines.extend(f"- `{canonical}` ← `{column}`" for canonical, column in sheet["mapeamento_proposto"].items())
            if sheet["vazios"]:
                lines.extend(["", "Campos com valores vazios:"])
                lines.extend(f"- `{column}`: {count}" for column, count in sheet["vazios"].items())
    lines.extend(["", "## Pendências"])
    lines.extend(f"- {item}" for item in report["pendencias"])
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")

