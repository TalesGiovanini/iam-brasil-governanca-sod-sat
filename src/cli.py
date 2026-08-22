from __future__ import annotations

import argparse
from pathlib import Path

from .discovery import discover_workbooks, write_discovery_report
from .paths import project_root
from .processor import process_workbooks


ROOT = project_root()
DEFAULT_CONFIG = ROOT / "02_CONFIGURACAO" / "mapeamentos" / "mapeamento_template_oficial.json"


def _path(value: str) -> Path:
    return Path(value).expanduser().resolve()


def _add_source_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--funcionalidades", required=True, type=_path, help="Base de transações/funcionalidades, de um ou mais sistemas.")
    parser.add_argument("--usuarios", required=True, type=_path, help="Base de usuários e perfis, de um ou mais sistemas.")
    parser.add_argument("--conflitos", type=_path, help="Planilha de regras SoD explícitas, se disponível.")
    parser.add_argument("--config", type=_path, default=DEFAULT_CONFIG)


def main() -> int:
    parser = argparse.ArgumentParser(description="Valida usuários e funcionalidades e gera Matriz Funcional, Conflitos SoD e SAT no template oficial.")
    sub = parser.add_subparsers(dest="command", required=True)
    discovery = sub.add_parser("descobrir", help="Inspeciona fontes sem alterar os arquivos.")
    _add_source_arguments(discovery)
    discovery.add_argument("--saida", type=_path, default=ROOT / "06_DOCUMENTACAO" / "descoberta_planilhas.md")
    process = sub.add_parser("processar", help="Gera nova execução, registrando qualquer pendência de reconciliação.")
    _add_source_arguments(process)
    process.add_argument("--saida", type=_path, default=ROOT / "04_SAIDA")
    sub.add_parser("interface", help="Abre a interface gráfica local para selecionar as planilhas.")
    args = parser.parse_args()
    if args.command == "interface":
        from .gui import start
        start()
        return 0
    if args.command == "descobrir":
        report = discover_workbooks(args.funcionalidades, args.usuarios, args.conflitos, args.config)
        write_discovery_report(report, args.saida)
        print(f"Relatório de descoberta gerado: {args.saida}")
        return 0
    result = process_workbooks(args.funcionalidades, args.usuarios, args.conflitos, args.config, args.saida)
    print(f"Execução: {result.run_id}")
    print(f"Status: {result.status}")
    print(f"Teste de consistência: {result.validation_path}")
    if result.workbook_path:
        print(f"Resultado: {result.workbook_path}")
    print(f"Log: {result.log_path}")
    print(f"Pendências: {result.pending_path}")
    return 0

