import argparse
from tales.core.engine import get_status


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="tales", description="Tales Agent CLI")
    sub = parser.add_subparsers(dest="command")
    sub.add_parser("status", help="Show local system status")
    return parser


def main() -> None:
    args = build_parser().parse_args()

    if args.command == "status":
        status = get_status()
        print("Tales Agent")
        print(f"Core:           {status.core}")
        print(f"Agent Engine:   {status.agent_engine}")
        print(f"Knowledge Base: {status.knowledge_base}")
        print(f"Inferencia:     {'ATIVA' if status.inference_enabled else 'INATIVA'}")
        print(f"Provedor:       {status.inference_provider}")
        return

    build_parser().print_help()


if __name__ == "__main__":
    main()
