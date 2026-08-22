"""Ponto de entrada da aplicação local IAM Matriz SoD/SAT."""
import sys


if __name__ == "__main__":
    # O executável gráfico não carrega pandas/openpyxl antes de o usuário pedir
    # uma análise. Isso reduz sensivelmente o tempo percebido de abertura.
    if getattr(sys, "frozen", False) and len(sys.argv) == 1:
        from src.gui import start

        start()
    else:
        from src.cli import main

        raise SystemExit(main())

