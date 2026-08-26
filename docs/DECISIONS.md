# Architecture Decision Log

## ADR-001 — Local-first
Decisão: armazenamento local é o padrão.
Motivo: independência, controle de dados e portabilidade.

## ADR-002 — IA opcional
Decisão: Core e workflows determinísticos não dependem de LLM.
Motivo: permitir uso offline/independente e evitar lock-in.

## ADR-003 — CLI-first
Decisão: funcionalidades entram primeiro pela CLI.
Motivo: reduzir complexidade e manter automação/scriptabilidade.

## ADR-004 — Deny by default
Decisão: acesso a recursos é negado sem política explícita.
Motivo: menor privilégio e segurança operacional.
