# Roadmap

## Fase 0 — Bootstrap
- [ ] Instalação local
- [ ] CLI `tales status`
- [ ] carregamento YAML
- [ ] testes básicos

## Fase 1 — Core seguro
- [ ] Policy Engine
- [ ] filesystem read-only
- [ ] auditoria local
- [ ] SQLite

## Fase 2 — Conhecimento local
- [ ] ingestão TXT/MD/JSON/CSV
- [ ] metadados
- [ ] hash e deduplicação
- [ ] SQLite FTS5
- [ ] `tales index`
- [ ] `tales search`

## Fase 3 — Agente independente
- [ ] executor
- [ ] workflows determinísticos
- [ ] planner baseado em regras
- [ ] `tales agent run`

## Fase 4 — Módulo de Inferência plugável
- [ ] interface InferenceProvider
- [ ] NullProvider (operação sem provedor)
- [ ] provedor externo opcional via configuração
- [ ] provedor local opcional
- [ ] inferência habilitada/desabilitada via config

## Fase 5 — RAG
- [ ] chunking
- [ ] embeddings opcionais
- [ ] retrieval híbrido
- [ ] citações de origem

## Fase 6 — Expansão
- [ ] DOCX/PDF robustos
- [ ] API local
- [ ] UI web/desktop
- [ ] acesso remoto seguro
- [ ] integrações externas
