# Comece aqui no Codex

Abra a pasta `tales-agent-starter` no VS Code e dê ao Codex a seguinte orientação inicial:

> Leia integralmente `AGENTS.md`, `README.md`, `docs/ARCHITECTURE.md`, `docs/SECURITY.md`, `docs/ROADMAP.md` e `docs/DECISIONS.md`. Não comece reestruturando o projeto. Primeiro valide o bootstrap e execute os testes existentes. Depois implemente apenas a Fase 0 do Roadmap: configuração real via YAML/env, `tales status` refletindo a configuração e testes. Preserve obrigatoriamente o funcionamento com IA desabilitada. Não implemente OpenAI ainda. Ao terminar, execute os testes e apresente alterações, decisões e riscos pendentes.

## Primeira execução no Windows
Abra PowerShell na raiz do projeto:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/bootstrap.ps1
```

Depois:

```powershell
tales status
pytest
```

## Regra essencial
Não peça ao Codex para "criar um agente de IA" de forma genérica. Trabalhe por fases do `ROADMAP.md`, pois isso reduz lock-in e evita que o projeto vire apenas um wrapper de LLM.
