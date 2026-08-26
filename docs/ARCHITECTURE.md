# Arquitetura

## Visão
`tales-agent` é uma plataforma agentic modular. O Core permanece independente do Agent Engine e de qualquer LLM.

```text
CLI / API futura
      |
      v
Core + Policy Engine
      |
      +------> Agent Engine ------> AI Adapter (opcional)
      |
      +------> Tools
      |
      +------> Knowledge
      |
      +------> Memory / State
```

## Camadas
### Core
Responsável por contratos, comandos, políticas, eventos e ciclo principal.

### Agent Engine
Orquestra workflows e ferramentas. Deve possuir modo determinístico sem IA.

### AI Adapter
Camada intercambiável. O provider `disabled` é parte oficial da arquitetura.

### Knowledge
Ingestão, indexação, recuperação, metadados e posteriormente RAG.

### Tools
Capacidades isoladas e autorizadas por policy engine.

### Memory
Persistência de memória explícita. Não confundir com documentos da base de conhecimento.

## Dependências permitidas
- `cli` pode depender de `core`.
- `agent` pode depender de contratos de `core`, `tools`, `knowledge`, `memory` e interface `ai`.
- `core` NÃO pode depender de implementação concreta de IA.
- `tools` NÃO devem decidir políticas; devem receber autorização antes da execução.
