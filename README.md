# IAM | Governança SoD & SAT + Tales Agent Platform

> Repositório unificado: análise de Segregação de Funções (SoD) e Sensitive Access Transactions (SAT), combinado com a plataforma agentic `tales-agent` — local-first, CLI-first e independente de fornecedor externo.

---

## IAM | Governança SoD & SAT

Aplicação Python local voltada à análise de bases de funcionalidades/transações e usuários/perfis. Organiza o resultado no template Excel definido pelo time responsável, registra a conciliação por sistema e produz diagnósticos acionáveis.

### O que a aplicação faz

- Identifica sistemas, perfis e funcionalidades em planilhas heterogêneas;
- Concilia as bases nos dois sentidos, por sistema e perfil;
- Bloqueia a geração do resultado quando não há sistema conciliado ou há divergências críticas;
- Produz Matriz Funcional, Atividades Conflitantes e SAT no modelo Excel configurado;
- Gera diagnóstico, pendências, análise SoD e minuta de e-mail somente quando aplicável;
- Preserva as fontes: a execução sempre acontece em uma nova pasta de saída.

### Uso local

```powershell
py -m pip install -r requirements.txt
py main.py interface
```

Também é possível executar por linha de comando:

```powershell
py main.py processar --funcionalidades .\funcionalidades.xlsx --usuarios .\usuarios.xlsx
```

---

## Tales Agent Platform

Plataforma agentic local-first, CLI-first e independente de fornecedor externo.

### Objetivo

Construir um sistema que combine:
- Core determinístico
- Agent Engine
- Knowledge Base local
- Memory local
- Tools controladas por política
- Módulo de inferência opcional e intercambiável
- CLI como primeira interface

A aplicação deve continuar funcionando sem provedor externo de processamento.

### Começo rápido

Windows PowerShell:
```powershell
./scripts/bootstrap.ps1
```

Linux/macOS:
```bash
chmod +x scripts/bootstrap.sh
./scripts/bootstrap.sh
```

Depois:
```bash
tales status
```

### Regra arquitetural principal
```text
Core != Provedor Externo
Agent != Provedor Externo
Inferência != Dependência Obrigatória
```

---

## Segurança e limites

O mecanismo é local e determinístico: não substitui a validação técnica e a aprovação dos responsáveis. Consulte [SECURITY.md](SECURITY.md) para reportar uma vulnerabilidade.

Nunca envie dados pessoais, bases de acesso, templates preenchidos, evidências de cliente ou resultados de análise.

## Licença

Disponibilizado sob a licença MIT. Consulte [LICENSE](LICENSE).

Leia `AGENTS.md` antes de usar Codex neste repositório.
