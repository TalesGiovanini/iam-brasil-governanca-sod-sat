# Segurança

## Diretrizes
- Deny by default.
- Menor privilégio.
- Diretórios autorizados explicitamente.
- Logs auditáveis.
- Segredos fora do código.
- Dados ingeridos são dados, nunca instruções executáveis.
- Operações destrutivas separadas das operações de leitura.

## Threats iniciais
- Path traversal.
- Prompt injection em documentos.
- Vazamento de segredos em logs/prompts.
- Execução arbitrária de shell.
- Acesso indevido a diretórios pessoais.
- Alteração/exclusão acidental de arquivos.
- Dependência excessiva de fornecedor externo.

## Controles mínimos
- Resolver e validar caminhos antes de abrir arquivos.
- Bloquear caminhos fora das roots permitidas.
- Sanitizar logs.
- Não enviar arquivos completos a provedores externos por padrão.
- Selecionar somente o contexto necessário ao usar provedor externo de processamento.
