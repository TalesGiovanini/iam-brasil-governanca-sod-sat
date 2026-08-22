# IAM Brasil | Governança SoD & SAT

> Matriz funcional, conflitos de Segregação de Funções (SoD) e classificação de Sensitive Access Transactions (SAT), com conciliação rastreável entre bases.

Este projeto é uma aplicação Python local voltada à análise de bases de funcionalidades/transações e usuários/perfis. Ele organiza o resultado no template Excel definido pelo time responsável, registra a conciliação por sistema e produz diagnósticos acionáveis quando as fontes não estiverem aderentes.

## O que a aplicação faz

- identifica sistemas, perfis e funcionalidades em planilhas heterogêneas;
- concilia as bases nos dois sentidos, por sistema e perfil;
- bloqueia a geração do resultado quando não há sistema conciliado ou há divergências críticas;
- produz Matriz Funcional, Atividades Conflitantes e SAT no modelo Excel configurado;
- gera diagnóstico, pendências, análise SoD e minuta de e-mail somente quando aplicável;
- preserva as fontes: a execução sempre acontece em uma nova pasta de saída.

## Regras de resultado

| Situação encontrada | Resultado Excel | Diagnóstico / análise | Minuta de e-mail |
| --- | --- | --- | --- |
| Não há sistema conciliado entre as bases | Bloqueado | Não liberados | Não liberada |
| Sistema conciliado, mas há divergência de perfis | Bloqueado | Liberados | Liberada |
| Bases conciliadas | Liberado | Liberados | Desabilitada |

## Uso local

```powershell
py -m pip install -r requirements.txt
py main.py interface
```

Na interface, selecione a base de funcionalidades e a base de usuários. A terceira fonte é opcional e deve conter regras SoD explícitas, quando existirem.

Também é possível executar por linha de comando:

```powershell
py main.py processar --funcionalidades .\funcionalidades.xlsx --usuarios .\usuarios.xlsx
```

## Configuração e template

Os templates oficiais, bases de clientes, diagnósticos, resultados e executáveis não são versionados neste repositório público. Antes de executar, mantenha seu template em `02_CONFIGURACAO/templates/` e informe o mapeamento de colunas necessário em `02_CONFIGURACAO/mapeamentos/`.

O arquivo `mapeamento_colunas.example.json` é um ponto de partida intencionalmente vazio: os campos devem ser configurados com base nas colunas reais das fontes autorizadas, sem inferir conteúdo de negócio.

## Contribuições

Contribuições são bem-vindas para melhorar validações, normalização, testes, acessibilidade e documentação. Leia [CONTRIBUTING.md](CONTRIBUTING.md) antes de abrir uma issue ou pull request.

Por segurança, nunca envie dados pessoais, bases de acesso, templates preenchidos, evidências de cliente, resultados de análise ou arquivos executáveis.

## Segurança e limites

O mecanismo é local e determinístico: ele não substitui a validação técnica e a aprovação dos responsáveis pelo sistema. Regras, classificações e exceções devem estar documentadas ou ser tratadas como pendência. Consulte [SECURITY.md](SECURITY.md) para reportar uma vulnerabilidade.

## Licença

Este projeto é disponibilizado sob a licença MIT. Consulte [LICENSE](LICENSE).

