# Como contribuir

Obrigado por contribuir com a evolução do IAM | Governança SoD & SAT.

## Princípios do projeto

- Não use dados reais de usuários, perfis, sistemas ou clientes em issues, commits ou anexos.
- Não altere templates ou bases originais; crie sempre cópias de teste sintéticas.
- Não introduza regras SoD, critérios de SAT ou classificações sem uma fonte explícita e revisável.
- Preserve a rastreabilidade: toda decisão de bloqueio, alerta ou geração deve ser explicável.

## Antes de enviar um pull request

1. Abra uma issue para mudanças relevantes e descreva o problema, o comportamento esperado e dados sintéticos mínimos para reproduzi-lo.
2. Mantenha o escopo pequeno e inclua testes para regras de validação ou geração afetadas.
3. Execute a suíte local:

```powershell
py -m unittest discover -s tests -v
```

4. Informe no pull request quais arquivos foram alterados, qual validação foi executada e o que ainda precisa de revisão humana.

## Convenções

- Python com tipagem onde ela esclarece contratos.
- Mensagens e documentação em português claro.
- Funções de negócio devem preferir comportamento conservador: quando não houver evidência suficiente, registrar pendência em vez de inventar um resultado.

## Reporte de problema

Inclua a versão, o sistema operacional, os passos com dados sintéticos, o resultado atual e o resultado esperado. Nunca anexe planilhas de acesso reais.

