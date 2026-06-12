# excel-network-report

Add-in Excel (Office Scripts) para automação do relatório operacional de disponibilidade de rede móvel.

## Contexto

No monitoramento de redes móveis, a equipe de operações realiza extrações periódicas de uma base de sites com dezenas de colunas e centenas de registros. Antes deste add-in, o processo de formatação, classificação e consolidação desses dados para o relatório gerencial era feito **manualmente**, consumindo tempo e sujeito a erros.

Este add-in automatiza todo esse fluxo diretamente no Excel, com dois botões: **Formatar** e **Atualizar**.

---

## O que o script faz

### Formatar
- Copia a aba `Base` e renomeia automaticamente com horário e índice (`Online_22h_1`, por exemplo)
- Reestrutura as colunas conforme o layout esperado do relatório
- Classifica cada site nas seguintes categorias de **Pareto**:
  - `Field` — falha de campo
  - `PROVEDOR` — falha de provedor (Tx Terceira, Tx SWAP, Ran-sharing)
  - `REGIONAL` — falha regional
  - `2G` — sites de tecnologia 2G
- Categoriza o **downtime** de cada site em faixas horárias:
  - `00h – 06h`
  - `06h – 12h`
  - `12h – 24h`
  - `+24h`
- Adiciona colunas calculadas: `Pareto`, `Downtime`, `2G`, `IMP`, `Em Análise`, `Total`
- Aplica formatação visual: cabeçalhos em negrito, cores por grupo de colunas, freeze da primeira linha
- Remove linhas com flag zero (sites inativos/irrelevantes)

### Atualizar
- Lê todas as abas `Online_*` existentes na planilha
- Consolida os totais por categoria e faixa de downtime nas abas **Resumo** e **Resumo_Field**
- No `Resumo_Field`, detalha os dados de Field por estado: **AM, AP, MA, PA, RR**
- Ordena as abas numericamente antes de processar

---

## Tecnologias

- TypeScript (Office Scripts)
- Excel JavaScript API (`Excel.run`, `context.sync`)
- jQuery (para eventos de botão no painel de tarefas)

---

## Como usar

1. Abra o Excel com a planilha contendo a aba `Base` com os dados exportados do sistema de monitoramento
2. Acesse **Automatizar → Novo Script** ou importe o arquivo via Office Scripts
3. No painel de tarefas, clique em:
   - **Formatar** — para processar e estruturar a extração atual
   - **Atualizar** — para consolidar todas as extrações do dia nas abas de resumo

> Desenvolvido e utilizado como add-in via **Script Lab**.

---

## Estrutura esperada da planilha

| Aba | Descrição |
|-----|-----------|
| `Base` | Dados brutos exportados do sistema de monitoramento |
| `Online_*` | Abas geradas automaticamente pelo botão Formatar |
| `Resumo` | Consolidação geral por categoria e downtime |
| `Resumo_Field` | Consolidação de falhas Field por estado |

---

## Motivação

Desenvolvido para otimizar o processo de reporte operacional em ambiente de engenharia de redes móveis. A automação eliminou etapas manuais repetitivas e padronizou o formato do relatório entregue periodicamente para gestores e equipe técnica.
