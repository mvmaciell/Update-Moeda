O *Report* ZFI_COTACAO_MOEDAS é um programa desenvolvido para automatizar o processo de atualização das taxas de câmbio na tabela de moeda estrangeira dentro do sistema SAP. Este report é utilizado para garantir que as cotações das principais moedas estrangeiras estejam sempre atualizadas com base nos dados fornecidos pelo Banco Central do Brasil (BCB). Abaixo estão os principais detalhes e funcionalidades do programa:

# Funcionalidades

1. **Atualização Automática das Taxas de Câmbio**: 
   - O programa obtém diariamente as cotações de moedas como USD, EUR, e GBP, utilizando a interface do Banco Central do Brasil, e atualiza a tabela de moedas no sistema SAP. A atualização é feita para categorias específicas e pode incluir tanto taxas de compra quanto de venda.

2. **Parâmetros de Seleção**:
   - **Período de Atualização**: Permite especificar a data inicial e final para o período de atualização das cotações.
   - **Categoria M**: Uma opção para categorizar as taxas atualizadas.
   - **Envio de E-mails**: Possibilidade de enviar e-mails para destinatários específicos com o relatório das taxas de câmbio.

3. **Tratamento de Dias Úteis e Feriados**:
   - O programa verifica se o dia é útil e, caso contrário, utiliza as cotações do último dia útil para realizar as atualizações. Isso garante que, mesmo em feriados ou fins de semana, o sistema tenha as cotações mais recentes possíveis.

4. **Envio de Relatórios por E-mail**:
   - O report possui uma funcionalidade para enviar um resumo das cotações por e-mail, incluindo detalhes como as taxas de câmbio para diferentes moedas. O conteúdo do e-mail é formatado em HTML e pode ser enviado para uma lista de destinatários configurada nos parâmetros de seleção.

### Histórico de Modificações

- **26/06/2024**: Desenvolvimento inicial do report.
- **12/07/2024**: Ajustes na funcionalidade de envio de e-mails e comunicação HTTP.
- **30/07/2024**: Inclusão de novas cotações e ajustes adicionais.

### Utilização e Benefícios

Este programa garante que as transações envolvendo moedas estrangeiras sejam realizadas com base em cotações atualizadas e precisas. Além disso, a automação do processo reduz significativamente o tempo e o esforço necessários para a manutenção manual das tabelas de moedas, minimizando erros e aumentando a eficiência operacional.
