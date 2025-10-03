 
README - Projeto VBA de Automação de Relatórios Financeiros

OBJETIVO

Este projeto automatiza a rotina de processamento, atualização e salvamento de relatórios financeiros diários para uma empresa do ramo financeiro.
As informações oficiais foram substituídas, restando apenas a lógica de execução de um projeto VBA em ambiente profissional.
O projeto é apresentado de forma simplificada para fins educacionais e pode ser adaptado conforme a necessidade de cada ambiente profissional.

ESTRUTURA DO PROJETO

- Arquivo principal: Executável.xlsm (formato habilitado para macros)
- Aba de controle: "EXEC"
- Módulos VBA:
  - Modulo1: Executável principal, declara variáveis e chama funções
  - Modulo2: Funções divididas por etapas

CONFIGURAÇÕES INICIAIS

Na aba "EXEC", configure:
- MÊS e ANO: para organização de pastas e subpastas
- CAMINHO: define o endereço completo e nome do arquivo final, incluindo hora da execução

IMPORTANTE: Substitua a partição inicial do caminho (“K:\”) pelo caminho correspondente à sua máquina local.

EXEMPLO DE CAMINHO EM MINHA MÁQUINA
K:\Automação com VBA\RELATÓRIOS ACOMP. VENDAS DIÁRIO\2025\OUTUBRO\ACOMP. VENDAS DIÁRIO - 2 OUTUBRO 2025 - 17.24.32Hrs

FUNCIONAMENTO DO CÓDIGO VBA

Modulo1 - Executável principal:
- Declara variáveis de caminho
- Chama funções do Modulo2

Modulo2 - Funções por etapa:
1. Atualiza consultas do PowerQuery
2. Copia tabela da aba "BASE" do arquivo EXEC
3. Abre arquivo Matriz e limpa conteúdo da aba "BASE PRODUÇÃO" (linha 3 em diante)
4. Reabre arquivos e realiza cópia entre EXEC e Matriz
5. Atualiza o dia útil do mês (aba "CONSULTOR - RENDA FIXA", célula “O2”)
6. Cria array com nomes das abas e limpa fórmulas via looping
7. Salva o arquivo final no caminho especificado na célula “C2”

COMO EXECUTAR
1. Habilite a guia "Desenvolvedor" no Excel
2. Acesse o editor VBA: Desenvolvedor > Visual Basic
3. Navegue pelos módulos e execute o código conforme necessário

OBSERVAÇÕES
- Certifique-se de que os arquivos referenciados estejam no local correto
- O formato ".xlsm" é obrigatório para execução de macros
- O projeto pode ser adaptado para diferentes rotinas de trabalho

USO PROFISSIONAL OU EDUCACIONAL
Este projeto pode ser utilizado como base para estudos de automação com VBA,
integração com PowerQuery e manipulação de dados em Excel.