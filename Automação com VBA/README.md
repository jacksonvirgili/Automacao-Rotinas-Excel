# Projeto de Automação de Relatórios Comerciais
## Sobre o Projeto

Este projeto foi desenvolvido com foco na acessibilidade de informações para toda a rede de colaboradores.
Seu principal objetivo é fornecer indicadores (KPI’s) comerciais atualizados diariamente, permitindo que coordenadores e supervisores tomem decisões rápidas e baseadas em dados.

Trata-se de uma solução desenvolvida e aplicada para a equipe comercial de uma empresa do setor financeiro.
Todas as informações sigilosas foram substituídas, mantendo apenas a estrutura lógica e funcional do projeto.

## Estrutura de Pastas

📁 ARQUIVOS DA REDE
    └── Contém os arquivos com informações sobre hierarquias e metas.

📁 RELATÓRIOS ACOMP. VENDAS DIÁRIO
    ├── Matriz de KPI’s.
    └── Subpastas organizadas por ano e mês (organizam o armazenamento dos relatórios finais).

📄 Executável.xlsm
    └── Arquivo principal habilitado para macros (VBA).

📄 Readme.txt
    └── Este documento.

## Configuração Inicial

Dentro do arquivo Executável.xlsm, acesse a aba “EXEC” e configure:

Célula C2 → Insira o **caminho** dos diretórios e o **nome** desejado para o arquivo final (com data).

💡 Exemplo de Caminho (na minha máquina)

K:\Automação com VBA\RELATÓRIOS ACOMP. VENDAS DIÁRIO\2025\OUTUBRO\ (caminho)
ACOMP. VENDAS DIÁRIO - 2 OUTUBRO 2025 - 17.24.32Hrs (nome arquivo)

## Observações

- Certifique-se de que os diretórios e subpastas já existam antes da execução.

- As macros devem estar habilitadas para o funcionamento correto do script.

## Resumo de Funções

1. Atualiza consultas do PowerQuery
2. Copia tabela da aba "BASE" do arquivo EXEC
3. Abre arquivo Matriz e limpa conteúdo da aba "BASE PRODUÇÃO" (linha 3 em diante)
4. Reabre arquivos e realiza cópia entre EXEC e Matriz
5. Atualiza o dia útil do mês (aba "CONSULTOR - RENDA FIXA", célula “O2”)
6. Cria array com nomes das abas e limpa fórmulas via looping
7. Salva o arquivo final no caminho especificado na célula “C2”

# USO PROFISSIONAL OU EDUCACIONAL

Este projeto pode ser utilizado como base para estudos de automação com VBA,

integração com PowerQuery e manipulação de dados em Excel.

