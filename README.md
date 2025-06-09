# 📚 Gerador de Planilha de Estoque de Livros

Este projeto em Python gera automaticamente uma planilha Excel simulando o estoque de uma livraria. Ele cria dados fictícios, calcula o preço de venda com margem de lucro, além de estimar o lucro total e o valor total do estoque.

Depois disso, os dados são **tratados e visualizados no Power BI**, permitindo análises gráficas e dashboards interativos.

## 🚀 Funcionalidades

- Geração automática de dados para livros fictícios
- Cálculo de preço de venda com base na margem de lucro
- Estimativa de lucro total por produto
- Cálculo do valor total do estoque
- Ajuste automático da largura das colunas no Excel
- Integração com Power BI para visualização e análise dos dados

## 🛠 Tecnologias utilizadas

- [Python 3](https://www.python.org/)
- [openpyxl](https://openpyxl.readthedocs.io/) — para manipulação de arquivos Excel
- [Power BI](https://powerbi.microsoft.com/) — para visualização dos dados
- Bibliotecas padrão: `random`, `datetime`

## 📊 Tratamento e Análise com Power BI

Após a geração do arquivo `estoque_livros.xlsx`, o Power BI é utilizado para:

- Limpeza e transformação dos dados (Power Query)
- Criação de gráficos de:
  - Livros mais lucrativos
  - Editoras com maior quantidade de livros
  - Distribuição de notas
- Cálculo de KPIs como:
  - Valor total em estoque
  - Lucro total estimado
  - Margens médias por editora

## 💼 Estrutura da planilha gerada

A planilha contém as seguintes colunas:

- **Nome do Livro**
- **Editora**
- **Valor de Compra**
- **Notas**
- **Quantidade**
- **Preço de Venda**
- **Lucro Total**
- **Valor Total**

As fórmulas de **lucro** e **valor total** são calculadas diretamente no Excel.

## 🧪 Como usar

1. Clone este repositório:
   ```bash
   git remote set-url origin https://github.com/OFabioSilvaa/Gerador-de-Planilha-de-Estoque-de-Livros.git
