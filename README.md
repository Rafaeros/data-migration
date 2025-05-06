# Projeto de Automação de Pedidos de Compra

Este repositório contém o código-fonte para o projeto de migração de dados, que tem como objetivo processar e analisar dados de estoque. O projeto utiliza diversas bibliotecas Python, como `pandas` e `openpyxl`, para manipulação e tratamento dos dados.

## Estrutura do Projeto

- **main.py**: Arquivo principal para execução do projeto.
- **core/get_data.py**: Script responsável por obter dados de uma planilha Excel.
- **core/create_orders.py**: Script responsável por criar pedidos de compra utilizando Selenium.
- **data_analysis.ipynb**: Notebook Jupyter utilizado para análise dos dados extraídos.
- **uv.lock**: Arquivo de bloqueio contendo as dependências do projeto.

## Como Executar

1. Certifique-se de ter o Python 3.10 instalado.
2. Instale as dependências necessárias utilizando:

   ```bash
   pip install -r pyproject.toml
   ```

3. Execute o script principal:

   ```bash
   python main.py
   ```

## Dependências

As principais bibliotecas utilizadas no projeto são:

- `pandas`
- `openpyxl`
- `selenium`
- `rich`

## Licença

Este projeto está licenciado sob a MIT License - veja o arquivo LICENSE para mais detalhes.