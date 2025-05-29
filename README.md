<hr>

# 📦 Projeto de Automação de Pedidos de Compra e Nota Fiscal
Este repositório contém o código-fonte para um projeto de automação de pedidos de compra, focado na migração e análise de dados de estoque. Utilizando bibliotecas como pandas, openpyxl e selenium, o projeto automatiza a extração de dados de planilhas Excel e a criação de pedidos de compra em sistemas web.


## 📋 Pré-requisitos
Antes de executar o projeto, verifique se os seguintes softwares estão instalados no seu sistema:

* Git
* Python 3.10+
* Outlook
* uv (gerenciador de dependências)

Para instalar o uv, utilize:

```bash
pip install uv
```

Se estiver utilizando PowerShell no Windows, altere a política de execução de scripts para permitir scripts locais:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 📁 Estrutura do Projeto
```pgsql
data-migration/
├── main.py
├── core/
│   ├── utils/
│   |   └── send_email.py
│   ├── get_data.py
|   |── create_invoice.py
│   └── create_orders.py
├── data_analysis.ipynb
├── pyproject.toml
├── uv.lock
├── .python-version
└── .gitignore
main.py: Arquivo principal para execução do projeto.
```

- **core/get_data.py**: Responsável por obter dados de uma planilha Excel.
- **core/create_orders.py**: Cria pedidos de compra utilizando Selenium.
- **core/create_invoice.py**: Cria notas fiscais utilizando Selenium.
- **core/utils/send_email.py**: Envia e-mails utilizando a biblioteca PyWin32.
- **data_analysis.ipynb**: Notebook Jupyter para análise dos dados extraídos.
- **pyproject.toml**: Arquivo de configuração das dependências do projeto.
- **uv.lock: Arquivo** de bloqueio contendo as versões exatas das dependências.
- **.python-version**: Especifica a versão do Python utilizada no projeto.

## 🚀 Como Executar

Clone o repositório:

```bash
git clone https://github.com/Rafaeros/data-migration.git
cd data-migration
```

Crie e ative o ambiente virtual:

No Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
```

No Unix/MacOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instale as dependências com uv:

```bash
uv pip install -r pyproject.toml
```

Execute o script principal:

```bash
python main.py
```

## 🛠️ Dependências Principais
- **pandas**: Manipulação e análise de dados.
- **openpyxl**: Manipulação de planilhas Excel.
- **selenium**: Automação de navegadores web.
- **pywin32**: Interação com o Outlook para envio de e-mails.
- **rich**: Biblioteca de formatação de texto para terminal.


## 📊 Análise de Dados

O notebook data_analysis.ipynb fornece uma análise exploratória dos dados extraídos, permitindo insights sobre o estoque e os pedidos de compra.


## 📄 Licença

Este projeto está licenciado sob a [MIT License](LICENSE).  
Você pode utilizá-lo, modificá-lo e distribuí-lo conforme os termos dessa licença.

<hr>