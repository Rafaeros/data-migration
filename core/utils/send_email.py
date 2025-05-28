"""
Module to send email to finance department
"""

import win32com.client as win32
import pandas as pd
from rich.console import Console


def send_order_email(order_data, rateio) -> None:
    """Send email to finance department"""
    outlook = win32.Dispatch("Outlook.Application")

    tipo_proprietario: str = order_data["tipo_proprietario"]
    pedido_numero: int = order_data["pedido_numero"]
    quantidade_itens: int = order_data["quantidade_itens"]
    itens = pd.DataFrame(order_data["itens"]).reset_index(drop=True)
    itens.index += 1
    itens["CUSTO UNITARIO"] = itens["CUSTO UNITARIO"].apply(
        lambda x: str(x).replace(".", ",")
    )
    itens_html = itens.to_html(col_space=50, justify="center")

    style: str = """
        <style>
        /* Aplica padding e cor de texto em todos os elementos */
        * {
            padding: 5px;
            color: black;
            box-sizing: border-box; /* Inclui padding e border na largura e altura total do elemento */
        }
        
        p {
            font-size: 16px;
            font_weight: bold;
        }

        /* Estilo para o cabeçalho da tabela */
        thead {
            text-align: center;
            background-color: #6c17ec;
            color: white; /* Texto branco para contraste */
        }

        /* Estilo para linhas, cabeçalhos e células da tabela */
        tr, th, td {
            text-align: center;
            vertical-align: middle; /* Alinhamento vertical ao centro */
            padding: 10px; /* Padding para melhorar espaçamento */
        }

        /* Estilo alternado para linhas da tabela */
        tr:nth-child(even) {
            background-color: #f2f2f2; /* Fundo cinza claro */
        }

        /* Estilo para células específicas */
        td:nth-child(5) {
            text-align: left;
            background-color: #FFCDD2; /* Vermelho claro para melhor legibilidade */
        }

        /* Estilo para a borda da tabela */
        table {
            border-collapse: collapse; /* Remove espaçamento entre células */
            width: 100%; /* Largura total */
        }

        /* Estilo para bordas das células */
        th, td {
            border: 1px solid #ddd; /* Borda cinza clara */
        }

        /* Estilo para hover em linhas da tabela */
        tr:hover {
            background-color: #ddd; /* Fundo cinza claro ao passar o mouse */
        }
        </style>
        """

    email_body: str = f"""
        <!DOCTYPE html>
        <html>
            <head>
                {style}
            </head>
            <body>
                <h1>Equipe Financeira, Nota Fiscal Armazenada com as informações abaixo:</h1>
                
                <h2>Pedido de Compra - N°{pedido_numero} Criado</h2>
                <h3>Tipo de Proprietário: {tipo_proprietario}</h3>
                <h3>Quantidade de Itens no Pedido: {quantidade_itens}, Rateio: {rateio}</h3>
                </br>
                {itens_html}
                </br>
                
                <p>Por favor, verifiquem os dados e procedam com a emissão da NF no sistema.</p>
                
                <p>Caso haja alguma divergência ou necessidade de ajustes, entrem em contato comigo para corrigir.</p>
                
                <p>Agradeço pela atenção!</p>
                
                <p>Esta é uma mensagem automática. Por favor, não responda diretamente a este e-mail.</p>
            </body>
        </html>
    """
    console = Console()
    try:
        mail = outlook.CreateItem(0)
        mail.To = "email@domain"
        mail.CC = "email@domain"
        mail.Subject = (
            f"NF Armazenada - Pedido de Compra N° {pedido_numero} - {tipo_proprietario}"
        )
        mail.HTMLBody = email_body
        mail.Send()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        outlook.Quit()

    console.print("[bold green]Email enviado com sucesso![/bold green]")


if __name__ == "__main__":
    send_order_email(
        {
            "tipo_proprietario": "Embalagem / Estoque",
            "pedido_numero": 1,
            "quantidade_itens": 2,
            "itens": [
                {
                    "CODÍGO": "CXPP01G",
                    "FINALIDADE": "Embalagens",
                    "PROPRIETARIO": "Estoque",
                    "SALDO TOTAL": 128.0,
                    "UNIDADE": "un",
                    "CUSTO UNITARIO": 0.9291,
                    "DECLARA": "DECLARA",
                    "NCM": 41854,
                },
                {
                    "CODÍGO": "CXPP01M",
                    "FINALIDADE": "Embalagens",
                    "PROPRIETARIO": "Estoque",
                    "SALDO TOTAL": 240.0,
                    "UNIDADE": "un",
                    "CUSTO UNITARIO": 0.170416667,
                    "DECLARA": "DECLARA",
                    "NCM": 546465,
                },
            ],
        },
        "EMBALAGEM",
    )
