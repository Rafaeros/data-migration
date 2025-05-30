"""
This script is used to read an Excel file and print the names of the sheets in it.
"""

import os
import json
from getpass import getpass
import pathlib
from math import ceil
import pandas as pd
from rich.console import Console
from rich.text import Text
from core.create_orders import create_orders


def format_sheet_data(df: pd.DataFrame) -> list[str]:
    """Format the sheet data by renaming columns and filtering rows."""
    df = df[
        [
            "CODÍGO",
            "FINALIDADE",
            "PROPRIETARIO",
            "SALDO TOTAL",
            "UNIDADE",
            "CUSTO UNITARIO",
            "TIPO/PROPRIETARIO",
            "DECLARA",
            "NCM",
        ]
    ]
    owners: list[str] = [
        "F&K GROUP TECNOLOGIA EM SISTEMAS AUTOMOTIVOS LTDA.",
        "ESTOQUE DEUTSCH",
        "ITENS OBSOLETOS",
        "OVERSTOCK",
    ]
    df["SALDO TOTAL"] = pd.to_numeric(df["SALDO TOTAL"], errors="coerce")
    df = df[df["SALDO TOTAL"] > 0]
    df = df[df["NCM"] > 0]
    df = df[df["PROPRIETARIO"].isin(owners)]

    groups = df.groupby("TIPO/PROPRIETARIO")
    json_file_paths: list[str] = []
    for owner_type, group in groups:
        group = group.reset_index(drop=True)
        order_number = ceil(len(group) / 15)
        orders: list[dict] = []

        for i in range(order_number):
            fat = group.iloc[i * 15 : (i + 1) * 15]

            orders.append(
                {
                    "tipo_proprietario": owner_type,
                    "pedido_numero": i + 1,
                    "quantidade_itens": len(fat),
                    "itens": fat.drop(columns=["TIPO/PROPRIETARIO"]).to_dict(
                        orient="records"
                    ),
                }
            )

        tmp_folder = pathlib.Path("./tmp/json/")
        tmp_folder.mkdir(parents=True, exist_ok=True)

        if "." in owner_type[-1]:
            json_file: str = (
                "./tmp/json/" + owner_type.lower().replace("/", "_") + "json"
            )
        else:
            json_file: str = (
                "./tmp/json/" + owner_type.lower().replace("/", "_") + ".json"
            )
        json_file_paths.append(json_file)
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(orders, f, indent=4, ensure_ascii=False)

    print(f"Total de pedidos: {len(orders)}")
    if os.path.exists("./tmp/pedidos_enviados.txt"):
        with open("./tmp/pedidos_enviados.txt", "r", encoding="utf-8") as f:
            orders_sent = f.read().splitlines()
            sent_orders: list[str] = []
            for order in orders_sent:
                if order in json_file_paths:
                    sent_orders.append(order)
                    json_file_paths.remove(order)

            print("Pedidos já Criados:")
            for i, order in enumerate(sent_orders):
                print(f"{i + 1} - {order}")

            return json_file_paths

    return json_file_paths


def get_data(path: str) -> None:
    """Read the Excel file and print the names of the sheets."""
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    console = Console()
    console.print("[bold blue]Planilhas encontradas:[/bold blue]")

    for i, sheet in enumerate(sheets):
        text = Text(f"{i + 1} - {sheet}")
        text.stylize("bold green")
        console.print(text)
    console.print("[bold blue]Total de planilhas:[/bold blue]", len(sheets))

    df = pd.DataFrame()

    while df.empty:
        console.print(
            "[bold blue]Digite o número da planilha que deseja ler:[/bold blue]"
        )
        try:
            sheet_number: int = int(input())
            if 1 <= sheet_number <= len(sheets):
                sheet_option = sheets[sheet_number - 1]
                console.print(
                    f"[bold blue]Você escolheu a planilha:[/bold blue] {sheet_option}"
                )
                df = pd.read_excel(path, sheet_name=sheet_option)
            else:
                console.print("[bold red]Número inválido! Tente novamente.[/bold red]")
        except ValueError:
            console.print("[bold red]Entrada inválida! Tente novamente.[/bold red]")

    console.print("[bold green]Planilha lida com sucesso![/bold green]")
    console.print(
        "[bold blue]Iniciando modulo de criação dos pedidos de compra...[/bold blue]"
    )
    json_file_paths: list[str] = format_sheet_data(df)
    console.print("[bold blue]Insira suas credenciais:[/bold blue]")
    console.print("[bold blue]Usuário:[/bold blue]")
    username: str = input()
    console.print("[bold blue]Senha:[/bold blue]")
    password: str = getpass()
    console.print("[bold blue]Iniciando login...[/bold blue]")
    create_orders(username, password, json_file_paths)


if __name__ == "__main__":
    get_data(r"C:\Users\user\Documents\Analise.xlsx")
    print("Script executed successfully")
