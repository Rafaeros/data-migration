"""
Main file for the project.
"""

from rich.console import Console

from core.get_data import get_data


def main() -> None:
    """
    Main function to run the project.
    """
    console = Console()
    console.print("[bold magenta]Bem vindo ao projeto![/bold magenta]")
    console.print("[bold green]Feito por: Rafael Costa[/bold green]")
    console.print(
        "[bold green]Github:[/bold green] [link=https://github.com/rafaeros]https://github.com/rafaeros[/link]"
    )
    console.print("[bold blue]Coloque o caminho do arquivo Excel:[/bold blue]")
    path: str = input()
    get_data(path)
    print("[bold green]Arquivo lido com sucesso![/bold green]")


if __name__ == "__main__":
    main()
