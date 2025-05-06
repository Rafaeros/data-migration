"""
This module is responsible for creating orders on the website.
It uses Selenium to automate the login process and navigate through the website.
"""

import json
import time
import pyautogui as pygui
from selenium import webdriver
from selenium.webdriver.common.by import By
import selenium.common.exceptions
from rich.console import Console


def create_orders(username: str, password: str, json_file_paths: list[str]) -> None:
    """Login to the website using Selenium."""
    login_url: str = "https://app.cargamaquina.com.br/site/login?c=31.1~78%2C8%5E56%2C8"
    driver = webdriver.Chrome()
    try:
        driver.get(login_url)
        time.sleep(10)
        driver.find_element(By.ID, "LoginForm_username").send_keys(username)
        driver.find_element(By.ID, "LoginForm_password").send_keys(password)
        driver.find_element(By.NAME, "yt0").click()
        time.sleep(10)
    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.ElementNotInteractableException as e:
        print(f"Error: {e}")

    console = Console()
    console.print("[bold green]Login realizado com sucesso![/bold green]")
    console.print("[bold blue]Iniciando criação dos pedidos de compra...[/bold blue]")
    driver.get("https://app.cargamaquina.com.br/compra/pedidoCompra")
    time.sleep(10)
    for json_file_path in json_file_paths:
        with open(json_file_path, "r", encoding="utf-8") as f:  # type: ignore
            data = json.load(f)
            for orders in data:
                try:
                    driver.find_element(By.ID, "btIncluir").click()
                    driver.implicitly_wait(10)
                    supplier = driver.find_element(
                        By.XPATH, "//*[@id='s2id_sel2Fornecedor']"
                    )
                    supplier.click()
                    pygui.write("f&k group tecnologia em sistemas automotivos ltda")
                    pygui.press("enter")
                    for item in orders["itens"]:
                        driver.find_element(By.ID, "adicionarItensCompra").click()
                        time.sleep(10)
                        pygui.press("tab", presses=2)
                        pygui.write(item["CODÍGO"])
                        time.sleep(5)
                        pygui.press("enter")
                        pygui.write(str(item["SALDO TOTAL"]).replace(".", ","))
                        pygui.press("tab")
                        time.sleep(2)
                        if item["CUSTO UNITARIO"] == 0:
                            pygui.write("1,00")
                        else:
                            pygui.write(str(item["CUSTO UNITARIO"]).replace(".", ","))
                        time.sleep(2)
                        pygui.press("tab", presses=3)
                        time.sleep(2)
                        pygui.write("15/06/2025")
                        driver.find_element(By.ID, "adicionarItemCompraGrid").click()
                        driver.implicitly_wait(10)
                except selenium.common.exceptions.NoSuchElementException as e:
                    print(f"Error: {e}")
                except selenium.common.exceptions.ElementNotInteractableException as e:
                    print(f"Error: {e}")
                except selenium.common.exceptions.StaleElementReferenceException as e:
                    print(f"Error: {e}")
                except Exception as e:
                    print(f"Error: {e}")
                    input("Press Enter to continue...")


if __name__ == "__main__":
    create_orders(
        "username",
        "password",
        ["data.json"]
    )
