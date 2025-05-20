"""
This module is responsible for creating orders on the website.
It uses Selenium to automate the login process and navigate through the website.
"""

import json
import time
import pyautogui as pygui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import selenium.common.exceptions
from rich.console import Console
from rich.text import Text

from utils.send_email import send_order_email

def create_orders(username: str, password: str, json_file_paths: list[str]) -> None:
    """Login to the website using Selenium."""
    login_url: str = (
        "https://web.cargamaquina.com.br/site/login?c=31.1%7E78%2C8%5E56%2C8"
    )
    options = Options()
    options.add_argument("--force-device-scale-factor=0.75")  # 80% de zoom
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    wait = WebDriverWait(driver, 15)
    try:
        driver.get(login_url)
        driver.implicitly_wait(10)
        driver.find_element(By.ID, "LoginForm_username").send_keys(username)
        driver.find_element(By.ID, "LoginForm_password").send_keys(password)
        driver.find_element(By.NAME, "yt0").click()
        driver.implicitly_wait(10)
    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.ElementNotInteractableException as e:
        print(f"Error: {e}")

    console = Console()
    console.print("[bold green]Login realizado com sucesso![/bold green]")
    console.print("[bold blue]Iniciando criação dos pedidos de compra...[/bold blue]")
    time.sleep(3)
    driver.get("https://web.cargamaquina.com.br/compra/pedidoCompra")
    pygui.shortcut("alt", "tab")
    order_option: str = ""
    rateios: list[str] = [
        "MATERIAL DE USO E CONSUMO [MAT-USO]",
        "CUSTOS E SERVIÇOS - IMPORTAÇÃO [CUS-IMP]",
        "EQUIPAMENTOS PROTEÇÃO INDIVIDUAL/UNIFORME [EPI]",
        "FEIRAS - EVENTOS [DESP-FEIR]",
        "MANUTENÇÃO DE MAQUINAS E EQUIPAMENTOS [MAN-MAQ]",
        "SOFTWARE E HARDWARE [MAN-SOFT]",
        "MANUTENÇÃO DE VEICULO [MAN-VEIC]",
        "MANUTENÇÃO PREDIAL [MAN-PRE]",
        "MATERIAL DE ESCRITÓRIO [MAT-ESCR]",
        "MATERIAL DE LIMPEZA [MAT-LIMP]",
        "PROPAGANDA E PUBLICIDADE [PRO-PUB]",
        "SERVIÇOS DE TERCEIROS PRODUÇÃO [SERV-TER]",
        "VIAGENS / ESTADIAS / PEDÁGIOS [VIA-EST]",
        "MATERIA-PRIMA [MAT-MP]",
        "MATERIA PRIMA REVENDA [MP-REV]",
        "IMOBILIZADO [IMO]",
        "GASTOS GERAIS DE FABRICAÇÃO [KANBAN]",
        "MATÉRIA PRIMA CABOS [MP-CAB]",
        "NOVO PROJETO [NPROJ]",
        "CUSTO FIXO [CFIX]",
        "INVESTIMENTO [INVST]",
        "ALIMENTAÇÃO/BENEFÍCIOS FUNCIONÁRIOS [ALIMBEF]",
        "MATERIA PRIMA INDUSTRIALIZAÇÃO [MP-IND]",
        "NOVO PROJETO - ELETRONICO [PROJ - ELE]",
        "EMBALAGEM (MAT EMBALAGEM) [MAT EMB]",
        "CONFRATERNIZAÇÃO [CONF MKT]",
        "CONFRATERNIZAÇÃO [FESTA]",
        "DECORAÇÃO [DEC MKT]",
        "DECORAÇÃO [ENDO MKT]",
        "BENEFÍCIOS [CEST BAS]",
        "COPA E COZINHA [COPA]",
        "HIGIENE E LIMPEZA [HIG LIMP]",
        "MATERIAL DE ESCRITÓRIO [MAT ESC]",
        "FERRAMENTAS [FERRAM-PRO]",
        "FERRAMENTAS [FERRAM]",
        "GASTOS GERAIS PREDIAL [GAST PRED]",
    ]

    while order_option != "quit":
        for i, json_file in enumerate(json_file_paths):
            text = Text(f"{i + 1} - {json_file}\n")
            text.stylize("bold green")
            console.print(text)
        console.print(
            "[bold blue]Digite o número do pedido que deseja criar:[/bold blue]"
        )
        order_option: str = input()
        time.sleep(1)

        for i, rateio in enumerate(rateios):
            text = Text(f"{i + 1} - {rateio}\n")
            text.stylize("bold green")
            console.print(text)
        console.print("[bold blue]Selecione o rateio dos pedidos:[/bold blue]")
        rateio_option: str = input()
        time.sleep(1)
        pygui.shortcut("alt", "tab")
        rateio = rateios[int(rateio_option) - 1]
        json_file_path = json_file_paths[int(order_option) - 1]
        time.sleep(1)
        with open(json_file_path, "r", encoding="utf-8") as f:  # type: ignore
            orders = json.load(f)
            for order in orders:
                try:
                    console.print(
                        f"[bold blue]Criando pedido de compra:[/bold blue] {order['pedido_numero']}"
                    )
                    console.print(
                        f"[bold blue]TIPO/PROPRIETARIO[/bold blue] {order['tipo_proprietario']}"
                    )
                    console.print(f"[bold blue]Rateio: [/bold blue] {rateio}")

                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='btIncluir']")
                        )
                    ).click()
                    supplier = wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='s2id_sel2Fornecedor']")
                        )
                    )
                    supplier.click()
                    time.sleep(2)
                    pygui.write("f&k group tecnologia em sistemas automotivos ltda")
                    pygui.press("enter")
                    time.sleep(2)
                    # Adding items in order
                    for item in order["itens"]:
                        time.sleep(2)
                        console.print(
                            f"[bold blue]Adicionando item:[/bold blue] {item['CODÍGO']}"
                        )
                        wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='adicionarItensCompra']")
                            )
                        ).click()
                        wait.until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//*[@id='modal-item-compra-form']/div")
                            )
                        )
                        time.sleep(2)
                        pygui.press("tab", presses=2)
                        time.sleep(2)
                        pygui.write(item["CODÍGO"])
                        time.sleep(2)
                        pygui.press("enter")
                        time.sleep(2)
                        pygui.write(str(item["SALDO TOTAL"]).replace(".", ","))
                        pygui.press("tab")
                        time.sleep(2)
                        if item["CUSTO UNITARIO"] == 0:
                            pygui.write("1,00")
                            item["CUSTO UNITARIO"] = 1
                        elif item["CUSTO UNITARIO"] < 0:
                            pygui.write(
                                str(abs(item["CUSTO UNITARIO"])).replace(".", ",")
                            )
                            item["CUSTO UNITARIO"] = abs(item["CUSTO UNITARIO"])
                        else:
                            pygui.write(str(item["CUSTO UNITARIO"]).replace(".", ","))
                        time.sleep(2)
                        pygui.press("tab", presses=3)
                        time.sleep(2)
                        pygui.write("20/06/2025")
                        wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='adicionarItemCompraGrid']")
                            )
                        ).click()
                    # Finished to add the 15 items in order, now we need to go to the invoice
                    time.sleep(2)
                    # Invoice
                    invoice = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                '//*[@id="tabDadosGerais"]/div[5]/div/div/ul/li/a',
                            )
                        )
                    )
                    driver.execute_script("arguments[0].click();", invoice)

                    invoice_time = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//*[@id="s2id_sel2PrazoPR"]')
                        )
                    )
                    driver.execute_script(
                        "arguments[0].scrollIntoView();", invoice_time
                    )
                    time.sleep(2)
                    invoice_time.click()
                    time.sleep(2)
                    pygui.write("30 dias")
                    time.sleep(2)
                    pygui.press("enter")
                    time.sleep(2)

                    wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='gerarParcelas']")
                        )
                    ).click()
                    time.sleep(2)
                    pygui.shortcut("ctrl", "end")
                    # Choosing the purpose of the items in the order
                    purpose = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='s2id_sel2Rateio']")
                        )
                    )
                    time.sleep(2)
                    purpose.click()
                    time.sleep(2)
                    pygui.write(rateio)
                    time.sleep(2)
                    pygui.press("enter")
                    time.sleep(2)

                    wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='btnGravarCompra']")
                        )
                    ).click()

                    time.sleep(2)
                    console.print(
                        f"[bold green]Pedido de compra criado com sucesso:[/bold green] {order['pedido_numero']}"
                    )

                    send_order_email(order, rateio)
                    pygui.shortcut("alt", "tab")

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
    create_orders("username", "password", ["./data.json"])
