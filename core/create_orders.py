"""
This module is responsible for creating orders on the website.
It uses Selenium to automate the login process and navigate through the website.
"""

import os
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

from core.create_invoice import create_invoice


def create_orders(username: str, password: str, json_file_paths: list[str]) -> None:
    """Login to the website using Selenium."""
    login_url: str = "https://v2.cargamaquina.com.br/site/login/c/3.1~13,3%5e17,7"
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
    time.sleep(4)
    driver.get("https://v2.cargamaquina.com.br/compra/pedidoCompra")
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
        time.sleep(2)

        for i, rateio in enumerate(rateios):
            text = Text(f"{i + 1} - {rateio}\n")
            text.stylize("bold green")
            console.print(text)
        console.print("[bold blue]Selecione o rateio dos pedidos:[/bold blue]")
        rateio_option: str = input()
        time.sleep(2)
        rateio = rateios[int(rateio_option) - 1]
        json_file_path = json_file_paths[int(order_option) - 1]
        time.sleep(2)
        pygui.shortcut("alt", "tab")
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
                    time.sleep(4)
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
                    time.sleep(4)
                    pygui.write("f&k group tecnologia em sistemas automotivos ltda",
                                interval=0.1)
                    pygui.press("enter")
                    time.sleep(4)
                    # Adding items in order
                    for item in order["itens"]:
                        time.sleep(4)
                        console.print(
                            f"[bold blue]Adicionando item:[/bold blue] {item['CODÍGO']}"
                        )
                        wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='adicionarItensCompra']")
                            )
                        ).click()
                        time.sleep(4)
                        wait.until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//*[@id='modal-item-compra-form']/div")
                            )
                        )
                        time.sleep(4)
                        pygui.press("tab", presses=2)
                        time.sleep(4)
                        code = str(item["CODÍGO"])
                        pygui.write(code, interval=0.2)
                        time.sleep(4)
                        pygui.press("enter")
                        time.sleep(4)
                        qty = item["SALDO TOTAL"]
                        if float(qty).is_integer():
                            qty = str(int(qty))
                        else:
                            qty = str(qty).replace(".", ",")

                        pygui.write(qty, interval=0.1)
                        item["SALDO TOTAL"] = qty
                        pygui.press("tab")
                        time.sleep(4)

                        if item["CUSTO UNITARIO"] == 0:
                            pygui.write("1,00")
                        elif item["CUSTO UNITARIO"] < 0:
                            pygui.write(
                                str(abs(item["CUSTO UNITARIO"])).replace(".", ",")
                            )
                        else:
                            cost = f"{float(item['CUSTO UNITARIO']):.3f}".replace(
                                ".", ","
                            )
                            pygui.write(cost, interval=0.1)
                        time.sleep(4)
                        pygui.press("tab", presses=3)
                        time.sleep(4)
                        pygui.write("20/06/2025")
                        wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='adicionarItemCompraGrid']")
                            )
                        ).click()
                    # Finished to add the 15 items in order, now we need to go to the invoice
                    time.sleep(4)
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
                    pygui.shortcut("ctrl", "end")
                    time.sleep(4)
                    invoice_time.click()
                    time.sleep(4)
                    pygui.write("30 dias")
                    time.sleep(4)
                    pygui.press("enter")
                    time.sleep(4)

                    wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='gerarParcelas']")
                        )
                    ).click()
                    time.sleep(4)
                    pygui.shortcut("ctrl", "end")
                    # Choosing the purpose of the items in the order
                    purpose = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='s2id_sel2Rateio']")
                        )
                    )
                    time.sleep(4)
                    purpose.click()
                    time.sleep(4)
                    pygui.write(rateio, interval=0.1)
                    time.sleep(4)
                    pygui.press("enter")
                    time.sleep(4)

                    console.print("[bold blue]Gravando pedido de compra[/bold blue]")
                    wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='btnGravarCompra']")
                        )
                    ).click()
                    time.sleep(4)

                    console.print(
                        "[bold blue]Pegando valor do pedido de compra[/bold blue]"
                    )
                    order_number = wait.until(
                        EC.presence_of_element_located(
                            (
                                By.XPATH,
                                "//*[@id='ordem-compra-grid']/table/tbody/tr[1]/td[3]",
                            )
                        )
                    ).text
                    time.sleep(4)
                    order["pedido_numero"] = order_number

                    time.sleep(4)
                    console.print(
                        f"[bold green]Pedido de compra criado com sucesso:[/bold green] {order['pedido_numero']}"
                    )
                except selenium.common.exceptions.NoSuchElementException as e:
                    print(f"Error: {e}")
                    return
                except selenium.common.exceptions.ElementNotInteractableException as e:
                    print(f"Error: {e}")
                    return
                except selenium.common.exceptions.StaleElementReferenceException as e:
                    print(f"Error: {e}")
                    return
                except Exception as e:
                    print(f"Error: {e}")
                    input("Press Enter to continue...")
                    return
            driver.quit()
            time.sleep(6)
            console.print("[bold blue]Iniciando criação das notas fiscais[/bold blue]")
            create_invoice(username, password, orders, rateio)

        sended_orders_path: str = "./tmp/pedidos_enviados.txt"
        if not os.path.exists(sended_orders_path):
            with open(sended_orders_path, "w", encoding="utf-8") as f:
                f.write("pedidos_enviados\n")
                f.write(f"{json_file_path}\n")
                json_file_paths.remove(json_file_path)
        else:
            print("O arquivo de pedidos enviados já existe!")
            with open(sended_orders_path, "a", encoding="utf-8") as f:
                f.write(f"{json_file_path}\n")
                json_file_paths.remove(json_file_path)

        pygui.shortcut("alt", "tab")


if __name__ == "__main__":
    create_orders(
        "username",
        "password",
        ["./data.json"],
    )
