"""
This module is responsible for creating orders on the website.
It uses Selenium to automate the login process and navigate through the website.
"""

import re
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

from core.utils.send_email import send_order_email
from core.create_invoice_request import create_invoice_request


def create_orders(username: str, password: str, json_file_paths: list[str]) -> None:
    """Login to the website using Selenium."""
    login_lx: str = "https://v2.cargamaquina.com.br/site/login/c/3.1~13,3%5e17,7"
    login_fk: str = "https://app.cargamaquina.com.br/site/login?c=31.1~78%2C8%5E56%2C8"
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--force-device-scale-factor=0.65")  # 80% de zoom
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)
    console = Console()
    try:
        driver.get(login_lx)
        driver.implicitly_wait(10)
        driver.find_element(By.ID, "LoginForm_username").send_keys(username)
        driver.find_element(By.ID, "LoginForm_password").send_keys(password)
        driver.find_element(By.NAME, "yt0").click()
        driver.implicitly_wait(10)
        console.print("[bold green]Login realizado com sucesso na LANX![/bold green]")
        driver.get("https://v2.cargamaquina.com.br/compra/pedidoCompra")
        driver.implicitly_wait(10)
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(login_fk)
        driver.implicitly_wait(10)
        driver.find_element(By.ID, "LoginForm_username").send_keys(username)
        driver.find_element(By.ID, "LoginForm_password").send_keys(password)
        driver.find_element(By.NAME, "yt0").click()
        driver.implicitly_wait(10)
        console.print("[bold green]Login realizado com sucesso na F&K![/bold green]")
        driver.get("https://app.cargamaquina.com.br/fiscal/nfe/saida")
        driver.implicitly_wait(10)
        time.sleep(4)
        selenium_cookies = driver.get_cookies()
        request_cookies = {
            cookie["name"]: cookie["value"] for cookie in selenium_cookies
        }
        scripts = driver.find_elements("tag name", "script")
        csrf_token: str = ""
        for script in scripts:
            content = script.get_attribute("innerHTML")
            if content and "YII_CSRF_TOKEN" in content:
                match = re.search(r"window\.YII_CSRF_TOKEN\s*=\s*'([^']+)'", content)
                if match:
                    csrf_token = match.group(1)
                    break

        print(f"Token CSRF encontrado: {csrf_token}")
        time.sleep(3)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(3)
    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.ElementNotInteractableException as e:
        print(f"Error: {e}")

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
            "[bold blue]Digite o número do tipo/proprietário que deseja criar:[/bold blue]"
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
            emails: list[str] = [
                "vera.cristina@lanxcables.com.br",
                "adriana.damas@lanxcables.com.br",
                "renata.perez@lanxcables.com.br",
                "lucineide@lanxcables.com.br",
            ]
            # Create orders
            for i, order in enumerate(orders[:]):
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
                    pygui.write(
                        "f&k group tecnologia em sistemas automotivos ltda",
                    )
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
                        pygui.write(code)
                        time.sleep(5)
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
                        pygui.write("30/06/2025")
                        wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='adicionarItemCompraGrid']")
                            )
                        ).click()
                    create_invoice_request(order, cookies=request_cookies, crsf_token=csrf_token)

                    current_email = emails[i % len(emails)]
                    send_order_email(order, rateio, current_email)
                    time.sleep(4)

                    orders.remove(order)
                    with open(json_file_path, "w", encoding="utf-8") as f:
                        json.dump(orders, f, indent=4, ensure_ascii=False)
                except selenium.common.exceptions.NoSuchElementException as e:
                    print(f"Error: {e}")
                    input("Press Enter to continue...")
                    return
                except selenium.common.exceptions.ElementNotInteractableException as e:
                    print(f"Error: {e}")
                    input("Press Enter to continue...")
                    return
                except selenium.common.exceptions.StaleElementReferenceException as e:
                    print(f"Error: {e}")
                    input("Press Enter to continue...")
                    return
                except Exception as e:
                    print(f"Error: {e}")
                    input("Press Enter to continue...")
                    return
            driver.quit()
            time.sleep(6)

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
