"""
This module is responsible for creating invoices on the website.
It uses Selenium to automate the login process and navigate through the website.
"""

import time
import pyautogui as pygui
from rich.console import Console
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import selenium.common.exceptions

from core.utils.send_email import send_order_email


def create_invoice(
    username: str, password: str, orders: list[dict], rateio: str
) -> None:
    """Create invoice"""
    # Login
    login_url: str = "https://app.cargamaquina.com.br/site/login?c=31.1~78%2C8%5E56%2C8"
    options = Options()
    options.add_argument("--force-device-scale-factor=0.50")
    d = webdriver.Chrome(options=options)
    d.maximize_window()
    wait = WebDriverWait(d, 15)
    console = Console()
    try:
        d.get(login_url)
        d.implicitly_wait(15)
        d.find_element(By.ID, "LoginForm_username").send_keys(username)
        d.find_element(By.ID, "LoginForm_password").send_keys(password)
        d.find_element(By.NAME, "yt0").click()
        console.print("[bold green]Login na F&K realizado com sucesso![/bold green]")
        d.implicitly_wait(15)
    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.ElementNotInteractableException as e:
        print(f"Error: {e}")

    # Emails
    emails: list[str] = [
        "vera.cristina@lanxcables.com.br",
        "adriana.damas@lanxcables.com.br",
        "lucineide@lanxcables.com.br",
    ]

    # Create
    time.sleep(6)
    try:
        d.get("https://app.cargamaquina.com.br/fiscal/nfe/saida")
        time.sleep(6)
        for i, order in enumerate(orders):
            time.sleep(3)
            wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Incluir"))).click()
            time.sleep(3)
            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='s2id_sel2Natureza']")
                )
            ).click()

            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='select2-drop']/div/input")
                )
            ).send_keys("5.102")
            pygui.press("enter")
            time.sleep(2)
            console.print(
                "[bold green]Natureza da Operação incluido com Sucesso[/bold green]"
            )

            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='s2id_sel2PessoaDestinatario']")
                )
            ).click()
            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='select2-drop']/div/input")
                )
            ).send_keys("60.347.923/0001-46")
            pygui.press("enter")
            time.sleep(2)
            console.print("[bold green]Destinatário incluido com Sucesso[/bold green]")

            comp = wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='txtComplemento']")
                )
            )
            d.execute_script("arguments[0].scrollIntoView();", comp)
            time.sleep(3)

            for item in order["itens"]:
                time.sleep(3)
                console.print(
                    "[bold blue]Abrindo Formulário para Adicionar Itens na NF[/bold blue]"
                )
                add_form = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//*[@id='collapseItens']/div/div/div/div[2]/div[1]/div[1]/div[1]/a",
                        )
                    )
                )

                add_form.click()
                time.sleep(3)

                wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//*[@id='tabDados']/div/div[1]/div[2]/div[3]/a")
                    )
                ).click()
                time.sleep(3)
                pygui.press("tab", presses=2, interval=0.2)
                pygui.write(item["CODÍGO"])
                time.sleep(2)
                pygui.press("enter")
                time.sleep(3)

                wait.until(
                    EC.visibility_of_element_located(
                        (
                            By.XPATH,
                            "//*[@id='dialogSelecionarProduto']/div/div/div[3]/a[1]",
                        )
                    )
                ).click()
                time.sleep(3)
                console.print(
                    f"[bold green]Item: {item['CODÍGO']} incluido com Sucesso[/bold green]"
                )

                wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@id='quantidadeCom']")
                    )
                ).send_keys(item["SALDO TOTAL"])

                cost: str = ""
                if item["CUSTO UNITARIO"] == 0:
                    cost = "1,00"
                elif item["CUSTO UNITARIO"] < 0:
                    cost = str(abs(item["CUSTO UNITARIO"])).replace(".", ",")
                else:
                    cost = str(item["CUSTO UNITARIO"]).replace(".", ",")

                wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@id='valorUnitario']")
                    )
                ).send_keys(cost)
                console.print(
                    "[bold green]Quantidade e Custo UN incluido com Sucesso[/bold green]"
                )

                time.sleep(3)
                wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//*[@id='modal-item-form']/div/div[2]/div/div/ul/li[2]/a",
                        )
                    )
                ).click()
                time.sleep(3)
                wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//*[@id='accordionTributosICMS']")
                    )
                ).click()

                pygui.press("tab", presses=2, interval=0.2)
                pygui.write("0 - Nacional")
                time.sleep(3)
                pygui.press("enter")
                time.sleep(3)

                icms_input = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@id='icmsAliquota']")
                    )
                )
                icms_value = icms_input.get_attribute("value")
                if icms_value is not None:
                    icms_value = icms_value.replace(".", ",")
                    icms_input.clear()
                    time.sleep(2)
                    icms_input.send_keys(icms_value)
                    time.sleep(2)

                add_item = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@id='btnGravarItem']"))
                )
                add_item.click()
                console.print("[bold green]ICMS Incluido com Sucesso[/bold green]")
                time.sleep(3)

            time.sleep(3)
            pygui.shortcut("ctrl", "end")

            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='s2id_sel2PrazoPR']")
                )
            ).click()

            time.sleep(2)
            pygui.write("30 dias")
            time.sleep(2)
            pygui.press("enter")

            time.sleep(3)

            select_element = wait.until(
                EC.visibility_of_element_located((By.XPATH, "//*[@id='formaPag']"))
            )
            select_element.click()
            time.sleep(3)
            pygui.write("Bolet")
            time.sleep(3)
            pygui.press("enter")
            time.sleep(2)

            wait.until(
                EC.visibility_of_element_located((By.XPATH, "//*[@id='gerarParcelas']"))
            ).click()
            console.print(
                "[bold green]Condição de Pagamento incluida com sucesso[/bold green]"
            )
            time.sleep(3)
            order_info = f"Pedido de Compra: {order['pedido_numero']}"
            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='infoAdicionais']")
                )
            ).send_keys(order_info)
            time.sleep(3)
            pygui.shortcut("ctrl", "end")
            time.sleep(2)

            # Style Grid and Button
            generate_invoice_grid = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='mainGridControl']/div/div[1]/div")
                )
            )
            d.execute_script(
                "arguments[0].style.padding = '50px';", generate_invoice_grid
            )
            time.sleep(3)
            btn_style = (
                "height: 80px;"
                "text-align: center;"
                "display: flex;"
                "justify-content: center;"
                "align-items: center;"
                "gap: 2px;"
            )
            time.sleep(3)
            wait.until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='btnGerarMais']"))
            ).click()
            pygui.shortcut("ctrl", "end")
            time.sleep(3)
            generate_invoice_btn = wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='divBtnGerar']/ul/li/a")
                )
            )
            d.execute_script(
                f"arguments[0].setAttribute('style', '{btn_style}');",
                generate_invoice_btn,
            )
            generate_invoice_btn.click()
            console.print("[bold green]Nota Fiscal criada com sucesso![/bold green]")
            current_email = emails[i % len(emails)]
            send_order_email(order, rateio, current_email)
            time.sleep(2)
    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.ElementNotInteractableException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.StaleElementReferenceException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.TimeoutException as e:
        print(f"Error: {e}")
    except selenium.common.exceptions.InvalidSessionIdException as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Error: {e}")
        input("Press Enter to continue...")
    finally:
        d.quit()


if __name__ == "__main__":
    orders_data = [
        {
            "tipo_proprietario": "Embalagem / Estoque",
            "pedido_numero": 1,
            "quantidade_itens": 2,
            "itens": [
                {
                    "CODÍGO": "CXPP01GKF",
                    "FINALIDADE": "Embalagens",
                    "PROPRIETARIO": "Estoque",
                    "SALDO TOTAL": 165,
                    "UNIDADE": "un",
                    "CUSTO UNITARIO": 0.80625,
                    "DECLARA": "DECLARA",
                    "NCM": 363,
                },
                {
                    "CODÍGO": "CXPP01M",
                    "FINALIDADE": "Embalagens",
                    "PROPRIETARIO": "Estoque",
                    "SALDO TOTAL": 100,
                    "UNIDADE": "un",
                    "CUSTO UNITARIO": 0.80625,
                    "DECLARA": "DECLARA",
                    "NCM": 363,
                },
            ],
        }
    ]

    create_invoice("username", "password", orders_data, "EMBALAGEM")
