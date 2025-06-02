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

from core.utils.send_email import send_order_email


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
        driver.switch_to.window(driver.window_handles[0])
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

                    # Create Invoice
                    driver.switch_to.window(driver.window_handles[1])
                    time.sleep(4)
                    wait.until(
                        EC.element_to_be_clickable((By.LINK_TEXT, "Incluir"))
                    ).click()
                    time.sleep(4)
                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='s2id_sel2Natureza']")
                        )
                    ).click()

                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='select2-drop']/div/input")
                        )
                    ).send_keys("5.949-")
                    pygui.press("enter")
                    time.sleep(4)
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
                    time.sleep(4)
                    console.print(
                        "[bold green]Destinatário incluido com Sucesso[/bold green]"
                    )

                    comp = wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='txtComplemento']")
                        )
                    )
                    driver.execute_script("arguments[0].scrollIntoView();", comp)
                    time.sleep(4)

                    for item in order["itens"]:
                        time.sleep(4)
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
                        time.sleep(4)

                        wait.until(
                            EC.element_to_be_clickable(
                                (
                                    By.XPATH,
                                    "//*[@id='tabDados']/div/div[1]/div[2]/div[3]/a",
                                )
                            )
                        ).click()
                        time.sleep(4)
                        pygui.press("tab", presses=2, interval=0.2)
                        code = str(item["CODÍGO"])
                        pygui.write(code)
                        time.sleep(4)
                        pygui.press("enter")
                        time.sleep(4)

                        wait.until(
                            EC.visibility_of_element_located(
                                (
                                    By.XPATH,
                                    "//*[@id='dialogSelecionarProduto']/div/div/div[3]/a[1]",
                                )
                            )
                        ).click()
                        time.sleep(4)
                        console.print(
                            f"[bold green]Item: {item['CODÍGO']} incluido com Sucesso[/bold green]"
                        )
                        time.sleep(4)

                        ncm = wait.until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//*[@id='s2id_txtNCM']/a/span[1]")
                            )
                        )
                        while ncm.text == "":
                            if ncm.text == "":
                                pygui.shortcut("alt", "tab")
                                console.print(
                                    f"[bold red]NCM não encontrado verifique o código do material: {code}[/bold red]",
                                )
                                input(
                                    "Após adicionar o código manualmente, volte aqui e pressione enter para continuar..."
                                )

                                time.sleep(4)
                                pygui.shortcut("alt", "tab")
                                time.sleep(4)

                        qty = item["SALDO TOTAL"]
                        if float(qty).is_integer():
                            qty = str(int(qty))
                        else:
                            qty = str(qty).replace(".", ",")

                        pygui.write(qty)
                        item["SALDO TOTAL"] = qty

                        wait.until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//*[@id='quantidadeCom']")
                            )
                        ).send_keys(qty)

                        cost: str = ""
                        if item["CUSTO UNITARIO"] == 0:
                            cost = "1,00"
                        elif item["CUSTO UNITARIO"] < 0:
                            cost = str(abs(item["CUSTO UNITARIO"])).replace(".", ",")
                        else:
                            cost = f"{float(item['CUSTO UNITARIO']):.3f}".replace(
                                ".", ","
                            )

                        wait.until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//*[@id='valorUnitario']")
                            )
                        ).send_keys(cost)
                        console.print(
                            "[bold green]Quantidade e Custo UN incluido com Sucesso[/bold green]"
                        )

                        time.sleep(4)
                        wait.until(
                            EC.element_to_be_clickable(
                                (
                                    By.XPATH,
                                    "//*[@id='modal-item-form']/div/div[2]/div/div/ul/li[2]/a",
                                )
                            )
                        ).click()
                        time.sleep(4)
                        wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='accordionTributosICMS']")
                            )
                        ).click()

                        pygui.press("tab", presses=2, interval=0.2)
                        pygui.write("0 - Nacional")
                        time.sleep(4)
                        pygui.press("enter")
                        time.sleep(4)

                        add_item = wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//*[@id='btnGravarItem']")
                            )
                        )
                        add_item.click()
                        console.print(
                            "[bold green]ICMS Incluido com Sucesso[/bold green]"
                        )
                        time.sleep(4)

                    time.sleep(4)
                    pygui.shortcut("ctrl", "end")

                    frete = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='s2id_sel2ModalidadeFrete']")
                        )
                    )
                    driver.execute_script("arguments[0].scrollIntoView();", frete)
                    time.sleep(4)
                    frete.click()
                    pygui.write("Sem frete", interval=0.1)
                    time.sleep(4)
                    pygui.press("enter")
                    time.sleep(4)

                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='s2id_sel2PrazoPR']")
                        )
                    ).click()

                    time.sleep(4)
                    pygui.write("30 dias")
                    time.sleep(4)
                    pygui.press("enter")

                    time.sleep(4)

                    select_element = wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='formaPag']")
                        )
                    )
                    select_element.click()
                    time.sleep(4)
                    pygui.write("Bolet")
                    time.sleep(4)
                    pygui.press("enter")
                    time.sleep(4)

                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='gerarParcelas']")
                        )
                    ).click()
                    console.print(
                        "[bold green]Condição de Pagamento incluida com sucesso[/bold green]"
                    )
                    time.sleep(4)
                    order_info = f"Pedido de Compra: {order['pedido_numero']}"
                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='infoAdicionais']")
                        )
                    ).send_keys(order_info)
                    time.sleep(4)
                    pygui.shortcut("ctrl", "end")
                    time.sleep(4)

                    # Style Grid and Button
                    generate_invoice_grid = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='mainGridControl']/div/div[1]/div")
                        )
                    )
                    driver.execute_script(
                        "arguments[0].style.padding = '50px';", generate_invoice_grid
                    )
                    time.sleep(4)
                    btn_style = (
                        "height: 80px;"
                        "text-align: center;"
                        "display: flex;"
                        "justify-content: center;"
                        "align-items: center;"
                        "gap: 2px;"
                    )
                    time.sleep(4)
                    wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='btnGerarMais']")
                        )
                    ).click()
                    pygui.shortcut("ctrl", "end")
                    time.sleep(4)
                    generate_invoice_btn = wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "//*[@id='divBtnGerar']/ul/li/a")
                        )
                    )
                    driver.execute_script(
                        f"arguments[0].setAttribute('style', '{btn_style}');",
                        generate_invoice_btn,
                    )
                    time.sleep(4)
                    generate_invoice_btn.click()
                    console.print(
                        "[bold green]Nota Fiscal criada com sucesso![/bold green]"
                    )
                    current_email = emails[i % len(emails)]
                    send_order_email(order, rateio, current_email)
                    time.sleep(4)
                    driver.switch_to.window(driver.window_handles[0])
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
            console.print("[bold blue]Iniciando criação das notas fiscais[/bold blue]")

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
