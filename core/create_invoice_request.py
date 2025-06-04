"""
module to create an invoice request for a given order.
This module defines a function to send a POST request to the Cargamaquina
to create an invoice based on the provided order data and authentication cookies.
"""

import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import selenium.common.exceptions


def get_cookies(username: str, password: str) -> dict:
    """
    Retrieve cookies and CSRF token from the Cargamaquina login page.
    :param username: Username for Cargamaquina login.
    :param password: Password for Cargamaquina login.
    :return: A tuple containing cookies and CSRF token."""
    try:
        driver = webdriver.Chrome()
        options = Options()
        options.add_argument("--headless")  # Run in headless mode
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options=options)
        driver.get("https://app.cargamaquina.com.br/site/login/c/31.1~78,8%5E56,8")
        driver.find_element(By.ID, "LoginForm_username").send_keys(username)
        driver.find_element(By.ID, "LoginForm_password").send_keys(password)
        driver.find_element(By.ID, "LoginForm_submit").click()
        selenium_cookies = driver.get_cookies()
        request_cookies: dict = {cookie["name"]: cookie["value"] for cookie in selenium_cookies}
    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Element not found: {e}")
        return {}
    except selenium.common.exceptions.TimeoutException as e:
        print(f"Timeout while trying to find elements: {e}")
        return {}
    except selenium.common.exceptions.WebDriverException as e:
        print(f"WebDriver error: {e}")
        return {}
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return {}
    finally:
        driver.quit()
    print("Cookies and CSRF token retrieved successfully.")
    return request_cookies


def create_request_item(index: int, item: dict) -> tuple[dict[str, str], float]:
    """
    Format the request data for creating an invoice.
    This function is a placeholder and should be implemented to return the actual request data.
    """
    try:
        item_number: int = index + 1
        code: str = item["CODÍGO"]
        description: str = item["DESCRIÇÃO"][:120]
        unit: str = item["UNIDADE"]
        qty = item["SALDO TOTAL"]
        unit_price: str = f"{item['CUSTO UNITARIO']:.3f}"
        ncm = item["NCM"]
        item_total_price: str = f"{qty * float(unit_price):.2f}"
        pis_price: str = f"{float(item_total_price) * 0.0165:.2f}"
        cofins_price: str = f"{float(item_total_price) * 0.076:.2f}"

        if float(qty).is_integer():
            qty = str(int(qty))
        else:
            qty = str(qty)

        item_request: dict[str, str] = {
            f"ItensNota[{index}][numero_item]": f"{item_number}",
            f"ItensNota[{index}][codigo_produto]": f"{code}",
            f"ItensNota[{index}][descricao]": f"{description}",
            f"ItensNota[{index}][unidade_comercial]": f"{unit}",
            f"ItensNota[{index}][cfop]": "4660576",
            f"ItensNota[{index}][quantidade_comercial]": f"{qty}",
            f"ItensNota[{index}][valor_unitario_comercial]": f"{unit_price}",
            f"ItensNota[{index}][valor_bruto]": f"{item_total_price}",
            f"ItensNota[{index}][codigo_ncm]": f"{ncm}",
            f"ItensNota[{index}][tipi_id]": f"{ncm}",
            f"ItensNota[{index}][cest]": "",
            f"ItensNota[{index}][unidade_tributavel]": "",
            f"ItensNota[{index}][quantidade_tributavel]": "",
            f"ItensNota[{index}][valor_unitario_tributavel]": "",
            f"ItensNota[{index}][valor_frete]": "",
            f"ItensNota[{index}][valor_seguro]": "",
            f"ItensNota[{index}][valor_desconto]": "",
            f"ItensNota[{index}][valor_outras_despesas]": "",
            f"ItensNota[{index}][pedido_compra]": "",
            f"ItensNota[{index}][numero_item_pedido_compra]": "",
            f"ItensNota[{index}][numero_fci]": "",
            f"ItensNota[{index}][icms_situacao_tributaria]": "41",
            f"ItensNota[{index}][icms_origem]": "0",
            f"ItensNota[{index}][fcp_percentual]": "",
            f"ItensNota[{index}][fcp_base_calculo]": "",
            f"ItensNota[{index}][fcp_valor]": "",
            f"ItensNota[{index}][icms_modalidade_base_calculo]": "0",
            f"ItensNota[{index}][ipi_situacao_tributaria]": "51",
            f"ItensNota[{index}][ipi_base_calculo]": "",
            f"ItensNota[{index}][ipi_valor]": "0",
            f"ItensNota[{index}][ipi_aliquota]": "",
            f"ItensNota[{index}][ipi_codigo_enquadramento_legal]": "999",
            f"ItensNota[{index}][icms_aliquota]": "",
            f"ItensNota[{index}][icms_base_calculo]": "",
            f"ItensNota[{index}][icms_valor]": "0",
            f"ItensNota[{index}][icms_reducao_base_calculo]": "",
            f"ItensNota[{index}][icms_percentual_diferimento]": "",
            f"ItensNota[{index}][icms_valor_operacao]": "",
            f"ItensNota[{index}][icms_valor_diferido]": "",
            f"ItensNota[{index}][pis_situacao_tributaria]": "01",
            f"ItensNota[{index}][pis_base_calculo]": f'{item_total_price.replace(".", ",")}',
            f"ItensNota[{index}][pis_aliquota_valor]": "",
            f"ItensNota[{index}][pis_aliquota_porcentual]": "1,65",
            f"ItensNota[{index}][pis_valor]": f'{pis_price.replace(".", ",")}',
            f"ItensNota[{index}][pis_quantidade_vendida]": "",
            f"ItensNota[{index}][cofins_situacao_tributaria]": "01",
            f"ItensNota[{index}][cofins_base_calculo]": f'{item_total_price.replace(".", ",")}',
            f"ItensNota[{index}][cofins_aliquota_valor]": "",
            f"ItensNota[{index}][cofins_aliquota_porcentual]": "7,6",
            f"ItensNota[{index}][cofins_valor]": f'{cofins_price.replace(".", ",")}',
            f"ItensNota[{index}][cofins_quantidade_vendida]": "",
            f"ItensNota[{index}][issqn_base_calculo]": "",
            f"ItensNota[{index}][issqn_aliquota]": "",
            f"ItensNota[{index}][issqn_valor]": "",
            f"ItensNota[{index}][ii_base_calculo]": f'{item_total_price.replace(".", ",")}',
            f"ItensNota[{index}][ii_aliquota]": "",
            f"ItensNota[{index}][ii_valor]": "",
            f"ItensNota[{index}][ii_despesas_aduaneiras]": "",
            f"ItensNota[{index}][ii_valor_iof]": "",
            f"ItensNota[{index}][informacoes_adicionais_item]": "",
            f"ItensNota[{index}][_produtoId]": "",
            f"ItensNota[{index}][_itemVendaId]": "",
            f"ItensNota[{index}][rastros]": "",
            f"ItensNota[{index}][documentos_importacao]": "",
            f"ItensNota[{index}][icms_aliquota_credito_simples]": "",
            f"ItensNota[{index}][icms_valor_credito_simples]": "",
            f"ItensNota[{index}][icms_modalidade_base_calculo_st]": "6",
            f"ItensNota[{index}][icms_reducao_base_calculo_st]": "",
            f"ItensNota[{index}][icms_margem_valor_adicionado_st]": "",
            f"ItensNota[{index}][icms_base_calculo_st]": "",
            f"ItensNota[{index}][icms_aliquota_st]": "",
            f"ItensNota[{index}][icms_valor_st]": "",
            f"ItensNota[{index}][icms_base_calculo_retido_remetente]": "",
            f"ItensNota[{index}][icms_valor_retido_remetente]": "",
            f"ItensNota[{index}][icms_valor_desonerado]": "",
            f"ItensNota[{index}][icms_motivo_desoneracao]": "0",
            f"ItensNota[{index}][icms_deducao_desoneracao]": "0",
            f"ItensNota[{index}][codigo_beneficio_fiscal]": "PR800006",
            f"ItensNota[{index}][reducao_aliquota_icms]": "",
            f"ItensNota[{index}][inclui_no_total]": "1",
            f"ItensNota[{index}][item_retorno_xml]": "",
            f"ItensNota[{index}][icms_base_calculo_uf_destino]": "",
            f"ItensNota[{index}][fcp_base_calculo_uf_destino]": "",
            f"ItensNota[{index}][fcp_percentual_uf_destino]": "",
            f"ItensNota[{index}][fcp_base_calculo_st]": "",
            f"ItensNota[{index}][fcp_percentual_st]": "",
            f"ItensNota[{index}][fcp_valor_st]": "",
            f"ItensNota[{index}][icms_aliquota_interna_uf_destino]": "",
            f"ItensNota[{index}][icms_aliquota_interestadual]": "",
            f"ItensNota[{index}][icms_percentual_partilha]": "",
            f"ItensNota[{index}][fcp_valor_uf_destino]": "",
            f"ItensNota[{index}][icms_valor_uf_destino]": "",
            f"ItensNota[{index}][icms_valor_uf_remetente]": "",
            f"ItensNota[{index}][icms_base_calculo_retido_st]": "",
            f"ItensNota[{index}][icms_aliquota_final]": "",
            f"ItensNota[{index}][icms_valor_substituto]": "",
            f"ItensNota[{index}][icms_valor_retido_st]": "",
            f"ItensNota[{index}][icms_reducao_base_calculo_efetiva]": "",
            f"ItensNota[{index}][icms_base_calculo_efetiva]": "",
            f"ItensNota[{index}][icms_aliquota_efetiva]": "",
            f"ItensNota[{index}][icms_valor_efetivo]": "",
            f"ItensNota[{index}][total_tributo_federal]": "",
            f"ItensNota[{index}][total_tributo_estadual]": "",
            f"ItensNota[{index}][total_tributo_municipal]": "",
            f"ItensNota[{index}][combustivel_codigo_anp]": "",
            f"ItensNota[{index}][combustivel_descricao_anp]": "",
            f"ItensNota[{index}][combustivel_sigla_uf]": "",
            f"ItensNota[{index}][combustivel_percentual_glp]": "",
            f"ItensNota[{index}][combustivel_percentual_biodiesel]": "",
            f"ItensNota[{index}][combustivel_percentual_gas_natural_nacional]": "",
            f"ItensNota[{index}][combustivel_valor_partida]": "",
            f"ItensNota[{index}][combustivel_registro_codif]": "",
            f"ItensNota[{index}][combustivel_percentual_gas_natural_importado]": "",
            f"ItensNota[{index}][quantidade_tributada_combustivel]": "",
            f"ItensNota[{index}][aliquota_ad_rem_combustivel]": "",
            f"ItensNota[{index}][valor_icms_proprio_combustivel]": "",
            f"ItensNota[{index}][quantidade_tributada_sujeita_retencao_combustivel]": "",
            f"ItensNota[{index}][aliquota_ad_rem_retencao_combustivel]": "",
            f"ItensNota[{index}][valor_icms_retencao_combustivel]": "",
            f"ItensNota[{index}][valor_icms_operacao_combustivel]": "",
            f"ItensNota[{index}][percentucal_diferimento_combustivel]": "",
            f"ItensNota[{index}][valor_icms_diferido_combustivel]": "",
            f"ItensNota[{index}][quantidade_tributada_retida_anteriormente_combustivel]": "",
            f"ItensNota[{index}][aliquota_ad_rem_retido_anteriormente_combustivel]": "",
            f"ItensNota[{index}][valor_imcs_retido_anteriormente_combustivel]": "",
            f"ItensNota[{index}][icms_base_calculo_mono_retido]": "",
            f"ItensNota[{index}][icms_aliquota_retido]": "",
            f"ItensNota[{index}][icms_valor_mono_retido]": "",
            f"ItensNota[{index}][codigo_remessa_xml]": "",
            f"ItensNota[{index}][nota_remessa_id]": "",
        }
    except KeyError as e:
        print(f"KeyError: {e} in item {index} with data {item}")
        item_request = {}
    except ValueError as e:
        print(f"ValueError: {e} in item {index} with data {item}")
        item_request = {}
    except TypeError as e:
        print(f"TypeError: {e} in item {index} with data {item}")
        item_request = {}
    except Exception as e:
        print(f"Unexpected error: {e} in item {index} with data {item}")
        item_request = {}

    sum_total_price = float(item_total_price.replace(",", "."))

    return item_request, sum_total_price


def create_invoice_request(order: dict, cookies: dict, crsf_token: str):
    """
    Create an invoice request for a given order.

    :param order_data: Dictionary containing order details.
    :param cookie: Cookie string for authentication.
    """
    url: str = "https://app.cargamaquina.com.br/speed/speedNfe/saida"
    all_items: dict = {}
    data_request: dict = {}
    total_nfe_price = 0.0
    for i, item in enumerate(order["itens"]):
        item_request, item_price = create_request_item(i, item)
        all_items.update(item_request)
        total_nfe_price += item_price
    total_nfe_price = f"{total_nfe_price:.2f}"
    data = {
        "YII_CSRF_TOKEN": f"{crsf_token}",
        "Nfe[_fonteTributacao]": "",
        "Nfe[_artigoTributacao]": "",
        "Nfe[armazenar]": "1",
        "Nfe[_origemNota]": "",
        "Nfe[identificador_unico]": "9364683e370b6ff160.27403748",
        "Nfe[tipo_operacao]": "1",
        "Nfe[natureza_id]": "37310453",
        "Nfe[finalidade_emissao]": "1",
        "Nfe[consumidor_final]": "0",
        "Nfe[destino_operacao]": "1",
        "Nfe[consumidor_final_presente]": "0",
        "Nfe[chave_referenciada]": "",
        "Nfe[uf_local_embarque]": "",
        "Nfe[local_embarque]": "",
        "Nfe[emitente_id]": "6676429",
        "Nfe[emitenteCNPJ]": "09.652.475/0001-37",
        "Nfe[emitenteRazaoSocial]": "F&K Group Tecnologia em Sistemas Automotivos Ltda",
        "Nfe[emitenteNomeFantasia]": "F&K Group",
        "Nfe[emitenteInscricaoEstadual]": "9045253241",
        "Nfe[emitenteInscricaoEstadualST]": "",
        "Nfe[emitenteInscricaoMunicipal]": "",
        "Nfe[emitenteCodigoRegimeTributario]": "3",
        "Nfe[emitenteEnderecoCEP]": "86050450",
        "Nfe[emitenteEnderecoLogradouro]": "R JOAO WYCLIF",
        "Nfe[emitenteEnderecoNumero]": "111",
        "Nfe[emitenteEnderecoBairro]": "GLEBA FAZENDA PALHANO",
        "Nfe[estadoCidadeEmitenteId]": "3446",
        "Nfe[cidadeEmitenteId]": "6838",
        "Nfe[emitenteTelefone]": "(43) 3032-5556",
        "Nfe[tipo_destinatario]": [
            "",
            "N",
        ],
        "Nfe[destinatario_id]": "37289422",
        "tituloBack": "Destinatário",
        "[hidenExceto]": "0",
        "Nfe[tipo_documento_destinatario]": "0",
        "Nfe[cnpj_destinatario]": "60.347.923/0001-46",
        "Nfe[cpf_destinatario]": "",
        "Nfe[razao_social_destinatario]": "LANX CABLES LTDA",
        "Nfe[tipo_inscricao_estadual_destinatario]": "1",
        "Nfe[inscricao_estadual_destinatario]": "91141047-84",
        "Nfe[inscricao_municipal_destinatario]": "",
        "Nfe[inscricao_suframa_destinatario]": "",
        "Nfe[email_destinatario]": "",
        "Nfe[cep_destinatario]": "86065000",
        "Nfe[logradouro_destinatario]": "ARTHUR THOMAS",
        "Nfe[numero_destinatario]": "1795",
        "Nfe[bairro_destinatario]": "RODOCENTRO",
        "Nfe[pais_destinatario]": "",
        "Nfe[uf_destinatario]": "3446",
        "Nfe[municipio_destinatario]": "6838",
        "Nfe[telefone_destinatario]": "433032-4292",
        "Nfe[complemento_destinatario]": "",
        "Nfe[cnpj_entrega]": "",
        "Nfe[cpf_entrega]": "",
        "Nfe[inscricao_estadual_entrega]": "",
        "Nfe[nome_entrega]": "",
        "Nfe[email_entrega]": "",
        "Nfe[cep_entrega]": "",
        "Nfe[logradouro_entrega]": "",
        "Nfe[numero_entrega]": "",
        "Nfe[bairro_entrega]": "",
        "Nfe[pais_entrega]": "",
        "Nfe[uf_entrega]": "",
        "Nfe[telefone_entrega]": "",
        "Nfe[complemento_entrega]": "",
        "Nfe[_totalNotaImportacao]": "",
        "processo": "",
        "infoItem": "G",
        "Nfe[transportadora_id]": "",
        "Nfe[modalidade_frete]": "9",
        "Nfe[valorFrete]": "",
        "Nfe[transportadorTipPessoa]": "0",
        "Nfe[cnpj_transportador]": "",
        "Nfe[cpf_transportador]": "",
        "Nfe[inscricao_estadual_transportador]": "",
        "Nfe[endereco_transportador]": "",
        "Nfe[uf_transportador]": "",
        "Nfe[municipio_transportador]": "",
        "Nfe[valor_desconto]": "",
        "Nfe[veiculo_placa]": "",
        "Nfe[veiculo_uf]": "",
        "Nfe[veiculo_rntc]": "",
        "Nfe[veiculo_identificacao_vagao]": "",
        "Nfe[veiculo_identificacao_balsa]": "",
        "Nfe[volume_quantidade]": "",
        "Nfe[volume_especie]": "",
        "Nfe[volume_marca]": "",
        "Nfe[volume_numero]": "",
        "Nfe[volume_peso_liquido]": "",
        "Nfe[volume_peso_bruto]": "",
        "Nfe[condicao_pagamento_id]": "6091530",
        "condicao": "30",
        "Nfe[forma_pagamento]": "15",
        "Nfe[bandeira_pagamento]": "01",
        "Nfe[total_nota]": f"{total_nfe_price.replace('.', ',')}",
        "Nfe[_totalIpi]": "0",
        "PagamentosNota[0][numero]": "1",
        "PagamentosNota[0][vencimento]": "02/07/2025",
        "PagamentosNota[0][valor]": f"{total_nfe_price.replace('.', ',')}",
        "PagamentosNota[0][forma]": "15",
        "PagamentosNota[0][_bandeira]": "01",
        "PagamentosNota[0][_aVista]": "1",
        "Nfe[info_adicionais]": f"{order['pedido_numero']}",
    }

    for k, v in data.items():
        if k == "infoItem":
            data_request[k] = v
            data_request.update(all_items)
        else:
            data_request[k] = v

    try:
        response = requests.post(url, cookies=cookies, data=data, timeout=15)
        if not response.ok:
            print(f"Failed to create invoice request: {response.status_code}")
            return
        print("Invoice request created successfully.")
        print("Response:", response.status_code, response.text)
    except requests.RequestException as e:
        print(f"An error occurred while creating the invoice request: {e}")


if __name__ == "__main__":
    cookies, csrf_token = get_cookies("username", "password")
    if not cookies or not csrf_token:
        print("Failed to retrieve cookies or CSRF token.")
    create_invoice_request(
        order={
            "pedido_numero": "123456789",
            "itens": [
                {
                    "CODÍGO": "123456",
                    "DESCRIÇÃO": "Produto de Teste",
                    "UNIDADE": "UN",
                    "SALDO TOTAL": 10,
                    "CUSTO UNITARIO": 25.99,
                    "NCM": "12345678",
                },
                {
                    "CODÍGO": "654321",
                    "DESCRIÇÃO": "Outro Produto de Teste",
                    "UNIDADE": "KG",
                    "SALDO TOTAL": 5,
                    "CUSTO UNITARIO": 15.50,
                    "NCM": "87654321",
                },
            ],
        },
        cookies=cookies,
        crsf_token=csrf_token
    )
