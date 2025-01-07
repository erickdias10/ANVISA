import logging
import nest_asyncio
import time
import getpass
import os
import unicodedata
import re
import spacy

from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from webdriver_manager.chrome import ChromeDriverManager
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt

# Configuração de logs
logging.basicConfig(level=logging.INFO)

# Aplicação do nest_asyncio para permitir múltiplos loops de eventos (necessário se for rodar em notebook)
nest_asyncio.apply()

# Constantes de elementos
LOGIN_URL = "https://sei.anvisa.gov.br/sip/login.php?sigla_orgao_sistema=ANVISA&sigla_sistema=SEI"
IFRAME_VISUALIZACAO_ID = "ifrVisualizacao"
BUTTON_XPATH_ALT = '//img[@title="Gerar Arquivo PDF do Processo"]/parent::a'


def create_driver(download_dir=None):
    """
    Configura e retorna uma instância do Selenium WebDriver,
    forçando o download de PDF ao invés de abrir no Chrome.
    """
    if download_dir is None:
        # Diretório padrão de downloads
        download_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-notifications")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def wait_for_element(driver, by, value, timeout=20):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return element
    except Exception as e:
        logging.error(f"Erro ao localizar o elemento {value}: {e}")
        raise


def handle_alert(driver):
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = Alert(driver)
        alert_text = alert.text
        alert.accept()
        return alert_text
    except TimeoutException:
        return None


def login(driver, username, password):
    """
    Realiza o login no sistema SEI.
    """
    driver.get(LOGIN_URL)
    wait_for_element(driver, By.ID, "txtUsuario").send_keys(username)
    driver.find_element(By.ID, "pwdSenha").send_keys(password)
    driver.find_element(By.ID, "sbmAcessar").click()


def access_process(driver, process_number):
    """
    Acessa um processo pelo número no sistema SEI.
    """
    search_field = wait_for_element(driver, By.ID, "txtPesquisaRapida")
    search_field.send_keys(process_number)
    search_field.send_keys("\n")
    time.sleep(3)


def generate_pdf(driver):
    """
    Gera o PDF do processo no iframe correspondente.
    """
    try:
        driver.switch_to.frame(wait_for_element(driver, By.ID, IFRAME_VISUALIZACAO_ID))
        gerar_pdf_button = wait_for_element(driver, By.XPATH, BUTTON_XPATH_ALT)
        driver.execute_script("arguments[0].click();", gerar_pdf_button)
    except Exception as e:
        logging.error(f"Erro ao gerar o PDF: {e}")
        raise
    finally:
        driver.switch_to.default_content()
        time.sleep(5)


def download_pdf(driver, option="Todos os documentos disponíveis"):
    """
    Realiza o clique no botão 'Gerar Arquivo PDF do Processo' e seleciona a opção desejada.
    """
    try:
        driver.switch_to.default_content()
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="menu-opcao"]//button'))
        )
        for option_button in dropdown_options:
            if option_button.text.strip() == option:
                driver.execute_script("arguments[0].click();", option_button)
                break
    except TimeoutException:
        logging.info("Opções de download não apareceram.")
    time.sleep(5)


def process_notification(username, password, process_number):
    """
    Orquestra o processo de login, acesso ao processo e geração/baixa do PDF.
    """
    driver = create_driver()
    try:
        login(driver, username, password)
        access_process(driver, process_number)
        generate_pdf(driver)
        download_pdf(driver)
        return "PDF gerado e baixado com sucesso."
    except Exception as e:
        logging.error(f"Erro durante o processamento: {e}")
        raise
    finally:
        driver.quit()


def _gerar_modelo_1(doc, info, enderecos, numero_processo):
    try:
        # Adiciona o cabeçalho do documento
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        adicionar_paragrafo(doc, 
            "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", 
            negrito=True
        )
        adicionar_paragrafo(doc, 
            f"Referência: Processo Administrativo Sancionador nº: {numero_processo} ", 
            negrito=True
        )
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Prezado(a) Senhor(a),")
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, 
            "Informamos que foi proferido julgamento pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias no processo administrativo sancionador em referência, conforme decisão em anexo."
        )
        doc.add_paragraph("\n")

        adicionar_paragrafo(doc, "O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?", negrito=True)
        adicionar_paragrafo(doc, 
            "Sendo aplicada a penalidade de multa, esta notificação estará acompanhada de boleto bancário, que deverá ser pago até o vencimento."
        )
        adicionar_paragrafo(doc, 
            "O valor da multa poderá ser pago com 20% de desconto caso seja efetuado em até 20 dias contados de seu recebimento. "
            "Incorrerá em ilegalidade o usufruto do desconto em data posterior ao prazo referido, mesmo que a data impressa no boleto permita pagamento, "
            "sendo a diferença cobrada posteriormente pela Gerência de Gestão de Arrecadação (GEGAR). "
            "O pagamento da multa implica em desistência tácita do recurso, conforme art. 21 da Lei nº 6.437/1977."
        )
        adicionar_paragrafo(doc, 
            "O não pagamento do boleto sem que haja interposição de recurso, acarretará, sucessivamente: "
            "i) a inscrição do devedor no Cadastro Informativo de Crédito não Quitado do Setor Público Federal (CADIN); "
            "ii) a inscrição do débito em dívida ativa da União; iii) o ajuizamento de ação de execução fiscal contra o devedor; "
            "e iv) a comunicação aos cartórios de registros de imóveis, dos devedores inscritos em dívida ativa ou execução fiscal."
        )
        adicionar_paragrafo(doc, 
            "Esclarecemos que o valor da multa foi atualizado pela taxa Selic acumulada nos termos do art. 37-A da Lei 10.522/2002 "
            "e no art. 5º do Decreto-Lei 1.736/79."
        )
        doc.add_paragraph("\n")

        adicionar_paragrafo(doc, "COMO FAÇO PARA INTERPOR RECURSO DA DECISÃO?", negrito=True)
        adicionar_paragrafo(doc, 
            "Havendo interesse na interposição de recurso administrativo, este poderá ser interposto no prazo de 20 dias contados do recebimento desta notificação, "
            "conforme disposto no art. 9º da RDC nº 266/2019."
        )
        adicionar_paragrafo(doc, 
            "O protocolo do recurso deverá ser feito exclusivamente, por meio de peticionamento intercorrente no processo indicado no campo assunto desta notificação, "
            "pelo Sistema Eletrônico de Informações (SEI). Para tanto, é necessário, primeiramente, fazer o cadastro como usuário externo SEI-Anvisa. "
            "Acesse o portal da Anvisa https://www.gov.br/anvisa/pt-br > Sistemas > SEI > Acesso para Usuários Externos (SEI) e siga as orientações. "
            "Para maiores informações, consulte o Manual do Usuário Externo Sei-Anvisa, que está disponível em https://www.gov.br/anvisa/pt-br/sistemas/sei."
        )
        doc.add_paragraph("\n")

        adicionar_paragrafo(doc, "QUAIS DOCUMENTOS DEVEM ACOMPANHAR O RECURSO?", negrito=True)
        adicionar_paragrafo(doc, "a) Autuado pessoa jurídica:")
        adicionar_paragrafo(doc, "1. Contrato ou estatuto social da empresa, com a última alteração;")
        adicionar_paragrafo(doc, 
            "2. Procuração e documento de identificação do outorgado (advogado ou representante), caso constituído para atuar no processo. "
            "Somente serão aceitas procurações e substabelecimentos assinados eletronicamente, com certificação digital no padrão da "
            "Infraestrutura de Chaves Públicas Brasileira (ICP-Brasil) ou pelo assinador Gov.br."
        )
        adicionar_paragrafo(doc, 
            "3. Ata de eleição da atual diretoria quando a procuração estiver assinada por diretor que não conste como sócio da empresa;"
        )
        adicionar_paragrafo(doc, 
            "4. No caso de contestação sobre o porte da empresa considerado para a dosimetria da pena de multa: comprovação do porte econômico "
            "referente ao ano em que foi proferida a decisão (documentos previstos no art. 50 da RDC nº 222/2006)."
        )
        adicionar_paragrafo(doc, "b) Autuado pessoa física:")
        adicionar_paragrafo(doc, "1. Documento de identificação do autuado;")
        adicionar_paragrafo(doc, 
            "2. Procuração e documento de identificação do outorgado (advogado ou representante), caso constituído para atuar no processo."
        )
        doc.add_paragraph("\n")

    except Exception as e:
        print(f"Erro ao gerar o documento no modelo 1: {e}")

def _gerar_modelo_2(doc, info, enderecos, numero_processo):
    try:
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "MODELO 2 - Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        adicionar_paragrafo(doc, "Assunto: Modelo 2 - Detalhes Específicos.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº: {numero_processo} ", negrito=True)
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Este é o modelo 2 do documento.")
        doc.add_paragraph("\n")

    except Exception as e:
        print(f"Erro ao gerar o documento no modelo 2: {e}")

def _gerar_modelo_3(doc, info, enderecos, numero_processo):
    try:
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "MODELO 3 - Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        adicionar_paragrafo(doc, "Assunto: Modelo 3 - Informações Personalizadas.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº: {numero_processo} ", negrito=True)
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Este é o modelo 3 do documento.")
        doc.add_paragraph("\n")

    except Exception as e:
        print(f"Erro ao gerar o documento no modelo 3: {e}")

def escolher_enderecos(enderecos):
    """
    Permite ao usuário escolher quais endereços deseja manter e editar
    cada campo, se desejar.
    """
    if not enderecos:
        print("Nenhum endereço encontrado para editar.")
        return []

    print("\nForam encontrados os seguintes endereços:")
    selected_addresses = []

    for i, end in enumerate(enderecos, start=1):
        print(f"\n[{i}] Endereço detectado:")
        print(f"  Endereço: {end['endereco']}")
        print(f"  Cidade:   {end['cidade']}")
        print(f"  Bairro:   {end['bairro']}")
        print(f"  Estado:   {end['estado']}")
        print(f"  CEP:      {end['cep']}")

        # Perguntar se deseja manter
        keep = input("Deseja manter este endereço? (S/N): ")
        if keep.strip().lower() == 's':
            # Permitir edição
            edit = input("Deseja editar este endereço? (S/N): ")
            if edit.strip().lower() == 's':
                # Para cada campo, damos a opção de editar (ENTER para manter)
                novo_endereco = input(f"Endereço [{end['endereco']}]: ").strip()
                if novo_endereco:
                    end['endereco'] = novo_endereco

                nova_cidade = input(f"Cidade [{end['cidade']}]: ").strip()
                if nova_cidade:
                    end['cidade'] = nova_cidade

                novo_bairro = input(f"Bairro [{end['bairro']}]: ").strip()
                if novo_bairro:
                    end['bairro'] = novo_bairro

                novo_estado = input(f"Estado [{end['estado']}]: ").strip()
                if novo_estado:
                    end['estado'] = novo_estado

                novo_cep = input(f"CEP [{end['cep']}]: ").strip()
                if novo_cep:
                    end['cep'] = novo_cep

            # Após possível edição, adicionamos à lista
            selected_addresses.append(end)

    return selected_addresses

def get_latest_downloaded_file(download_directory):
    """
    Retorna o caminho do último arquivo baixado no diretório especificado.
    """
    try:
        files = [os.path.join(download_directory, f) for f in os.listdir(download_directory)]
        files = [f for f in files if os.path.isfile(f)]  # Filtra apenas arquivos
        latest_file = max(files, key=os.path.getmtime) if files else None
        return latest_file
    except Exception as e:
        print(f"Erro ao acessar o diretório de downloads: {e}")
        return None


def main():
    username = input("Digite seu usuário: ")
    password = getpass.getpass("Digite sua senha: ")
    process_number = input("Digite o número do processo: ")

    try:
        result = process_notification(username, password, process_number)
        print(result)
    except Exception as e:
        print(f"Erro: {e}")


if __name__ == "__main__":
    main()
