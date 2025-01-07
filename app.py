import streamlit as st
import logging
import nest_asyncio
import time
import getpass
import os
import re
import unicodedata
import tempfile

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from webdriver_manager.chrome import ChromeDriverManager

from docx import Document
from docx.shared import Pt
from PyPDF2 import PdfReader

# -------------------------------------------------------
# Aplicação do nest_asyncio para permitir múltiplos loops de eventos
# (às vezes necessário em notebooks, mas não atrapalha no Streamlit)
nest_asyncio.apply()

# -------------------------------------------------------
# Configuração de logs
logging.basicConfig(level=logging.INFO)

# -------------------------------------------------------
# Constantes de elementos (URLs e XPaths)
LOGIN_URL = "https://sei.anvisa.gov.br/sip/login.php?sigla_orgao_sistema=ANVISA&sigla_sistema=SEI"
IFRAME_VISUALIZACAO_ID = "ifrVisualizacao"
BUTTON_XPATH_ALT = '//img[@title="Gerar Arquivo PDF do Processo"]/parent::a'


# =========================
# Funções Auxiliares Selenium
# =========================
def create_driver(download_dir=None):
    """
    Configura e retorna uma instância do Selenium WebDriver,
    forçando o download de PDF ao invés de abrir no Chrome.
    No Streamlit Cloud, devemos usar o Chrome em modo headless.
    """
    if download_dir is None:
        # Cria uma pasta temporária para downloads
        download_dir = tempfile.mkdtemp(prefix="downloads_")

    chrome_options = Options()
    # Execução headless obrigatória em muitos ambientes online
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-dev-shm-usage")  # Pode ajudar em ambientes limitados de memória

    # Configura o Chrome para baixar PDFs sem abrir
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.set_capability("unhandledPromptBehavior", "ignore")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver, download_dir


def wait_for_element(driver, by, value, timeout=20):
    """
    Aguarda até que um elemento esteja presente no DOM.
    """
    try:
        logging.info(f"Aguardando elemento: {value}")
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return element
    except Exception as e:
        logging.error(f"Erro ao localizar o elemento: {value}")
        raise Exception(f"Elemento {value} não encontrado na página.") from e


def handle_alert(driver):
    """
    Captura e trata alertas inesperados sem recarregar a página.
    """
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = Alert(driver)
        alert_text = alert.text
        logging.warning(f"Alerta inesperado encontrado: {alert_text}")
        alert.accept()
        return alert_text
    except Exception:
        logging.info("Nenhum alerta encontrado.")
        return None


def login(driver, username, password):
    """
    Realiza o login no sistema SEI.
    """
    logging.info("Acessando a página de login.")
    driver.get(LOGIN_URL)
    wait_for_element(driver, By.ID, "txtUsuario").send_keys(username)
    driver.find_element(By.ID, "pwdSenha").send_keys(password)
    driver.find_element(By.ID, "sbmAcessar").click()
    time.sleep(3)


def access_process(driver, process_number):
    """
    Acessa um processo pelo número no sistema SEI.
    """
    search_field = wait_for_element(driver, By.ID, "txtPesquisaRapida")
    search_field.send_keys(process_number)
    search_field.send_keys("\n")
    logging.info("Processo acessado com sucesso.")
    time.sleep(3)


def generate_pdf(driver):
    """
    Gera o PDF do processo no iframe correspondente.
    Se o SEI exibir outro iframe ou outro comportamento, ajustar aqui.
    """
    try:
        driver.switch_to.frame(
            wait_for_element(driver, By.ID, IFRAME_VISUALIZACAO_ID)
        )
        gerar_pdf_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, BUTTON_XPATH_ALT))
        )
        driver.execute_script("arguments[0].click();", gerar_pdf_button)
        logging.info("Clique no botão 'Gerar Arquivo PDF do Processo' realizado.")
        return "PDF gerado com sucesso."
    except Exception as e:
        logging.error(f"Erro ao gerar o PDF: {e}")
        raise Exception("Erro ao gerar o PDF do processo.")
    finally:
        driver.switch_to.default_content()
        time.sleep(5)


def download_pdf(driver, option="Todos os documentos disponíveis"):
    """
    Realiza o clique no botão 'Gerar Arquivo PDF do Processo' e seleciona a opção desejada.
    :param driver: Instância do WebDriver.
    :param option: Opção de download: "Todos os documentos disponíveis", 
                   "Todos exceto selecionados" ou "Apenas selecionados".
    """
    from selenium.common.exceptions import TimeoutException

    try:
        # TENTATIVA: Acessar o iframe 'ifrVisualizacao'
        try:
            driver.switch_to.frame(wait_for_element(driver, By.ID, "ifrVisualizacao"))
            gerar_pdf_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="divInfraBarraComandosSuperior"]/button[1]')
                )
            )
            driver.execute_script("arguments[0].click();", gerar_pdf_button)
            logging.info("Botão 'Gerar Arquivo PDF do Processo' clicado no iframe 'ifrVisualizacao'.")
        except Exception as e:
            logging.warning(f"Falha ao encontrar ou clicar no botão 'Gerar Arquivo PDF': {e}")
            raise Exception("Erro ao clicar no botão de geração de PDF.")

        # Voltar ao contexto principal antes de selecionar opções
        driver.switch_to.default_content()

        # Aguardar as opções de download aparecerem
        try:
            dropdown_options = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, '//div[@class="menu-opcao"]//button'))
            )
            logging.info("Opções de download detectadas.")

            # Selecionar a opção desejada
            for option_button in dropdown_options:
                if option_button.text.strip() == option:
                    driver.execute_script("arguments[0].click();", option_button)
                    logging.info(f"Opção '{option}' selecionada com sucesso.")
                    break
            else:
                logging.warning(f"Opção '{option}' não encontrada. Prosseguindo sem selecionar opção.")
        except TimeoutException:
            logging.warning("Opções de download não apareceram. Prosseguindo...")

        # Aguardar o início do download
        time.sleep(5)
        logging.info("Download iniciado (ou já realizado com sucesso).")

    except Exception as e:
        logging.error(f"Erro ao tentar baixar o PDF: {e}")
        raise Exception("Erro durante o processo de download do PDF.") from e


def process_notification(username: str, password: str, process_number: str):
    """
    Orquestra o processo de login, acesso ao processo e geração/baixa do PDF.
    Retorna (caminho_da_pasta_download, mensagem_final).
    """
    # Cria o driver e o diretório de download
    driver, download_dir = create_driver()
    try:
        # Passo 1: Login
        login(driver, username, password)

        # Passo 2: Acessa o processo
        access_process(driver, process_number)

        # Passo 3: Gera PDF
        generate_pdf(driver)
        
        # Passo 4: Tenta baixar o PDF
        try:
            download_pdf(driver, option="Todos os documentos disponíveis")
        except Exception as e:
            logging.warning(f"Erro não crítico no download_pdf: {e}")

        # Passo 5: Aguarda alguns segundos para o PDF ser baixado
        logging.info("Aguardando alguns segundos para permitir o download do PDF...")
        time.sleep(10)

        return download_dir, "PDF gerado e baixado automaticamente."
    except Exception as e:
        logging.exception("Erro durante o processamento.")
        raise e
    finally:
        driver.quit()


# =========================
# Funções Auxiliares de Extração de Texto e Geração DOCX
# =========================

def normalize_text(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)  # Remove múltiplos espaços
    return text.strip()

def corrigir_texto(texto):
    """
    Ajuste manual de possíveis problemas de encoding.
    Inclua aqui outras correções necessárias.
    """
    substituicoes = {
        'Ã©': 'é',
        'Ã§Ã£o': 'ção',
        'Ã³': 'ó',
        'Ã': 'à',
    }
    for errado, correto in substituicoes.items():
        texto = texto.replace(errado, correto)
    return texto

def extract_text_with_pypdf2(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        print(f"Erro ao processar PDF {pdf_path}: {e}")
        return ''

def extract_information(text):
    autuado_pattern = r"(?:NOME AUTUADO|Autuado|Empresa|Razão Social):\s*([\w\s,.-]+)"
    cnpj_cpf_pattern = r"(?:CNPJ|CPF):\s*([\d./-]+)"
    socios_adv_pattern = r"(?:Sócio|Advogado|Responsável|Representante Legal):\s*([\w\s]+)"
    email_pattern = r"(?:E-mail|Email):\s*([\w.-]+@[\w.-]+\.[a-z]{2,})"

    info = {
        "nome_autuado": re.search(autuado_pattern, text).group(1) if re.search(autuado_pattern, text) else None,
        "cnpj_cpf": re.search(cnpj_cpf_pattern, text).group(1) if re.search(cnpj_cpf_pattern, text) else None,
        "socios_advogados": re.findall(socios_adv_pattern, text) or [],
        "emails": re.findall(email_pattern, text) or [],
    }
    return info

def extract_addresses(text):
    endereco_pattern = r"(?:Endereço|End|Endereco):\s*([\w\s.,ºª-]+)"
    cidade_pattern = r"Cidade:\s*([\w\s]+(?: DE [\w\s]+)?)"
    bairro_pattern = r"Bairro:\s*([\w\s]+)"
    estado_pattern = r"Estado:\s*([A-Z]{2})"
    cep_pattern = r"CEP:\s*(\d{2}\.\d{3}-\d{3}|\d{5}-\d{3})"

    addresses = []
    seen_addresses = set()

    endereco_matches = re.findall(endereco_pattern, text)
    cidade_matches = re.findall(cidade_pattern, text)
    bairro_matches = re.findall(bairro_pattern, text)
    estado_matches = re.findall(estado_pattern, text)
    cep_matches = re.findall(cep_pattern, text)

    for i in range(
        max(
            len(endereco_matches),
            len(cidade_matches),
            len(bairro_matches),
            len(estado_matches),
            len(cep_matches)
        )
    ):
        endereco = endereco_matches[i].strip() if i < len(endereco_matches) else None
        cidade = cidade_matches[i].strip() if i < len(cidade_matches) else None
        bairro = bairro_matches[i].strip() if i < len(bairro_matches) else None
        estado = estado_matches[i].strip() if i < len(estado_matches) else None
        cep = cep_matches[i].strip() if i < len(cep_matches) else None

        if endereco and endereco not in seen_addresses:
            seen_addresses.add(endereco)
            addresses.append({
                "endereco": endereco,
                "cidade": cidade,
                "bairro": bairro,
                "estado": estado,
                "cep": cep
            })

    return addresses or []

def adicionar_paragrafo(doc, texto="", negrito=False, tamanho=12):
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return paragrafo

def extract_process_number(file_name):
    base_name = os.path.splitext(file_name)[0]
    if base_name.startswith("SEI"):
        base_name = base_name[3:].strip()
    return base_name

def gerar_documento_streamlit(info, enderecos, numero_processo):
    """
    Gera um documento DOCX com informações do processo e endereços extraídos.
    Retorna o caminho do arquivo gerado.
    """
    try:
        output_directory = tempfile.mkdtemp(prefix="docx_output_")
        output_path = os.path.join(
            output_directory, f"Notificacao_Processo_Nº_{numero_processo}.docx"
        )

        doc = Document()
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

        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº: {numero_processo} ", negrito=True)
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Prezado(a) Senhor(a),")
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Informamos que foi proferido julgamento ... (texto exemplo).")
        doc.add_paragraph("\n")

        # Exemplo de subtítulo e texto explicativo
        adicionar_paragrafo(doc, "O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?", negrito=True)
        adicionar_paragrafo(doc, "Sendo aplicada a penalidade de multa, ...")
        # etc... (adicione o restante do texto)

        # Salva o documento
        doc.save(output_path)
        return output_path
    except Exception as e:
        print(f"Erro ao gerar o documento DOCX: {e}")
        return None

def get_latest_downloaded_file(download_directory):
    """
    Retorna o caminho do último arquivo baixado no diretório especificado.
    """
    files = [os.path.join(download_directory, f) for f in os.listdir(download_directory)]
    files = [f for f in files if os.path.isfile(f)]
    if not files:
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

# =========================
# APP STREAMLIT
# =========================
def main():
    st.title("Automação SEI Anvisa - Exemplo Selenium + Streamlit")

    st.write("Este é um exemplo de aplicação Streamlit que faz login no SEI da Anvisa, gera e baixa o PDF de um processo, extrai dados e gera uma notificação em DOCX.")

    # Entradas do usuário
    username = st.text_input("Digite seu usuário (SEI)", value="")
    password = st.text_input("Digite sua senha (SEI)", value="", type="password")
    process_number = st.text_input("Digite o número do processo", value="")

    if st.button("Executar Processo"):
        if not username or not password or not process_number:
            st.warning("Por favor, preencha usuário, senha e número do processo.")
            return

        with st.spinner("Executando automação... aguarde..."):
            try:
                download_folder, msg = process_notification(username, password, process_number)
                st.success(msg)

                # Após o download, vamos tentar descobrir o PDF baixado
                pdf_baixado = get_latest_downloaded_file(download_folder)
                if pdf_baixado and pdf_baixado.lower().endswith(".pdf"):
                    st.info(f"PDF encontrado: {os.path.basename(pdf_baixado)}")

                    # Extrai texto do PDF
                    texto_extraido = extract_text_with_pypdf2(pdf_baixado)
                    if not texto_extraido:
                        st.warning("Não foi possível extrair texto do PDF.")
                    else:
                        st.success("Texto extraído com sucesso do PDF!")

                        # Descobre número do processo a partir do nome do arquivo
                        numero_processo_extraido = extract_process_number(os.path.basename(pdf_baixado))
                        st.write(f"Número de processo extraído: {numero_processo_extraido}")

                        # Extrai informações do texto
                        info = extract_information(texto_extraido)
                        enderecos = extract_addresses(texto_extraido)

                        # Gera documento .docx
                        docx_path = gerar_documento_streamlit(info, enderecos, numero_processo_extraido)
                        if docx_path:
                            st.success("Documento .docx gerado com sucesso.")
                            with open(docx_path, "rb") as file:
                                st.download_button(
                                    label="Baixar Notificação (DOCX)",
                                    data=file.read(),
                                    file_name=os.path.basename(docx_path),
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                        else:
                            st.error("Erro ao gerar o documento .docx.")
                else:
                    st.warning("Nenhum PDF foi encontrado no diretório de downloads.")
            
            except Exception as e:
                st.error(f"Ocorreu um erro na automação: {e}")

if __name__ == "__main__":
    main()
