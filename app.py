import streamlit as st
import logging
import nest_asyncio
import time
import os
import tempfile
import unicodedata
import re
import spacy
import shutil

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

# Aplicação do nest_asyncio para permitir múltiplos loops de eventos (necessário se for rodar em notebook)
nest_asyncio.apply()

# Configuração de logs para Streamlit
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Função para carregar o modelo spaCy
def load_spacy_model():
    model_name = "pt_core_news_sm"  # Use "pt_core_news_lg" se preferir um modelo maior
    if not spacy.util.is_package(model_name):
        from spacy.cli import download
        with st.spinner("Baixando modelo spaCy (isso pode levar alguns minutos)..."):
            download(model_name)
    return spacy.load(model_name)

# Carregar o modelo
nlp = load_spacy_model()

# Constantes de elementos
LOGIN_URL = "https://sei.anvisa.gov.br/sip/login.php?sigla_orgao_sistema=ANVISA&sigla_sistema=SEI"
IFRAME_VISUALIZACAO_ID = "ifrVisualizacao"
BUTTON_XPATH_ALT = '//img[@title="Gerar Arquivo PDF do Processo"]/parent::a'

# Funções existentes
def create_driver(download_dir=None):
    if download_dir is None:
        # Usar um diretório temporário
        download_dir = tempfile.mkdtemp()
    
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Executa o Chrome em modo headless
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")  # Necessário em alguns ambientes de nuvem
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-notifications")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    chrome_options.set_capability("unhandledPromptBehavior", "ignore")

    # Especificar o caminho do Chromium instalado via apt
    chrome_options.binary_location = "/usr/bin/chromium"

    # Utilizar webdriver-manager para gerenciar o ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def wait_for_element(driver, by, value, timeout=20):
    try:
        logger.info(f"Aguardando elemento: {value}")
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return element
    except Exception as e:
        logger.error(f"Erro ao localizar o elemento: {value}")
        raise Exception(f"Elemento {value} não encontrado na página.") from e

def handle_alert(driver):
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = Alert(driver)
        alert_text = alert.text
        logger.warning(f"Alerta inesperado encontrado: {alert_text}")
        alert.accept()
        return alert_text
    except Exception:
        logger.info("Nenhum alerta encontrado.")
        return None

def login(driver, username, password):
    logger.info("Acessando a página de login.")
    driver.get(LOGIN_URL)
    wait_for_element(driver, By.ID, "txtUsuario").send_keys(username)
    driver.find_element(By.ID, "pwdSenha").send_keys(password)
    driver.find_element(By.ID, "sbmAcessar").click()

def access_process(driver, process_number):
    search_field = wait_for_element(driver, By.ID, "txtPesquisaRapida")
    search_field.send_keys(process_number)
    search_field.send_keys("\n")
    logger.info("Processo acessado com sucesso.")
    time.sleep(3)

def generate_pdf(driver):
    try:
        driver.switch_to.frame(
            wait_for_element(driver, By.ID, IFRAME_VISUALIZACAO_ID)
        )
        gerar_pdf_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, BUTTON_XPATH_ALT))
        )
        driver.execute_script("arguments[0].click();", gerar_pdf_button)
        logger.info("Clique no botão 'Gerar Arquivo PDF do Processo' realizado.")
        return "PDF gerado com sucesso."
    except Exception as e:
        logger.error(f"Erro ao gerar o PDF: {e}")
        raise Exception("Erro ao gerar o PDF do processo.")
    finally:
        driver.switch_to.default_content()
        time.sleep(5)

def download_pdf(driver, option="Todos os documentos disponíveis"):
    try:
        # Acessar o iframe 'ifrVisualizacao' e selecionar a opção de download
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="menu-opcao"]//button'))
        )
        logger.info("Opções de download detectadas.")

        for option_button in dropdown_options:
            if option_button.text.strip() == option:
                driver.execute_script("arguments[0].click();", option_button)
                logger.info(f"Opção '{option}' selecionada com sucesso.")
                break
        else:
            logger.warning(f"Opção '{option}' não encontrada. Prosseguindo sem selecionar opção.")

        time.sleep(5)
        logger.info("Download iniciado (ou já realizado com sucesso).")

    except Exception as e:
        logger.error(f"Erro ao tentar baixar o PDF: {e}")
        raise Exception("Erro durante o processo de download do PDF.") from e

def process_notification(username: str, password: str, process_number: str, download_dir):
    driver = create_driver(download_dir)
    try:
        login(driver, username, password)
        access_process(driver, process_number)
        generate_pdf(driver)
        try:
            download_pdf(driver, option="Todos os documentos disponíveis")
        except Exception as e:
            logger.warning(f"Erro não crítico no download_pdf: {e}")

        logger.info("Aguardando alguns segundos para permitir o download do PDF...")
        time.sleep(10)

        # Encontrar o arquivo PDF baixado
        files = [f for f in os.listdir(download_dir) if f.lower().endswith('.pdf')]
        if not files:
            raise Exception("Nenhum arquivo PDF foi baixado.")
        latest_file = max([os.path.join(download_dir, f) for f in files], key=os.path.getmtime)

        return latest_file
    except Exception as e:
        logger.exception("Erro durante o processamento.")
        raise e
    finally:
        driver.quit()

# Funções Auxiliares
def normalize_text(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def corrigir_texto(texto):
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
            page_text = page.extract_text()
            if page_text:
                text += page_text
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        logger.error(f"Erro ao processar PDF {pdf_path}: {e}")
        return ''

def extract_information_spacy(text):
    doc = nlp(text)

    info = {
        "nome_autuado": None,
        "cnpj_cpf": None,
        "socios_advogados": [],
        "emails": [],
    }

    for ent in doc.ents:
        if ent.label_ in ["PER", "ORG"]:
            if not info["nome_autuado"]:
                info["nome_autuado"] = ent.text.strip()
        elif ent.label_ == "EMAIL":
            info["emails"].append(ent.text.strip())

    cnpj_cpf_pattern = r"(?:CNPJ|CPF):\s*([\d./-]+)"
    match = re.search(cnpj_cpf_pattern, text)
    if match:
        info["cnpj_cpf"] = match.group(1)

    socios_adv_pattern = r"(?:Sócio|Advogado|Responsável|Representante Legal):\s*([\w\s]+)"
    info["socios_advogados"] = re.findall(socios_adv_pattern, text) or []

    return info

def extract_addresses_spacy(text):
    doc = nlp(text)

    addresses = []
    seen_addresses = set()

    endereco_pattern = r"(?:Endereço|End|Endereco):\s*([\w\s.,ºª-]+)"
    cidade_pattern = r"Cidade:\s*([\w\s]+(?: DE [\w\s]+)?)"
    bairro_pattern = r"Bairro:\s*([\w\s]+)"
    estado_pattern = r"Estado:\s*([A-Z]{2})"
    cep_pattern = r"CEP:\s*(\d{2}\.\d{3}-\d{3}|\d{5}-\d{3})"

    endereco_matches = re.findall(endereco_pattern, text)
    cidade_matches = re.findall(cidade_pattern, text)
    bairro_matches = re.findall(bairro_pattern, text)
    estado_matches = re.findall(estado_pattern, text)
    cep_matches = re.findall(cep_pattern, text)

    for i in range(max(len(endereco_matches), len(cidade_matches), len(bairro_matches), len(estado_matches), len(cep_matches))):
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

def _gerar_modelo_1(doc, info, enderecos, numero_processo):
    try:
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
        logger.error(f"Erro ao gerar o documento no modelo 1: {e}")

def _gerar_modelo_2(doc, info, enderecos, numero_processo):
    try:
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "MODELO 2 - Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for endereco in enderecos:
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
        logger.error(f"Erro ao gerar o documento no modelo 2: {e}")

def _gerar_modelo_3(doc, info, enderecos, numero_processo):
    try:
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "MODELO 3 - Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for endereco in enderecos:
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
        logger.error(f"Erro ao gerar o documento no modelo 3: {e}")

def escolher_enderecos(enderecos):
    if not enderecos:
        st.warning("Nenhum endereço encontrado para editar.")
        return []

    selected_addresses = []
    st.subheader("Endereços Detectados")

    for i, end in enumerate(enderecos, start=1):
        with st.expander(f"Endereço {i}"):
            st.write(f"**Endereço:** {end['endereco']}")
            st.write(f"**Cidade:** {end['cidade']}")
            st.write(f"**Bairro:** {end['bairro']}")
            st.write(f"**Estado:** {end['estado']}")
            st.write(f"**CEP:** {end['cep']}")

            keep = st.checkbox(f"Deseja manter este endereço? (Endereço {i})", value=True, key=f"keep_{i}")
            if keep:
                edit = st.checkbox(f"Deseja editar este endereço? (Endereço {i})", key=f"edit_{i}")
                if edit:
                    end['endereco'] = st.text_input(f"Endereço [{end['endereco']}]:", value=end['endereco'], key=f"endereco_{i}")
                    end['cidade'] = st.text_input(f"Cidade [{end['cidade']}]:", value=end['cidade'], key=f"cidade_{i}")
                    end['bairro'] = st.text_input(f"Bairro [{end['bairro']}]:", value=end['bairro'], key=f"bairro_{i}")
                    end['estado'] = st.text_input(f"Estado [{end['estado']}]:", value=end['estado'], key=f"estado_{i}")
                    end['cep'] = st.text_input(f"CEP [{end['cep']}]:", value=end['cep'], key=f"cep_{i}")
                selected_addresses.append(end)

    return selected_addresses

def get_latest_downloaded_file(download_directory):
    try:
        files = [os.path.join(download_directory, f) for f in os.listdir(download_directory) if os.path.isfile(os.path.join(download_directory, f))]
        files = [f for f in files if f.lower().endswith('.pdf')]
        latest_file = max(files, key=os.path.getmtime) if files else None
        return latest_file
    except Exception as e:
        logger.error(f"Erro ao acessar o diretório de downloads: {e}")
        return None

# Interface do Streamlit
def main():
    st.title("Automação de Notificações SEI-Anvisa")

    st.sidebar.header("Configurações")

    # Inputs do Usuário
    username = st.sidebar.text_input("Usuário")
    password = st.sidebar.text_input("Senha", type="password")
    process_number = st.sidebar.text_input("Número do Processo")
    # Diretório de downloads será um diretório temporário
    download_directory = tempfile.mkdtemp(prefix="downloads_")

    st.sidebar.write("**Diretório de Downloads:**")
    st.sidebar.write(download_directory)

    if st.sidebar.button("Iniciar Processo"):
        if not username or not password or not process_number:
            st.error("Por favor, preencha todos os campos.")
        else:
            with st.spinner("Processando..."):
                try:
                    latest_pdf = process_notification(username, password, process_number, download_directory)
                    st.success("PDF gerado e baixado automaticamente.")

                    # Exibir o caminho do PDF (opcional)
                    st.write(f"PDF salvo em: {latest_pdf}")

                    # Extrair texto do PDF
                    text = extract_text_with_pypdf2(latest_pdf)

                    if text:
                        st.success("Texto extraído com sucesso!")
                        numero_processo = extract_process_number(os.path.basename(latest_pdf))
                        info = extract_information_spacy(text)
                        addresses = extract_addresses_spacy(text)

                        # Permitir ao usuário editar os endereços
                        addresses = escolher_enderecos(addresses)

                        # Escolher o modelo do documento
                        modelo = st.selectbox("Escolha o modelo do documento:", ["Modelo 1", "Modelo 2", "Modelo 3"])

                        if st.button("Gerar Documento"):
                            doc = Document()
                            if modelo == "Modelo 1":
                                _gerar_modelo_1(doc, info, addresses, numero_processo)
                                tipo_documento = 1
                            elif modelo == "Modelo 2":
                                _gerar_modelo_2(doc, info, addresses, numero_processo)
                                tipo_documento = 2
                            elif modelo == "Modelo 3":
                                _gerar_modelo_3(doc, info, addresses, numero_processo)
                                tipo_documento = 3

                            output_dir = tempfile.mkdtemp(prefix="output_")
                            output_path = os.path.join(output_dir, f"Notificacao_Processo_Nº_{numero_processo}_modelo_{tipo_documento}.docx")
                            doc.save(output_path)
                            st.success(f"Documento gerado com sucesso.")

                            # Fornecer link de download
                            with open(output_path, "rb") as file:
                                st.download_button(
                                    label="Baixar Documento",
                                    data=file,
                                    file_name=os.path.basename(output_path),
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )

                    else:
                        st.error("Não foi possível extrair texto do arquivo.")

                except Exception as e:
                    st.error(f"Ocorreu um erro: {e}")

    st.header("Gerar Documento a Partir do PDF")

    if st.button("Processar Último PDF Baixado"):
        with st.spinner("Processando o último PDF baixado..."):
            try:
                latest_file = get_latest_downloaded_file(download_directory)

                if latest_file:
                    st.write(f"Último arquivo encontrado: {os.path.basename(latest_file)}")
                    try:
                        numero_processo = extract_process_number(os.path.basename(latest_file))
                        text = extract_text_with_pypdf2(latest_file)

                        if text:
                            st.success(f"Texto extraído com sucesso! Número do processo: {numero_processo}")
                            info = extract_information_spacy(text)
                            addresses = extract_addresses_spacy(text)

                            # Permitir ao usuário editar os endereços
                            addresses = escolher_enderecos(addresses)

                            # Escolher o modelo do documento
                            modelo = st.selectbox("Escolha o modelo do documento:", ["Modelo 1", "Modelo 2", "Modelo 3"])

                            if st.button("Gerar Documento"):
                                doc = Document()
                                if modelo == "Modelo 1":
                                    _gerar_modelo_1(doc, info, addresses, numero_processo)
                                    tipo_documento = 1
                                elif modelo == "Modelo 2":
                                    _gerar_modelo_2(doc, info, addresses, numero_processo)
                                    tipo_documento = 2
                                elif modelo == "Modelo 3":
                                    _gerar_modelo_3(doc, info, addresses, numero_processo)
                                    tipo_documento = 3

                                output_dir = tempfile.mkdtemp(prefix="output_")
                                output_path = os.path.join(output_dir, f"Notificacao_Processo_Nº_{numero_processo}_modelo_{tipo_documento}.docx")
                                doc.save(output_path)
                                st.success(f"Documento gerado com sucesso.")

                                # Fornecer link de download
                                with open(output_path, "rb") as file:
                                    st.download_button(
                                        label="Baixar Documento",
                                        data=file,
                                        file_name=os.path.basename(output_path),
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                        else:
                            st.error("Não foi possível extrair texto do arquivo.")
                    except Exception as e:
                        st.error(f"Ocorreu um erro: {e}")
                else:
                    st.warning("Nenhum arquivo encontrado no diretório de downloads.")
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")

if __name__ == '__main__':
    main()
