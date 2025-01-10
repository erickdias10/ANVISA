import streamlit as st
import logging
import nest_asyncio
import time
import getpass
import os
import unicodedata
import re
import spacy

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt
from io import BytesIO

# Configuração de logs
logging.basicConfig(level=logging.INFO)

# Aplicação do nest_asyncio para permitir múltiplos loops de eventos (necessário se for rodar em notebook)
nest_asyncio.apply()

# Constantes de elementos
LOGIN_URL = "https://sei.anvisa.gov.br/sip/login.php?sigla_orgao_sistema=ANVISA&sigla_sistema=SEI"

def create_browser_context():
    """
    Configura e retorna uma instância do Playwright BrowserContext e Page,
    utilizando um diretório de dados de usuário persistente para manter a sessão.
    """
    download_dir = os.path.join(os.getcwd(), "downloads")
    os.makedirs(download_dir, exist_ok=True)

    # Diretório para armazenar dados de usuário persistentes
    user_data_dir = os.path.join(os.getcwd(), "user_data")
    os.makedirs(user_data_dir, exist_ok=True)

    playwright = sync_playwright().start()
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=user_data_dir,
        headless=True,  # True para rodar em modo headless no servidor
        accept_downloads=True,  # Permite downloads automáticos
        downloads_path=download_dir  # Define o diretório de downloads
    )
    page = context.new_page()
    return playwright, context, page

def wait_for_element(page, selector, timeout=20000):
    """
    Aguarda até que um elemento esteja presente no DOM.
    """
    try:
        logging.info(f"Aguardando elemento: {selector}")
        element = page.wait_for_selector(selector, timeout=timeout)
        if element:
            return element
    except PlaywrightTimeoutError:
        logging.error(f"Erro ao localizar o elemento: {selector}")
        raise Exception(f"Elemento {selector} não encontrado na página.")
    return None

def handle_download(download, download_dir):
    """
    Manipula o download, salvando-o no diretório de downloads especificado.
    """
    os.makedirs(download_dir, exist_ok=True)
    download_path = os.path.join(download_dir, download.suggested_filename)
    download.save_as(download_path)
    logging.info(f"Download salvo em: {download_path}")
    return download_path

def handle_alert(page):
    """
    Captura e trata alertas inesperados sem recarregar a página.
    """
    try:
        dialog = page.expect_event("dialog", timeout=5000)
        if dialog:
            alert_text = dialog.message
            logging.warning(f"Alerta inesperado encontrado: {alert_text}")
            dialog.accept()
            return alert_text
    except PlaywrightTimeoutError:
        logging.info("Nenhum alerta encontrado.")
        return None

def login(page, username, password):
    """
    Realiza o login no sistema SEI.
    """
    logging.info("Acessando a página de login.")
    page.goto(LOGIN_URL)

    # Aguarda o campo de usuário aparecer e preenche
    user_field = wait_for_element(page, "#txtUsuario")
    if user_field:
        user_field.fill(username)
        logging.info("Campo de usuário preenchido.")
    else:
        raise Exception("Campo de usuário não encontrado.")

    # Aguarda o campo de senha aparecer e preenche
    password_field = wait_for_element(page, "#pwdSenha")
    if password_field:
        password_field.fill(password)
        logging.info("Campo de senha preenchido.")
    else:
        raise Exception("Campo de senha não encontrado.")

    # Aguarda o botão de login aparecer e clica
    login_button = wait_for_element(page, "#sbmAcessar")
    if login_button:
        login_button.click()
        logging.info("Botão de login clicado.")
    else:
        raise Exception("Botão de login não encontrado.")

    # Aguarda a página principal carregar após login
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
        logging.info("Login realizado com sucesso.")
    except PlaywrightTimeoutError:
        logging.error("Tempo esgotado aguardando a página principal carregar após login.")
        raise Exception("Login pode não ter sido realizado com sucesso.")

def access_process(page, process_number):
    """
    Acessa um processo pelo número no sistema SEI.
    """
    search_field = wait_for_element(page, "#txtPesquisaRapida")
    search_field.fill(process_number)
    search_field.press("Enter")
    logging.info("Processo acessado com sucesso.")
    time.sleep(3)  # Aguarda a página carregar resultados

# Defina seus identificadores e XPaths
IFRAME_VISUALIZACAO_ID = "ifrVisualizacao"
BUTTON_XPATH_GERAR_PDF = '//*[@id="divArvoreAcoes"]/a[7]/img'  # XPath para o botão 'Gerar PDF'
BUTTON_XPATH_DOWNLOAD_OPTION = '//*[@id="divInfraBarraComandosSuperior"]/button[1]'  # XPath para a opção de download

def generate_and_download_pdf(page, download_dir):
    """
    Gera e baixa o PDF do processo no iframe correspondente.
    :param page: Instância do Playwright Page.
    :param download_dir: Diretório para salvar o download.
    """
    try:
        # Espera pelo iframe e obtém o handle
        logging.info(f"Esperando pelo iframe com ID {IFRAME_VISUALIZACAO_ID}")
        iframe_element = page.wait_for_selector(f'iframe#{IFRAME_VISUALIZACAO_ID}', timeout=10000)
        if not iframe_element:
            raise Exception(f"Iframe com ID {IFRAME_VISUALIZACAO_ID} não encontrado.")
        
        # Acessa o contexto do iframe
        iframe = iframe_element.content_frame()
        if not iframe:
            raise Exception("Não foi possível acessar o conteúdo do iframe.")

        logging.info("Iframe encontrado e acessado com sucesso.")

        # Espera pelo botão de gerar PDF dentro do iframe
        logging.info(f"Esperando pelo botão de gerar PDF com XPath {BUTTON_XPATH_GERAR_PDF}")
        gerar_pdf_button = iframe.wait_for_selector(f'xpath={BUTTON_XPATH_GERAR_PDF}', timeout=10000)
        if not gerar_pdf_button:
            raise Exception("Botão para gerar PDF não encontrado.")

        # Clicar no botão 'Gerar PDF' para abrir as opções de download
        logging.info("Clicando no botão 'Gerar Arquivo PDF do Processo'")
        gerar_pdf_button.click()
        logging.info("Clique no botão 'Gerar Arquivo PDF do Processo' realizado.")

        # Opcional: esperar que as opções de download apareçam
        time.sleep(2)  # Aguarda um breve momento para que as opções de download apareçam

        # Espera pelo botão de opção de download dentro do iframe
        logging.info(f"Esperando pelo botão de opção de download com XPath {BUTTON_XPATH_DOWNLOAD_OPTION}")
        download_option_button = iframe.wait_for_selector(f'xpath={BUTTON_XPATH_DOWNLOAD_OPTION}', timeout=10000)
        if not download_option_button:
            raise Exception("Botão de opção de download não encontrado.")

        # Clicar no botão de opção de download e capturar o download
        logging.info("Clicando no botão de opção de download.")
        with page.expect_download(timeout=60000) as download_info_option:
            download_option_button.click()
        download_option = download_info_option.value
        download_option_path = handle_download(download_option, download_dir)
        logging.info("Clique no botão de opção de download realizado.")

        # Retorna o caminho do download principal
        return download_option_path

    except PlaywrightTimeoutError as e:
        logging.error(f"Timeout ao gerar o PDF: {e}")
        raise Exception("Timeout ao gerar o PDF do processo.")
    except Exception as e:
        logging.error(f"Erro ao gerar o PDF: {e}")
        raise Exception("Erro ao gerar o PDF do processo.")
    finally:
        # Opcional: espera adicional se necessário
        time.sleep(5)

def process_notification(username, password, process_number):
    """
    Orquestra o processo de login, acesso ao processo e geração/baixa do PDF.
    """
    playwright, context, page = create_browser_context()
    download_dir = os.path.join(os.getcwd(), "downloads")
    try:
        # Passo 1: Login
        login(page, username, password)

        # Passo 2: Acessa o processo
        access_process(page, process_number)

        # Passo 3: Gera e baixa o PDF
        try:
            download_path = generate_and_download_pdf(page, download_dir)  # Função consolidada
            logging.info(f"PDF gerado e salvo em: {download_path}")
        except Exception as e:
            logging.error(f"Erro ao gerar o PDF: {e}")
            raise Exception("Erro durante o processo de geração do PDF.")

        logging.info("PDF gerado com sucesso.")

        return download_path  # Retorna o caminho do PDF baixado
    except Exception as e:
        logging.error(f"Erro durante o processamento: {e}")
        raise e
    finally:
        # Fechar o navegador
        context.close()
        playwright.stop()

def normalize_text(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)  # Remove múltiplos espaços
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
        st.error(f"Erro ao processar PDF {pdf_path}: {e}")
        return ''

def extract_information_spacy(text):
    """
    Extrai informações do texto utilizando spaCy.
    """
    doc = nlp(text)

    info = {
        "nome_autuado": None,
        "cnpj_cpf": None,
        "socios_advogados": [],
        "emails": [],
    }

    for ent in doc.ents:
        if ent.label_ in ["PER", "ORG"]:  # Pessoa ou Organização
            if not info["nome_autuado"]:
                info["nome_autuado"] = ent.text.strip()
        elif ent.label_ == "EMAIL":
            info["emails"].append(ent.text.strip())

    # Usar regex para complementar a extração de CNPJ/CPF
    cnpj_cpf_pattern = r"(?:CNPJ|CPF):\s*([\d./-]+)"
    match = re.search(cnpj_cpf_pattern, text)
    if match:
        info["cnpj_cpf"] = match.group(1)

    # Sócios ou Advogados mencionados
    socios_adv_pattern = r"(?:Sócio|Advogado|Responsável|Representante Legal):\s*([\w\s]+)"
    info["socios_advogados"] = re.findall(socios_adv_pattern, text) or []

    return info

def extract_addresses_spacy(text):
    """
    Extrai endereços do texto utilizando spaCy e complementa com regex.
    """
    doc = nlp(text)

    addresses = []
    seen_addresses = set()

    # Usar regex para capturar padrões específicos
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
        st.error(f"Erro ao gerar o documento no modelo 1: {e}")

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
        st.error(f"Erro ao gerar o documento no modelo 2: {e}")

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
        st.error(f"Erro ao gerar o documento no modelo 3: {e}")

def main():
    st.title("Gerador de Notificações SEI-Anvisa")

    st.sidebar.header("Informações de Login")
    username = st.sidebar.text_input("Usuário")
    password = st.sidebar.text_input("Senha", type="password")

    st.header("Processo Administrativo")
    process_number = st.text_input("Número do Processo")

    if st.button("Gerar Notificação"):
        if not username or not password or not process_number:
            st.error("Por favor, preencha todos os campos.")
            return

        with st.spinner("Processando..."):
            try:
                download_path = process_notification(username, password, process_number)
                st.success("PDF gerado com sucesso!")

                if download_path:
                    st.info(f"Arquivo PDF baixado: {download_path}")
                    numero_processo = extract_process_number(os.path.basename(download_path))
                    text = extract_text_with_pypdf2(download_path)

                    if text:
                        st.success("Texto extraído com sucesso!")
                        info = extract_information_spacy(text)
                        addresses = extract_addresses_spacy(text)

                        # Exibir informações extraídas
                        st.subheader("Informações Extraídas")
                        st.write(f"**Nome Autuado:** {info.get('nome_autuado', 'Não informado')}")
                        st.write(f"**CNPJ/CPF:** {info.get('cnpj_cpf', 'Não informado')}")
                        st.write(f"**Emails:** {', '.join(info.get('emails', []))}")
                        st.write(f"**Sócios/Advogados:** {', '.join(info.get('socios_advogados', []))}")

                        st.subheader("Endereços Encontrados")
                        for idx, end in enumerate(addresses, start=1):
                            st.write(f"**Endereço {idx}:**")
                            st.write(f"  - Endereço: {end['endereco']}")
                            st.write(f"  - Cidade: {end['cidade']}")
                            st.write(f"  - Bairro: {end['bairro']}")
                            st.write(f"  - Estado: {end['estado']}")
                            st.write(f"  - CEP: {end['cep']}")

                        # Permitir ao usuário editar os endereços
                        st.subheader("Editar Endereços")
                        edited_addresses = []
                        for idx, end in enumerate(addresses, start=1):
                            st.write(f"**Endereço {idx}:**")
                            endereco = st.text_input(f"Endereço {idx}", value=end['endereco'])
                            cidade = st.text_input(f"Cidade {idx}", value=end['cidade'])
                            bairro = st.text_input(f"Bairro {idx}", value=end['bairro'])
                            estado = st.text_input(f"Estado {idx}", value=end['estado'])
                            cep = st.text_input(f"CEP {idx}", value=end['cep'])
                            edited_addresses.append({
                                "endereco": endereco,
                                "cidade": cidade,
                                "bairro": bairro,
                                "estado": estado,
                                "cep": cep
                            })

                        # Escolha do modelo de documento
                        st.subheader("Escolha o Modelo do Documento")
                        modelo = st.selectbox("Selecione o modelo desejado:", ["Modelo 1", "Modelo 2", "Modelo 3"])

                        if st.button("Gerar Documento Word"):
                            doc = Document()
                            if modelo == "Modelo 1":
                                _gerar_modelo_1(doc, info, edited_addresses, numero_processo)
                            elif modelo == "Modelo 2":
                                _gerar_modelo_2(doc, info, edited_addresses, numero_processo)
                            elif modelo == "Modelo 3":
                                _gerar_modelo_3(doc, info, edited_addresses, numero_processo)

                            # Salvar documento em buffer
                            buffer = BytesIO()
                            doc.save(buffer)
                            buffer.seek(0)

                            # Nome do arquivo
                            output_filename = f"Notificacao_Processo_Nº_{numero_processo}_modelo_{modelo.split()[-1]}.docx"

                            st.download_button(
                                label="Baixar Documento",
                                data=buffer,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            st.success("Documento gerado e pronto para download!")

                    else:
                        st.error("Não foi possível extrair texto do arquivo PDF.")
                else:
                    st.error("Nenhum arquivo PDF encontrado no diretório de downloads.")
            except Exception as ex:
                st.error(f"Ocorreu um erro: {ex}")

if __name__ == '__main__':
    # Carregar o modelo spaCy para português
    try:
        nlp = spacy.load("pt_core_news_lg")
    except OSError:
        st.info("Modelo 'pt_core_news_lg' não encontrado. Instalando...")
        os.system("python -m spacy download pt_core_news_lg")
        nlp = spacy.load("pt_core_news_lg")

    main()
