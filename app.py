# Bloco 1: Importação de Bibliotecas
import re
from PyPDF2 import PdfReader
import unicodedata
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
import os
import glob
import time
import joblib
import streamlit as st

def predict_addresses_with_model(text, vectorizer_path="vectorizer.pkl", model_path="address_model.pkl"):
    """
    Prediz endereços em um texto usando um modelo treinado.

    Args:
        text (str): Texto a ser analisado.
        vectorizer_path (str): Caminho para o arquivo do vetorizar.
        model_path (str): Caminho para o modelo treinado.

    Returns:
        list: Lista de endereços previstos.
    """
    try:
        # Carrega o vetorizar e o modelo
        vectorizer = joblib.load(vectorizer_path)
        model = joblib.load(model_path)

        # Prepara o texto para predição
        text_vectorized = vectorizer.transform([text])
        predictions = model.predict(text_vectorized)

        # Retorna as previsões
        return predictions
    except Exception as e:
        print(f"Erro ao fazer predição com o modelo: {e}")
        return []

def predict_Nome_Email_with_model(text, vectorizer_path="vectorizer_Nome.pkl", model_path="modelo_Nome.pkl"):
    """
    Prediz Nomes dos Autuados, CPF, CNPJ e Emails em um texto usando um modelo treinado.

    Args:
        text (str): Texto a ser analisado.
        vectorizer_path (str): Caminho para o arquivo do vetorizar.
        model_path (str): Caminho para o modelo treinado.

    Returns:
        list: Lista de endereços previstos.
    """
    try:
        # Carrega o vetorizar e o modelo
        vectorizer = joblib.load(vectorizer_path)
        model = joblib.load(model_path)

        # Prepara o texto para predição
        text_vectorized = vectorizer.transform([text])
        predictions = model.predict(text_vectorized)

        # Retorna as previsões
        return predictions
    except Exception as e:
        print(f"Erro ao fazer predição com o modelo: {e}")
        return []

def buscar_ultimo_arquivo_baixado(diretorio_downloads):
    """
    Busca o último arquivo baixado no diretório especificado.

    Args:
        diretorio_downloads (str): Caminho para o diretório de downloads.

    Returns:
        str: Caminho completo do último arquivo baixado, ou None se nenhum arquivo for encontrado.
    """
    try:
        arquivos = glob.glob(os.path.join(diretorio_downloads, "*"))
        if not arquivos:
            print("Nenhum arquivo encontrado no diretório de downloads.")
            return None

        ultimo_arquivo = max(arquivos, key=os.path.getmtime)
        print(f"Último arquivo baixado encontrado: {ultimo_arquivo}")
        return ultimo_arquivo
    except Exception as e:
        print(f"Erro ao buscar o último arquivo baixado: {e}")
        return None

def normalize_text(text):
    """
    Remove caracteres especiais e normaliza o texto.
    """
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)  # Remove múltiplos espaços
    return text.strip()

def corrigir_texto(texto):
    """
    Corrige caracteres corrompidos em texto.
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
    """
    Extrai texto de PDFs usando PyPDF2.
    """
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        print(f"Erro ao processar PDF com PyPDF2 {pdf_path}: {e}")
        return ''

# Bloco 4: Processamento de Endereços e Formatação do Documento
def extract_addresses(text):
    """
    Extrai informações de endereço do texto usando expressões regulares.

    Args:
        text (str): Texto extraído do PDF.

    Returns:
        list: Lista de dicionários contendo os endereços extraídos.
    """
    addresses = []
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
        address = {
            "endereco": endereco_matches[i].strip() if i < len(endereco_matches) else None,
            "cidade": cidade_matches[i].strip() if i < len(cidade_matches) else None,
            "bairro": bairro_matches[i].strip() if i < len(bairro_matches) else None,
            "estado": estado_matches[i].strip() if i < len(estado_matches) else None,
            "cep": cep_matches[i].strip() if i < len(cep_matches) else None
        }
        if any(address.values()) and address not in addresses:
            addresses.append(address)

    return addresses

def adicionar_paragrafo(doc, texto="", negrito=False, tamanho=12):
    """
    Adiciona um parágrafo ao documento com texto opcionalmente em negrito e com tamanho de fonte ajustável.
    """
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return paragrafo

import re
from PyPDF2 import PdfReader
import unicodedata
from tkinter import Tk, filedialog
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
import pyautogui
import os
import glob
import time
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt


def move_and_click(x, y):
    """
    Move o cursor do mouse para as coordenadas especificadas e clica.

    Args:
        x (int): Coordenada X para mover o mouse.
        y (int): Coordenada Y para mover o mouse.
    """
    try:
        pyautogui.moveTo(x, y)
        pyautogui.click()
    except Exception as e:
        print(f"Erro ao clicar nas coordenadas ({x}, {y}): {e}")


def buscar_processo(processo):
    """
    Realiza a busca de um processo no sistema usando automação.
    
    Args:
        processo (str): Número do processo a ser buscado.
    """
    try:
        print("Buscando o processo...")
        move_and_click(1465, 199)  # Coordenadas do campo de busca
        pyautogui.write(processo)  # Digita o número do processo
        pyautogui.press("enter")  # Pressiona Enter para buscar
        time.sleep(10)  # Aguarda o carregamento
    except Exception as e:
        print(f"Erro ao buscar o processo: {e}")


def baixar_processo():
    """
    Realiza o download do processo em formato PDF usando automação.
    """
    try:
        print("Baixando o processo...")
        move_and_click(1084, 256)  # Botão para gerar em PDF
        time.sleep(10)  # Aguarda o carregamento do botão Gerar
        move_and_click(1792, 288)  # Botão para confirmar geração
        time.sleep(20)  # Aguarda o download do arquivo
    except Exception as e:
        print(f"Erro ao baixar o processo: {e}")


def buscar_ultimo_arquivo_baixado(diretorio_downloads):
    """
    Busca o último arquivo baixado no diretório especificado.
    
    Args:
        diretorio_downloads (str): Caminho para o diretório de downloads.
        
    Returns:
        str: Caminho completo do último arquivo baixado, ou None se nenhum arquivo for encontrado.
    """
    try:
        arquivos = glob.glob(os.path.join(diretorio_downloads, "*"))
        
        if not arquivos:
            print("Nenhum arquivo encontrado no diretório de downloads.")
            return None

        ultimo_arquivo = max(arquivos, key=os.path.getmtime)
        print(f"Último arquivo baixado encontrado: {ultimo_arquivo}")
        return ultimo_arquivo
    except Exception as e:
        print(f"Erro ao buscar o último arquivo baixado: {e}")
        return None


def normalize_text(text):
    """
    Remove caracteres especiais e normaliza o texto.
    """
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)  # Remove múltiplos espaços
    return text.strip()


def corrigir_texto(texto):
    """
    Corrige caracteres corrompidos em texto.
    """
    substituicoes = {
        'Ã©': 'é',
        'Ã§Ã£o': 'ção',
        'Ã³': 'ó',
        'Ã': 'à',
        # Adicione mais substituições conforme necessário
    }
    for errado, correto in substituicoes.items():
        texto = texto.replace(errado, correto)
    return texto


def clean_fragmented_text(text):
    """
    Limpa fragmentos de palavras coladas.
    """
    return re.sub(r"(\w)\s+(\w)", r"\1 \2", text).strip()


def extract_text_with_pypdf2(pdf_path):
    """
    Extrai texto de PDFs usando PyPDF2.
    """
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return clean_fragmented_text(text)
    except Exception as e:
        print(f"Erro ao processar PDF com PyPDF2 {pdf_path}: {e}")
        return ''


def extract_addresses(text):
    """
    Extrai informações de endereço do texto.
    """
    addresses = []
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
        address = {
            "endereco": endereco_matches[i].strip() if i < len(endereco_matches) else None,
            "cidade": cidade_matches[i].strip() if i < len(cidade_matches) else None,
            "bairro": bairro_matches[i].strip() if i < len(bairro_matches) else None,
            "estado": estado_matches[i].strip() if i < len(estado_matches) else None,
            "cep": cep_matches[i].strip() if i < len(cep_matches) else None
        }
        if any(address.values()) and address not in addresses:
            addresses.append(address)

    return addresses


def adicionar_paragrafo(doc, texto="", negrito=False, tamanho=12):
    """
    Adiciona um parágrafo ao documento com texto opcionalmente em negrito e com tamanho de fonte ajustável.
    
    Args:
        doc (Document): Documento onde o parágrafo será adicionado.
        texto (str): Texto do parágrafo.
        negrito (bool): Define se o texto será em negrito.
        tamanho (int): Tamanho da fonte.
    """
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return paragrafo

def gerar_documento_docx(process_number, info, enderecos, output_path="Notificacao_Processo_Nº_{process_number}.docx"):
    """
    Gera um documento DOCX com informações do processo e endereços extraídos.

    Args:
        process_number (str): Número do processo administrativo.
        info (dict): Dicionário com informações extraídas do texto.
        enderecos (list): Lista de dicionários contendo informações de endereços.
        output_path (str): Caminho para salvar o documento gerado.
    """
    try:
        diretorio_downloads = os.path.expanduser("~/Downloads")
        output_path = os.path.join(diretorio_downloads, f"Notificacao_Processo_Nº_{process_number}.docx")
        
        doc = Document()

        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "[Ao Senhor/À Senhora]")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        # Adiciona endereços
        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        # Corpo principal
            # Corpo principal
        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº {process_number}", negrito=True)
        doc.add_paragraph("\n")  # Quebra de linha
        adicionar_paragrafo(doc, "Prezado(a) Senhor(a),")
        doc.add_paragraph("\n")  # Quebra de linha
        adicionar_paragrafo(doc, "Informamos que foi proferido julgamento pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias no processo administrativo sancionador em referência, conforme decisão em anexo.")
        doc.add_paragraph("\n")  # Quebra de linha
        
        # O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?
        adicionar_paragrafo(doc, "O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?", negrito=True)
        adicionar_paragrafo(doc, "Sendo aplicada a penalidade de multa, esta notificação estará acompanhada de boleto bancário, que deverá ser pago até o vencimento.")
        adicionar_paragrafo(doc, "O valor da multa poderá ser pago com 20% de desconto caso seja efetuado em até 20 dias contados de seu recebimento. Incorrerá em ilegalidade o usufruto do desconto em data posterior ao prazo referido, mesmo que a data impressa no boleto permita pagamento, sendo a diferença cobrada posteriormente pela Gerência de Gestão de Arrecadação (GEGAR). O pagamento da multa implica em desistência tácita do recurso, conforme art. 21 da Lei nº 6.437/1977.")
        adicionar_paragrafo(doc, "O não pagamento do boleto sem que haja interposição de recurso, acarretará, sucessivamente: i) a inscrição do devedor no Cadastro Informativo de Crédito não Quitado do Setor Público Federal (CADIN); ii) a inscrição do débito em dívida ativa da União; iii) o ajuizamento de ação de execução fiscal contra o devedor; e iv) a comunicação aos cartórios de registros de imóveis, dos devedores inscritos em dívida ativa ou execução fiscal.")
        adicionar_paragrafo(doc, "Esclarecemos que o valor da multa foi atualizado pela taxa Selic acumulada nos termos do art. 37-A da Lei 10.522/2002 e no art. 5º do Decreto-Lei 1.736/79.")
        doc.add_paragraph("\n")  # Quebra de linha
        
        # COMO FAÇO PARA INTERPOR RECURSO DA DECISÃO?
        adicionar_paragrafo(doc, "COMO FAÇO PARA INTERPOR RECURSO DA DECISÃO?", negrito=True)
        adicionar_paragrafo(doc, "Havendo interesse na interposição de recurso administrativo, este poderá ser interposto no prazo de 20 dias contados do recebimento desta notificação, conforme disposto no art. 9º da RDC nº 266/2019.")
        adicionar_paragrafo(doc, "O protocolo do recurso deverá ser feito exclusivamente, por meio de peticionamento intercorrente no processo indicado no campo assunto desta notificação, pelo Sistema Eletrônico de Informações (SEI). Para tanto, é necessário, primeiramente, fazer o cadastro como usuário externo SEI-Anvisa. Acesse o portal da Anvisa https://www.gov.br/anvisa/pt-br > Sistemas > SEI > Acesso para Usuários Externos (SEI) e siga as orientações. Para maiores informações, consulte o Manual do Usuário Externo Sei-Anvisa, que está disponível em https://www.gov.br/anvisa/pt-br/sistemas/sei.")
        doc.add_paragraph("\n")  # Quebra de linha
        
        # Quais documentos devem acompanhar o recurso
        adicionar_paragrafo(doc, "QUAIS DOCUMENTOS DEVEM ACOMPANHAR O RECURSO?", negrito=True)
        adicionar_paragrafo(doc, "a) Autuado pessoa jurídica:")
        adicionar_paragrafo(doc, "1. Contrato ou estatuto social da empresa, com a última alteração;")
        adicionar_paragrafo(doc, "2. Procuração e documento de identificação do outorgado (advogado ou representante), caso constituído para atuar no processo. Somente serão aceitas procurações e substabelecimentos assinados eletronicamente, com certificação digital no padrão da Infraestrutura de Chaves Públicas Brasileira (ICP-Brasil) ou pelo assinador Gov.br.")
        adicionar_paragrafo(doc, "3. Ata de eleição da atual diretoria quando a procuração estiver assinada por diretor que não conste como sócio da empresa;")
        adicionar_paragrafo(doc, "4. No caso de contestação sobre o porte da empresa considerado para a dosimetria da pena de multa: comprovação do porte econômico referente ao ano em que foi proferida a decisão (documentos previstos no art. 50 da RDC nº 222/2006).")
        adicionar_paragrafo(doc, "b) Autuado pessoa física:")
        adicionar_paragrafo(doc, "1. Documento de identificação do autuado;")
        adicionar_paragrafo(doc, "2. Procuração e documento de identificação do outorgado (advogado ou representante), caso constituído para atuar no processo.")
        doc.add_paragraph("\n")  # Quebra de linha

        # Interface Streamlit
        st.title("Gerador de Documentos - Processos Administrativos")
        processo = st.text_input("Digite o número do processo:")

        uploaded_file = st.file_uploader("Envie o arquivo PDF do processo", type="pdf")
        
        # Fechamento
        advogado_nome = info.get('socios_advogados', ["[Nome não informado]"])
        advogado_nome = advogado_nome[0] if advogado_nome else "[Nome não informado]"
        
        advogado_email = info.get('emails', ["[E-mail não informado]"])
        advogado_email = advogado_email[0] if advogado_email else "[E-mail não informado]"
        
        adicionar_paragrafo(doc, f"Por fim, esclarecemos que foi concedido aos autos por meio do Sistema Eletrônico de Informações (SEI), por 180 (cento e oitenta) dias, ao usuário: {advogado_nome} – E-mail: {advogado_email}")
        adicionar_paragrafo(doc, "Atenciosamente,", negrito=True)

        
        # Salva o documento
        doc.save(output_path)
        print(f"Documento gerado com sucesso: {output_path}")
    except Exception as e:
        print(f"Erro ao gerar o documento DOCX: {e}")

def extract_information(text):
    """
    Extrai informações específicas do texto, como Nome do Autuado, CNPJ/CPF, Sócios/Advogados e E-mails.

    Args:
        text (str): Texto a ser analisado.

    Returns:
        dict: Dicionário com as informações extraídas.
    """
    autuado_pattern = r"(?:NOME AUTUADO|Autuado|Empresa|Razão Social):\s*([\w\s,.-]+)"
    cnpj_cpf_pattern = r"(?:CNPJ|CPF):\s*([\d./-]+)"
    socios_adv_pattern = r"(?:Sócio|Advogado|Responsável|Representante Legal):\s*([\w\s]+)"
    email_pattern = r"(?:E-mail|Email):\s*([\w.-]+@[\w.-]+\.[a-z]{2,})"

    info = {
        "nome_autuado": re.search(autuado_pattern, text).group(1) if re.search(autuado_pattern, text) else None,
        "cnpj_cpf": re.search(cnpj_cpf_pattern, text).group(1) if re.search(cnpj_cpf_pattern, text) else None,
        "socios_advogados": re.findall(socios_adv_pattern, text),
        "emails": re.findall(email_pattern, text),
    }
    return info

def main():
    processo = input("Digite o número do processo: ")
    
    # Busca o processo e faz o download
    buscar_processo(processo)
    baixar_processo()

    # Busca o último arquivo baixado
    diretorio_downloads = os.path.expanduser("~/Downloads")
    pdf_path = buscar_ultimo_arquivo_baixado(diretorio_downloads)  # pdf_path é atribuído aqui

    if pdf_path:  # Verifica se o arquivo foi encontrado
        texto_extraido = extract_text_with_pypdf2(pdf_path)  # Extrai o texto do PDF
        if texto_extraido:  # Verifica se algum texto foi extraído
            print("Texto extraído com sucesso.")

            # Extrair informações do texto
            info = extract_information(texto_extraido)
            if not info:
                print("Nenhuma informação extraída do texto.")
                return  # Para a execução se não houver informações

            # Extrair endereços
            enderecos = extract_addresses(texto_extraido)

            # Gerar o documento com base nas informações extraídas
            print("Gerando documento...")
            gerar_documento_docx(processo, info, enderecos)
        else:
            print("Nenhum texto foi extraído do PDF.")
    else:
        print("Nenhum arquivo foi encontrado.")

if __name__ == "__main__":
    main()
