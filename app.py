import re
import os
import unicodedata
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt
import joblib
import streamlit as st

# ---------------------------
# Modelo
# ---------------------------
VECTOR_PATH = r"C:\Users\erickd\OneDrive - Bem Promotora de Vendas e Servicos SA\Área de Trabalho\Projeto"

def predict_addresses_with_model(text, vectorizer_path="vectorizer.pkl", model_path="address_model.pkl"):
    try:
        vectorizer = joblib.load(vectorizer_path)
        model = joblib.load(model_path)
        text_vectorized = vectorizer.transform([text])
        predictions = model.predict(text_vectorized)
        return predictions
    except Exception as e:
        print(f"Erro ao fazer predição de endereços: {e}")
        return []

def predict_Nome_Email_with_model(text, vectorizer_path="vectorizer_Nome.pkl", model_path="modelo_Nome.pkl"):
    try:
        vectorizer = joblib.load(vectorizer_path)
        model = joblib.load(model_path)
        text_vectorized = vectorizer.transform([text])
        predictions = model.predict(text_vectorized)
        return predictions
    except Exception as e:
        print(f"Erro ao fazer predição de nomes e e-mails: {e}")
        return {}

# ---------------------------
# Funções de Processamento de Texto
# ---------------------------
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
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        print(f"Erro ao processar PDF {pdf_path}: {e}")
        return ''

# ---------------------------
# Funções de Extração de Dados
# ---------------------------
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

        # Filtra endereços com valor 'None' ou 'none'
        if any(value and value.lower() != 'none' for value in address.values()):
            addresses.append(address)

    # Filtra e mantém apenas o endereço mais completo em caso de repetição
    addresses = remove_duplicate_and_incomplete_addresses(addresses)
    
    return addresses or []

def remove_duplicate_and_incomplete_addresses(addresses):
    unique_addresses = []
    seen_addresses = set()

    for address in addresses:
        # Substituir None por uma string vazia
        address_tuple = tuple(sorted((
            address.get('endereco', ''),
            address.get('cidade', ''),
            address.get('bairro', ''),
            address.get('estado', ''),
            address.get('cep', '')
        )))

        # Verifica se o endereço já foi visto antes
        if address_tuple not in seen_addresses:
            seen_addresses.add(address_tuple)
            unique_addresses.append(address)
        else:
            # Verificar se o endereço já existe e, se for o caso, substituí-lo
            existing_address = next(
                (a for a in unique_addresses 
                 if tuple(sorted((
                    a.get('endereco', ''),
                    a.get('cidade', ''),
                    a.get('bairro', ''),
                    a.get('estado', ''),
                    a.get('cep', '')
                ))) == address_tuple), None
            )

            # Substituição com base no comprimento dos dados
            if existing_address:
                if len(address.get('endereco', '')) > len(existing_address.get('endereco', '')): 
                    unique_addresses.remove(existing_address)
                    unique_addresses.append(address)
                elif len(address.get('cidade', '')) > len(existing_address.get('cidade', '')):
                    unique_addresses.remove(existing_address)
                    unique_addresses.append(address)
                elif len(address.get('bairro', '')) > len(existing_address.get('bairro', '')):
                    unique_addresses.remove(existing_address)
                    unique_addresses.append(address)
                elif len(address.get('estado', '')) > len(existing_address.get('estado', '')):
                    unique_addresses.remove(existing_address)
                    unique_addresses.append(address)
                elif len(address.get('cep', '')) > len(existing_address.get('cep', '')):
                    unique_addresses.remove(existing_address)
                    unique_addresses.append(address)

    return unique_addresses

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

def gerar_documento_docx(info, enderecos, numero_processo):
    try:
        output_directory = "output"
        os.makedirs(output_directory, exist_ok=True)
        output_path = os.path.join(output_directory, f"Notificacao_Processo_Nº_{numero_processo}.docx")
        doc = Document()

        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "[Ao Senhor/À Senhora]")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        doc.save(output_path)
        return output_path
    except Exception as e:
        print(f"Erro ao gerar documento: {e}")
        return None

# ---------------------------
# Função de Processamento do PDF e Integração com Streamlit
# ---------------------------
def processar_pdf(file):
    texto_extraido = extract_text_with_pypdf2(file)
    info_extraida = extract_information(texto_extraido)
    enderecos = extract_addresses(texto_extraido)
    numero_processo = extract_process_number(file.name)
    docx_path = gerar_documento_docx(info_extraida, enderecos, numero_processo)
    return docx_path

# Interface Streamlit
st.title("Sistema de Extração e Geração de Documentos")

uploaded_file = st.file_uploader("Escolha um arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    st.write(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")

    # Processar o arquivo PDF
    docx_path = processar_pdf(uploaded_file)

    # Se o documento foi gerado com sucesso, oferece para download
    if docx_path:
        with open(docx_path, "rb") as f:
            st.download_button(
                label="Baixar Documento Gerado",
                data=f,
                file_name=os.path.basename(docx_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.write("Erro ao gerar o documento!")
