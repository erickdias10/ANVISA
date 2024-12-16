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
VECTOR_PATH = r"C:\\Users\\erickd\\OneDrive - Bem Promotora de Vendas e Servicos SA\\Área de Trabalho\\Projeto"

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
        if address["endereco"] and address["endereco"].lower() != 'none':
            addresses.append(address)

    # Filtra e mantém apenas o endereço mais completo em caso de repetição
    addresses = remove_duplicate_and_incomplete_addresses(addresses)
    
    return addresses or []

def remove_duplicate_and_incomplete_addresses(addresses):
    unique_addresses = []
    seen_addresses = set()

    for address in addresses:
        # Garantir que os valores sejam strings antes de chamar .lower()
        endereco = str(address.get('endereco', '')).lower()
        cidade = str(address.get('cidade', '')).lower()
        bairro = str(address.get('bairro', '')).lower()
        estado = str(address.get('estado', '')).lower()
        cep = str(address.get('cep', '')).lower()

        # Criar um tuple para verificação de duplicados
        address_tuple = (endereco, cidade, bairro, estado, cep)

        # Verifica se o endereço já foi visto antes
        if address_tuple not in seen_addresses:
            seen_addresses.add(address_tuple)
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

# ---------------------------
# Função de Processamento do PDF e Integração com Streamlit
# ---------------------------
def processar_pdf(file):
    try:
        # Extrair texto do PDF
        texto_extraido = extract_text_with_pypdf2(file)

        # Validar se o texto foi extraído com sucesso
        if not texto_extraido:
            raise ValueError("Não foi possível extrair texto do arquivo PDF.")

        # Extrair informações do texto
        info_extraida = extract_information(texto_extraido)

        # Garantir que todas as chaves existem em `info_extraida`
        info_extraida = {
            "nome_autuado": info_extraida.get("nome_autuado", "[Nome não informado]"),
            "cnpj_cpf": info_extraida.get("cnpj_cpf", "[CNPJ/CPF não informado]"),
            "socios_advogados": info_extraida.get("socios_advogados", []),
            "emails": info_extraida.get("emails", [])
        }

        # Extrair endereços
        enderecos = extract_addresses(texto_extraido)

        # Validar endereços extraídos
        if not enderecos:
            enderecos = [{"endereco": "[Endereço não informado]", "cidade": "", "bairro": "", "estado": "", "cep": ""}]

        # Extrair número do processo
        numero_processo = extract_process_number(file.name)

        # Validar o número do processo
        if not numero_processo:
            raise ValueError("Número do processo não identificado no nome do arquivo.")

        # Gerar documento
        docx_path = gerar_documento_docx(info_extraida, enderecos, numero_processo)
        return docx_path
    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
        return None

# ---------------------------
# Interface Streamlit
# ---------------------------
if __name__ == "__main__":
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
