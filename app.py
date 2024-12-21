# ---------------------------
# Importação de Bibliotecas
# ---------------------------
import re
from PyPDF2 import PdfReader
import unicodedata
from docx import Document
from docx.shared import Pt
import os
import streamlit as st
import spacy
from typing import ForwardRef

# Corrige temporariamente o erro de ForwardRef
ForwardRef._evaluate = lambda *args, **kwargs: None

# ---------------------------
# Inicialização do SpaCy com Fallback
# ---------------------------
@st.cache_resource
def load_spacy_model():
    try:
        return spacy.load("pt_core_news_lg")
    except OSError:
        st.warning("Modelo 'pt_core_news_lg' não encontrado. Instalando modelo pequeno como alternativa...")
        return spacy.load("pt_core_news_sm")
    except Exception as e:
        st.error(f"Erro ao carregar o modelo SpaCy: {e}")
        return None

# Carrega o modelo
nlp = load_spacy_model()

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
# Extração com SpaCy
# ---------------------------
def extract_information_with_spacy(text):
    doc = nlp(text)
    
    info = {
        "nomes": [ent.text for ent in doc.ents if ent.label_ == "PER"],
        "cnpj_cpf": re.findall(r"(?:CNPJ|CPF):\s*([\d./-]+)", text),
        "emails": [ent.text for ent in doc.ents if ent.label_ == "EMAIL"]
    }
    
    # Extração adicional de padrões não cobertos pelo SpaCy
    socios_adv_pattern = r"(?:Sócio|Advogado|Responsável|Representante Legal):\s*([\w\s]+)"
    info["socios_advogados"] = re.findall(socios_adv_pattern, text)

    return info

def extract_addresses_with_spacy(text):
    doc = nlp(text)
    addresses = []

    for ent in doc.ents:
        if ent.label_ == "LOC":
            addresses.append(ent.text)

    return addresses

def extract_process_number(file_name):
    base_name = os.path.splitext(file_name)[0]  # Remove a extensão
    if base_name.startswith("SEI"):
        base_name = base_name[3:].strip()  # Remove "SEI"
    return base_name

def adicionar_paragrafo(doc, texto="", negrito=False, tamanho=12):
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return paragrafo

# ---------------------------
# Interface Streamlit
# ---------------------------
st.title("Sistema de Extração e Geração de Notificações")

uploaded_file = st.file_uploader("Envie um arquivo PDF", type="pdf")

if uploaded_file:
    try:
        # Certifique-se de que o SpaCy foi carregado
        if nlp is None:
            st.error("O modelo SpaCy não pôde ser carregado. Verifique sua instalação.")
        else:
            file_name = uploaded_file.name
            numero_processo = extract_process_number(file_name)

            # Extrai texto do PDF
            text = extract_text_with_pypdf2(uploaded_file)
            if not text:
                st.error("O texto não pôde ser extraído do PDF. Verifique o arquivo.")
            else:
                st.success(f"Texto extraído com sucesso! Número do processo: {numero_processo}")

                # Extrai informações e endereços
                info = extract_information_with_spacy(text) or {}
                addresses = extract_addresses_with_spacy(text) or []

                # Gera o documento quando o botão for clicado
                if st.button("Gerar Documento"):
                    os.makedirs("output", exist_ok=True)  # Cria o diretório, se não existir

                    gerar_documento_docx(info, addresses, numero_processo)

                    # Verifica se o arquivo foi gerado
                    output_path = os.path.join("output", f"Notificacao_Processo_Nº_{numero_processo}.docx")
                    if os.path.exists(output_path):
                        with open(output_path, "rb") as file:
                            st.download_button(
                                label="Baixar Documento",
                                data=file,
                                file_name=f"Notificacao_Processo_Nº_{numero_processo}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    else:
                        st.error("Erro: O arquivo não foi gerado.")
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
