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

# ---------------------------
# Modelo SpaCy
# ---------------------------
try:
    nlp = spacy.load("pt_core_news_lg")
except OSError:
    os.system("python -m spacy download pt_core_news_lg")
    nlp = spacy.load("pt_core_news_lg")

def predict_with_spacy(text, entity_label):
    try:
        doc = nlp(text)
        entities = [ent.text for ent in doc.ents if ent.label_ == entity_label]
        return entities
    except Exception as e:
        st.error(f"Erro ao usar SpaCy para {entity_label}: {e}")
        return []

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

def extract_text_with_pypdf2(file):
    try:
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
        return ''

# ---------------------------
# Funções de Extração de Dados
# ---------------------------
def extract_information_with_spacy(text):
    info = {
        "nome_autuado": predict_with_spacy(text, "PER") or None,
        "cnpj_cpf": predict_with_spacy(text, "MISC") or None,
        "socios_advogados": predict_with_spacy(text, "ORG") or [],
        "emails": predict_with_spacy(text, "EMAIL") or [],
    }
    return info

def extract_addresses_with_spacy(text):
    addresses = predict_with_spacy(text, "LOC")
    return [{"endereco": addr} for addr in addresses]

def extract_process_number(file_name):
    base_name = os.path.splitext(file_name)[0]  # Remove a extensão
    if base_name.startswith("SEI"):
        base_name = base_name[3:].strip()  # Remove "SEI"
    return base_name

# ---------------------------
# Função de Geração de Documento
# ---------------------------
def gerar_documento_docx(info, enderecos, numero_processo):
    try:
        # Criação do documento
        doc = Document()

        doc.add_paragraph("\n")
        doc.add_paragraph(f"Processo: {numero_processo}")
        doc.add_paragraph(f"Nome Autuado: {info.get('nome_autuado', '[Não informado]')}")
        doc.add_paragraph(f"Endereço: {enderecos}")

        # Salvar como arquivo temporário
        output_path = f"Notificacao_Processo_Nº_{numero_processo}.docx"
        doc.save(output_path)
        return output_path
    except Exception as e:
        st.error(f"Erro ao gerar o documento DOCX: {e}")
        return None

# ---------------------------
# Interface Streamlit
# ---------------------------
st.title("Sistema de Extração e Geração de Notificações")

uploaded_file = st.file_uploader("Envie um arquivo PDF", type="pdf")

if uploaded_file:
    try:
        file_name = uploaded_file.name
        numero_processo = extract_process_number(file_name)
        text = extract_text_with_pypdf2(uploaded_file)
        if text:
            st.success(f"Texto extraído com sucesso! Número do processo: {numero_processo}")

            info = extract_information_with_spacy(text) or {}
            addresses = extract_addresses_with_spacy(text) or []

            if st.button("Gerar Documento"):
                doc_path = gerar_documento_docx(info, addresses, numero_processo)
                if doc_path:
                    with open(doc_path, "rb") as file:
                        st.download_button(
                            label="Baixar Documento",
                            data=file,
                            file_name=doc_path,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
