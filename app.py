# ---------------------------
# Importação de Bibliotecas
# ---------------------------
import spacy
import re
import unicodedata
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt
import os
import streamlit as st

# ---------------------------
# Carregar o modelo SpaCy
# ---------------------------
try:
    nlp = spacy.load("pt_core_news_md")
    st.success("Modelo SpaCy 'pt_core_news_md' carregado com sucesso!")
except OSError:
    nlp = None
    st.error("Erro ao carregar o modelo SpaCy. Certifique-se de que o modelo 'pt_core_news_lg' esteja instalado.")
    st.stop()  # Interrompe a execução do app se o modelo não estiver carregado


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
        st.error(f"Erro ao processar PDF: {e}")
        return ''

# ---------------------------
# Funções de Extração de Dados com SpaCy
# ---------------------------
def extract_information_with_spacy(text):
    doc = nlp(text)
    emails = [ent.text for ent in doc.ents if ent.label_ == "EMAIL"]
    pessoas = [ent.text for ent in doc.ents if ent.label_ == "PER"]
    cnpjs_cpfs = re.findall(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b|\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", text)
    
    return {
        "nome_autuado": pessoas[0] if pessoas else None,
        "cnpj_cpf": cnpjs_cpfs[0] if cnpjs_cpfs else None,
        "socios_advogados": pessoas,
        "emails": emails
    }

def extract_addresses_with_spacy(text):
    doc = nlp(text)
    addresses = []
    for ent in doc.ents:
        if ent.label_ == "LOC":
            addresses.append(ent.text)
    return addresses

# ---------------------------
# Funções de Criação de Documento
# ---------------------------
def adicionar_paragrafo(doc, texto="", negrito=False, tamanho=12):
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return paragrafo

def extract_process_number(file_name):
    base_name = os.path.splitext(file_name)[0]  # Remove a extensão
    if base_name.startswith("SEI"):
        base_name = base_name[3:].strip()  # Remove "SEI"
    return base_name

def gerar_documento_docx(info, enderecos, numero_processo):
    try:
        output_directory = "output"
        os.makedirs(output_directory, exist_ok=True)

        output_path = os.path.join(output_directory, f"Notificacao_Processo_Nº_{numero_processo}.docx")

        doc = Document()
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for endereco in enderecos:
            adicionar_paragrafo(doc, f"Endereço: {endereco}")

        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº: {numero_processo}", negrito=True)

        advogado_nome = info.get('socios_advogados', ["[Nome não informado]"])[0]
        advogado_email = info.get('emails', ["[E-mail não informado]"])[0]

        adicionar_paragrafo(doc, f"Atenciosamente,", negrito=True)
        adicionar_paragrafo(doc, f"{advogado_nome} – E-mail: {advogado_email}")

        doc.save(output_path)

        with open(output_path, "rb") as file:
            st.download_button(
                label="Baixar Documento",
                data=file,
                file_name=f"Notificacao_Processo_Nº_{numero_processo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"Erro ao gerar o documento DOCX: {e}")

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

            info = extract_information_with_spacy(text)
            addresses = extract_addresses_with_spacy(text)

            if st.button("Gerar Documento"):
                gerar_documento_docx(info, addresses, numero_processo)
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
