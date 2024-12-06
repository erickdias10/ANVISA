# Bloco 1: Importação de Bibliotecas
import os
import re
import unicodedata
import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt


# Funções Auxiliares
def normalize_text(text):
    """
    Remove caracteres especiais e normaliza o texto.
    """
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    return re.sub(r"\s{2,}", " ", text).strip()


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


def extract_text_with_pypdf2(pdf_file):
    """
    Extrai texto de PDFs usando PyPDF2.
    """
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return corrigir_texto(normalize_text(text))
    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
        return ''


def extract_information(text):
    """
    Extrai informações específicas do texto, como Nome do Autuado, CNPJ/CPF, Sócios/Advogados e E-mails.
    """
    try:
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
    except Exception as e:
        st.error(f"Erro ao extrair informações do texto: {e}")
        return {}


def extract_addresses(text):
    """
    Extrai informações de endereço do texto usando expressões regulares.
    """
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

    addresses = []
    for i in range(max(len(endereco_matches), len(cidade_matches), len(bairro_matches), len(estado_matches), len(cep_matches))):
        address = {
            "endereco": endereco_matches[i].strip() if i < len(endereco_matches) else None,
            "cidade": cidade_matches[i].strip() if i < len(cidade_matches) else None,
            "bairro": bairro_matches[i].strip() if i < len(bairro_matches) else None,
            "estado": estado_matches[i].strip() if i < len(estado_matches) else None,
            "cep": cep_matches[i].strip() if i < len(cep_matches) else None,
        }
        addresses.append(address)

    return addresses


def adicionar_paragrafo(doc, texto, negrito=False):
    """
    Adiciona um parágrafo ao documento.
    """
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(12)


def gerar_documento_docx(info, enderecos):
    """
    Gera um documento DOCX com informações do processo e endereços extraídos.

    Args:
        info (dict): Dicionário com informações extraídas do texto.
        enderecos (list): Lista de dicionários contendo informações de endereços.

    Returns:
        str: Caminho do arquivo gerado.
    """
    try:
        output_path = f"Notificacao_Processo_{info.get('nome_autuado', 'Desconhecido')}.docx"
        doc = Document()

        adicionar_paragrafo(doc, "[Ao Senhor/À Senhora]")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        if enderecos:
            for endereco in enderecos:
                doc.add_paragraph(f"Endereço: {endereco.get('endereco', '[Não informado]')}")
                doc.add_paragraph(f"Cidade: {endereco.get('cidade', '[Não informado]')}")
                doc.add_paragraph(f"Bairro: {endereco.get('bairro', '[Não informado]')}")
                doc.add_paragraph(f"Estado: {endereco.get('estado', '[Não informado]')}")
                doc.add_paragraph(f"CEP: {endereco.get('cep', '[Não informado]')}")
                doc.add_paragraph("\n")

        # Adiciona um fechamento básico
        adicionar_paragrafo(doc, "Atenciosamente,")
        doc.save(output_path)
        return output_path
    except Exception as e:
        st.error(f"Erro ao gerar o documento DOCX: {e}")
        return None


# Interface do Streamlit
st.title("Gerador de Documentos - Processos Administrativos")

uploaded_file = st.file_uploader("Envie o arquivo PDF do processo", type="pdf")

if uploaded_file:
    with st.spinner("Processando o arquivo..."):
        texto_extraido = extract_text_with_pypdf2(uploaded_file)
        if texto_extraido:
            info = extract_information(texto_extraido)
            enderecos = extract_addresses(texto_extraido)
            if info and enderecos:
                output_path = gerar_documento_docx(info, enderecos)
                if output_path:
                    with open(output_path, "rb") as file:
                        st.download_button("Baixar Documento Gerado", file, file_name=output_path)
                else:
                    st.error("Erro ao gerar o documento.")
            else:
                st.error("Informações ou endereços não extraídos corretamente.")
        else:
            st.error("Nenhum texto foi extraído do arquivo.")
