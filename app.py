# Bloco 1: Importação de Bibliotecas
import re
import os
import unicodedata
from docx import Document
from docx.shared import Pt
from PyPDF2 import PdfReader
import streamlit as st

# Bloco 2: Funções de Manipulação de Arquivos e Extração de Texto
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

def extract_text_with_pypdf2(pdf_file):
    """
    Extrai texto de PDFs usando PyPDF2.
    """
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
        return ''

# Bloco 3: Processamento de Endereços e Formatação do Documento
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

def gerar_documento_docx(process_number, enderecos, output_path="Notificacao_Processo.docx"):
    """
    Gera um documento DOCX com informações do processo e endereços extraídos.

    Args:
        process_number (str): Número do processo administrativo.
        enderecos (list): Lista de dicionários contendo informações de endereços.
        output_path (str): Caminho para salvar o documento gerado.
    """
    try:
        doc = Document()

        # Adiciona informações do processo e endereços
        adicionar_paragrafo(doc, "[Ao Senhor/À Senhora]")
        adicionar_paragrafo(doc, "NOME AUTUADO – CNPJ/CPF: [XXXXX]")
        doc.add_paragraph("\n")

        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        # Corpo principal do texto
        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº {process_number}", negrito=True)
        adicionar_paragrafo(doc, "Prezado(a) Senhor(a),")
        adicionar_paragrafo(doc, "Informamos que foi proferido julgamento pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias no processo administrativo sancionador em referência, conforme decisão em anexo.")

        # Salva o documento
        doc.save(output_path)
        st.success(f"Documento gerado com sucesso: {output_path}")
    except Exception as e:
        st.error(f"Erro ao gerar o documento DOCX: {e}")

# Bloco 4: Interface com Streamlit
def main():
    st.title("Gerador de Documentos - Processos Administrativos")

    uploaded_file = st.file_uploader("Envie o arquivo PDF do processo", type="pdf")
    if uploaded_file:
        with st.spinner("Processando o arquivo..."):
            texto_extraido = extract_text_with_pypdf2(uploaded_file)
            if texto_extraido:
                process_number = "12345"  # Exemplo de número de processo
                enderecos = extract_addresses(texto_extraido)
                output_path = f"Notificacao_Processo_{process_number}.docx"
                gerar_documento_docx(process_number, enderecos, output_path)
            else:
                st.error("Nenhum texto foi extraído do PDF.")
    else:
        st.info("Envie um arquivo PDF para começar.")

if __name__ == "__main__":
    main()