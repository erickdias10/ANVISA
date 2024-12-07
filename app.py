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
    Extrai informações específicas do texto.
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


def adicionar_paragrafo(doc, texto, negrito=False):
    """
    Adiciona um parágrafo ao documento DOCX com opções de negrito.
    """
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    if negrito:
        run.bold = True
    paragrafo.style.font.size = Pt(12)


def gerar_documento_docx(info, enderecos):
    """
    Gera um documento DOCX com base nas informações extraídas.
    """
    try:
        doc = Document()
        output_path = "documento_gerado.docx"

        adicionar_paragrafo(doc, "Gerador de Documentos - Processos Administrativos", negrito=True)

        # Adicionar informações ao documento
        adicionar_paragrafo(doc, f"Nome do Autuado: {info.get('nome_autuado', '[Não informado]')}")
        adicionar_paragrafo(doc, f"CNPJ/CPF: {info.get('cnpj_cpf', '[Não informado]')}")
        adicionar_paragrafo(doc, f"Endereços: {', '.join(enderecos) if enderecos else '[Não informado]'}")

        # Fechamento
        advogado_nome = info.get('socios_advogados', ["[Nome não informado]"])
        advogado_nome = advogado_nome[0] if advogado_nome else "[Nome não informado]"

        advogado_email = info.get('emails', ["[E-mail não informado]"])
        advogado_email = advogado_email[0] if advogado_email else "[E-mail não informado]"

        adicionar_paragrafo(doc, f"Por fim, esclarecemos que foi concedido aos autos por meio do Sistema Eletrônico de Informações (SEI), por 180 (cento e oitenta) dias, ao usuário: {advogado_nome} – E-mail: {advogado_email}")
        adicionar_paragrafo(doc, "Atenciosamente,", negrito=True)

        # Salvar o documento no caminho especificado
        doc.save(output_path)
        return output_path
    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")
        return None


# Interface do Streamlit
st.title("Gerador de Documentos - Processos Administrativos")

uploaded_file = st.file_uploader("Envie o arquivo PDF do processo", type="pdf")

if uploaded_file:
    with st.spinner("Processando o arquivo..."):
        texto_extraido = extract_text_with_pypdf2(uploaded_file)
        if texto_extraido:
            info = extract_information(texto_extraido)
            enderecos = []  # Corrigir função extract_addresses se necessário
            if info:
                st.success("Informações extraídas com sucesso!")

                # Exibir resultados na interface
                st.write("Informações Extraídas:")
                st.write(info)
                st.write("Endereços Extraídos:")
                st.write(enderecos)

                # Gerar documento DOCX
                output_path = gerar_documento_docx(info, enderecos)
                if output_path:
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="Baixar Documento Gerado",
                            data=file,
                            file_name=os.path.basename(output_path),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error("Erro ao gerar o documento.")
            else:
                st.error("Informações ou endereços não extraídos corretamente.")
        else:
            st.error("Nenhum texto foi extraído do arquivo.")