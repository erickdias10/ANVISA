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

        # Salvar o documento no caminho especificado
        doc.save(output_path)
        return output_path
    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")
        return None

   adicionar_paragrafo(doc, "b) Autuado pessoa física:")
        adicionar_paragrafo(doc, "1. Documento de identificação do autuado;")
        adicionar_paragrafo(doc, "2. Procuração e documento de identificação do outorgado (advogado ou representante), caso constituído para atuar no processo.")
        doc.add_paragraph("\n")  # Quebra de linha
        
        # Fechamento
        adicionar_paragrafo(doc, "Por fim, esclarecemos que foi concedido aos autos por meio do Sistema Eletrônico de Informações (SEI), por 180 (cento e oitenta) dias, ao usuário: [nome e e-mail.]")
        adicionar_paragrafo(doc, "Atenciosamente,", negrito=True)
        
        # Salva o documento
        doc.save(output_path)
        print(f"Documento gerado com sucesso: {output_path}")
    except Exception as e:
        print(f"Erro ao gerar o documento DOCX: {e}")


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

