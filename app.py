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