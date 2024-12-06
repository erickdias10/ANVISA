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


def extract_addresses(text):
    """
    Extrai informações de endereço do texto usando expressões regulares.
    """
    try:
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
                "endereco": endereco_matches[i].strip() if i < len(endereco_matches) else "[Não informado]",
                "cidade": cidade_matches[i].strip() if i < len(cidade_matches) else "[Não informado]",
                "bairro": bairro_matches[i].strip() if i < len(bairro_matches) else "[Não informado]",
                "estado": estado_matches[i].strip() if i < len(estado_matches) else "[Não informado]",
                "cep": cep_matches[i].strip() if i < len(cep_matches) else "[Não informado]",
            }
            addresses.append(address)

        return addresses
    except Exception as e:
        st.error(f"Erro ao extrair endereços: {e}")
        return []


def adicionar_paragrafo(doc, texto, negrito=False):
    """
    Adiciona um parágrafo ao documento.
    """
    try:
        paragrafo = doc.add_paragraph()
        run = paragrafo.add_run(texto)
        run.bold = negrito
        run.font.size = Pt(12)
    except Exception as e:
        raise RuntimeError(f"Erro ao adicionar parágrafo: {e}")


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
        # Verifica o número do processo
        process_number = info.get("process_number", "0000")  # Use um número padrão se não fornecido

        # Define o caminho de saída
        diretorio_downloads = os.getcwd()  # Altere para o diretório desejado
        output_path = os.path.join(diretorio_downloads, f"Notificacao_Processo_Nº_{process_number}.docx")

        # Cria o documento
        doc = Document()

        # Cabeçalho e informações do autuado
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "[Ao Senhor/À Senhora]")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        # Filtra endereços válidos
        enderecos_validos = [
            endereco for endereco in enderecos
            if any(endereco.get(campo) != "[Não informado]" for campo in ["endereco", "cidade", "bairro", "estado", "cep"])
        ]

        if enderecos_validos:
            adicionar_paragrafo(doc, "Endereços:")
            for endereco in enderecos_validos:
                adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
                adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
                adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
                adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
                adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
                doc.add_paragraph("\n")
        else:
            adicionar_paragrafo(doc, "Nenhum endereço válido encontrado.")

        # Texto principal do documento
        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº {process_number}", negrito=True)
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Prezado(a) Senhor(a),")
        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Informamos que foi proferido julgamento pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias no processo administrativo sancionador em referência, conforme decisão em anexo.")
        doc.add_paragraph("\n")

        # Informações sobre multa
        adicionar_paragrafo(doc, "O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?", negrito=True)
        adicionar_paragrafo(doc, "Sendo aplicada a penalidade de multa, esta notificação estará acompanhada de boleto bancário, que deverá ser pago até o vencimento.")
        # (Demais textos permanecem conforme fornecidos)

        # Recursos e anexos
        adicionar_paragrafo(doc, "COMO FAÇO PARA INTERPOR RECURSO DA DECISÃO?", negrito=True)
        adicionar_paragrafo(doc, "Havendo interesse na interposição de recurso administrativo, este poderá ser interposto no prazo de 20 dias contados do recebimento desta notificação.")
        # (Continue com o mesmo padrão)

        # Salvar o documento
        doc.save(output_path)
        return output_path
    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")
        return None


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

