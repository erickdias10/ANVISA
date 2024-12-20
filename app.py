import re
import os
import unicodedata
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt
import streamlit as st
from pdf2image import convert_from_bytes
import pytesseract

# ---------------------------
# Funções de Tratamento de Texto
# ---------------------------
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)  # Remove múltiplos espaços
    return text.strip()

def corrigir_texto(texto):
    substituicoes = {
        'Ã©': 'é', 'Ã§Ã£o': 'ção', 'Ã³': 'ó', 'Ã': 'à',
    }
    for errado, correto in substituicoes.items():
        texto = texto.replace(errado, correto)
    return texto

# ---------------------------
# Extração de Texto PDF
# ---------------------------
def extract_text_with_pypdf2(pdf_file):
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text() or ""
            text += extracted
        return corrigir_texto(normalize_text(text))
    except Exception as e:
        st.error(f"Erro ao extrair texto do PDF: {e}")
        return ""

def extract_text_with_ocr(pdf_file):
    try:
        images = convert_from_bytes(pdf_file.read())
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image, lang='por')
        return corrigir_texto(normalize_text(text))
    except Exception as e:
        st.error(f"Erro ao realizar OCR no PDF: {e}")
        return ""

def extract_text_with_fallback(pdf_file):
    text = extract_text_with_pypdf2(pdf_file)
    if not text:
        st.warning("Tentando extração via OCR, pois o PDF parece não conter texto extraível.")
        text = extract_text_with_ocr(pdf_file)
    return text

def get_process_number(uploaded_file):
    """
    Extrai o número do processo de um texto no formato específico.
    """
    texto = extract_text_with_fallback(uploaded_file)
    # Ajustar a expressão regular para o formato esperado do número de processo
    process_number_pattern = r"Processo(?: Administrativo Sancionador)?[^\d]*?n[ºo]:? (\d+)"
    match = re.search(process_number_pattern, texto)
    if match:
        return match.group(1)
    else:
        return "[Número de processo não encontrado]"

# ---------------------------
# Extração de Informações
# ---------------------------
def extract_information(text):
    autuado_pattern = r"(?:NOME AUTUADO|Autuado|Empresa|Razão Social):\s*([\w\s,.-]+)"
    cnpj_cpf_pattern = r"(?:CNPJ|CPF):\s*([\d./-]+)"

    info = {
        "nome_autuado": re.search(autuado_pattern, text).group(1) if re.search(autuado_pattern, text) else None,
        "cnpj_cpf": re.search(cnpj_cpf_pattern, text).group(1) if re.search(cnpj_cpf_pattern, text) else None
    }
    return info

def extract_addresses(text):
    endereco_pattern = (
        r"Endereço[:]? (.*?)\\n.*?CEP[:]? (\d{5}-\d{3}).*?Cidade[:]? ([\\w\\s]+).*?Estado[:]? ([A-Z]{2})"
    )
    matches = re.finditer(endereco_pattern, text, re.DOTALL)
    enderecos = []
    for match in matches:
        enderecos.append({
            "endereco": match.group(1).strip(),
            "cep": match.group(2),
            "cidade": match.group(3).strip(),
            "estado": match.group(4)
        })
    return enderecos if enderecos else [{"endereco": "[Endereço não encontrado]"}]

# ---------------------------
# Criação do Documento DOCX
# ---------------------------
def adicionar_paragrafo(doc, texto, negrito=False):
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(12)

# ---------------------------
# Criação do Documento DOCX
# ---------------------------
def gerar_documento_docx(process_number, info, enderecos):
    try:
        diretorio_downloads = os.path.expanduser("~/Downloads")
        output_path = os.path.join(diretorio_downloads, f"Notificacao_Processo_Nº_{process_number}.docx")
        doc = Document()

        adicionar_paragrafo(doc, "[Ao Senhor/À Senhora]")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        for endereco in enderecos:
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

                # Corpo principal
            # Corpo principal
        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº {process_number}", negrito=True)
        doc.add_paragraph("\n")  # Quebra de linha
        adicionar_paragrafo(doc, "Prezado(a) Senhor(a),")
        doc.add_paragraph("\n")  # Quebra de linha
        adicionar_paragrafo(doc, "Informamos que foi proferido julgamento pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias no processo administrativo sancionador em referência, conforme decisão em anexo.")
        doc.add_paragraph("\n")  # Quebra de linha
        
        # O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?
        adicionar_paragrafo(doc, "O QUE FAZER SE A DECISÃO TIVER APLICADO MULTA?", negrito=True)
        adicionar_paragrafo(doc, "Sendo aplicada a penalidade de multa, esta notificação estará acompanhada de boleto bancário, que deverá ser pago até o vencimento.")
        adicionar_paragrafo(doc, "O valor da multa poderá ser pago com 20% de desconto caso seja efetuado em até 20 dias contados de seu recebimento. Incorrerá em ilegalidade o usufruto do desconto em data posterior ao prazo referido, mesmo que a data impressa no boleto permita pagamento, sendo a diferença cobrada posteriormente pela Gerência de Gestão de Arrecadação (GEGAR). O pagamento da multa implica em desistência tácita do recurso, conforme art. 21 da Lei nº 6.437/1977.")
        adicionar_paragrafo(doc, "O não pagamento do boleto sem que haja interposição de recurso, acarretará, sucessivamente: i) a inscrição do devedor no Cadastro Informativo de Crédito não Quitado do Setor Público Federal (CADIN); ii) a inscrição do débito em dívida ativa da União; iii) o ajuizamento de ação de execução fiscal contra o devedor; e iv) a comunicação aos cartórios de registros de imóveis, dos devedores inscritos em dívida ativa ou execução fiscal.")
        adicionar_paragrafo(doc, "Esclarecemos que o valor da multa foi atualizado pela taxa Selic acumulada nos termos do art. 37-A da Lei 10.522/2002 e no art. 5º do Decreto-Lei 1.736/79.")
        doc.add_paragraph("\n")  # Quebra de linha
        
        # COMO FAÇO PARA INTERPOR RECURSO DA DECISÃO?
        adicionar_paragrafo(doc, "COMO FAÇO PARA INTERPOR RECURSO DA DECISÃO?", negrito=True)
        adicionar_paragrafo(doc, "Havendo interesse na interposição de recurso administrativo, este poderá ser interposto no prazo de 20 dias contados do recebimento desta notificação, conforme disposto no art. 9º da RDC nº 266/2019.")
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

        doc.save(output_path)
        return output_path
    except Exception as e:
        st.error(f"Erro ao gerar documento: {e}")
        return None
# ---------------------------
# Processamento do PDF
# ---------------------------
def processar_pdf(uploaded_file):
    texto = extract_text_with_fallback(uploaded_file)
    if not texto or texto.isspace():
        st.error("Nenhum texto foi extraído do PDF. O arquivo pode estar corrompido ou baseado em imagem.")
        return None

    st.write("Texto extraído (pré-processado):")
    st.code(texto[:1000])

    info = extract_information(texto) or {"nome_autuado": "[Não informado]", "cnpj_cpf": "[Não informado]"}
    enderecos = extract_addresses(texto) or [{"endereco": "[Endereço não encontrado]"}]

    process_number = get_process_number(uploaded_file)
    st.write(f"Número do processo: {process_number}")

    docx_path = gerar_documento_docx(process_number, info, enderecos)
    if docx_path:
        return docx_path
    else:
        st.error("Falha ao gerar o documento. Verifique os dados extraídos.")
        return None

# ---------------------------
# Interface Streamlit
# ---------------------------
if __name__ == "__main__":
    st.title("Extração e Geração de Documentos")
    uploaded_file = st.file_uploader("Carregue um PDF", type=["pdf"])

    if uploaded_file is not None:
        st.write(f"Processando o arquivo '{uploaded_file.name}'...")
        docx_path = processar_pdf(uploaded_file)

        if docx_path:
            with open(docx_path, "rb") as f:
                st.download_button(
                    label="Baixar Documento Gerado",
                    data=f,
                    file_name=os.path.basename(docx_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Não foi possível gerar o documento. Verifique o PDF.")
