# ---------------------------
# Importação de Bibliotecas
# ---------------------------
import re
from PyPDF2 import PdfReader
import unicodedata
from docx import Document
from docx.shared import Pt
import os
import joblib
import streamlit as st

# ---------------------------
# Funções Auxiliares
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

def extract_information(text):
    autuado_pattern = r"(?:NOME AUTUADO|Autuado|Empresa|Razão Social):\s*([\w\s,.-]+)"
    cnpj_cpf_pattern = r"(?:CNPJ|CPF):\s*([\d./-]+)"
    socios_adv_pattern = r"(?:Sócio|Advogado|Responsável|Representante Legal):\s*([\w\s]+)"
    email_pattern = r"(?:E-mail|Email):\s*([\w.-]+@[\w.-]+\.[a-z]{2,})"

    info = {
        "nome_autuado": re.search(autuado_pattern, text).group(1) if re.search(autuado_pattern, text) else None,
        "cnpj_cpf": re.search(cnpj_cpf_pattern, text).group(1) if re.search(cnpj_cpf_pattern, text) else None,
        "socios_advogados": re.findall(socios_adv_pattern, text) or [],
        "emails": re.findall(email_pattern, text) or [],
    }
    return info

def extract_addresses(text):
    endereco_pattern = r"(?:Endereço|End|Endereco):\s*([\w\s.,ºª-]+)"
    cidade_pattern = r"Cidade:\s*([\w\s]+(?: DE [\w\s]+)?)"
    bairro_pattern = r"Bairro:\s*([\w\s]+)"
    estado_pattern = r"Estado:\s*([A-Z]{2})"
    cep_pattern = r"CEP:\s*(\d{2}\.\d{3}-\d{3}|\d{5}-\d{3})"

    addresses = []
    seen_addresses = set()

    endereco_matches = re.findall(endereco_pattern, text)
    cidade_matches = re.findall(cidade_pattern, text)
    bairro_matches = re.findall(bairro_pattern, text)
    estado_matches = re.findall(estado_pattern, text)
    cep_matches = re.findall(cep_pattern, text)

    for i in range(max(len(endereco_matches), len(cidade_matches), len(bairro_matches), len(estado_matches), len(cep_matches))):
        endereco = endereco_matches[i].strip() if i < len(endereco_matches) else None
        cidade = cidade_matches[i].strip() if i < len(cidade_matches) else None
        bairro = bairro_matches[i].strip() if i < len(bairro_matches) else None
        estado = estado_matches[i].strip() if i < len(estado_matches) else None
        cep = cep_matches[i].strip() if i < len(cep_matches) else None

        if endereco and endereco not in seen_addresses:
            seen_addresses.add(endereco)
            addresses.append({
                "endereco": endereco,
                "cidade": cidade,
                "bairro": bairro,
                "estado": estado,
                "cep": cep
            })

    return addresses or []

def adicionar_paragrafo(doc, texto="", negrito=False, tamanho=12):
    paragrafo = doc.add_paragraph()
    run = paragrafo.add_run(texto)
    run.bold = negrito
    run.font.size = Pt(tamanho)
    return paragrafo

def extract_process_number(file_name):
    base_name = os.path.splitext(file_name)[0]
    if base_name.startswith("SEI"):
        base_name = base_name[3:].strip()
    return base_name

# ---------------------------
# Função de Geração de Documento
# ---------------------------
def gerar_documento_streamlit(info, enderecos, numero_processo):
    """
    Gera um documento DOCX com informações do processo e endereços extraídos.

    Args:
        info (dict): Dicionário com informações extraídas do texto.
        enderecos (list): Lista de dicionários contendo informações de endereços.
        numero_processo (str): Número do processo extraído do nome do arquivo.

    Returns:
        str: Caminho do arquivo gerado.
    """
    try:
        # Diretório seguro para salvar arquivos
        output_directory = "output"
        os.makedirs(output_directory, exist_ok=True)

        # Caminho completo do arquivo
        output_path = os.path.join(output_directory, f"Notificacao_Processo_Nº_{numero_processo}.docx")


        
        # Criação do documento
        doc = Document()

        doc.add_paragraph("\n")
        adicionar_paragrafo(doc, "Ao(a) Senhor(a):")
        adicionar_paragrafo(doc, f"{info.get('nome_autuado', '[Nome não informado]')} – CNPJ/CPF: {info.get('cnpj_cpf', '[CNPJ/CPF não informado]')}")
        doc.add_paragraph("\n")

        # Adiciona endereços
        for idx, endereco in enumerate(enderecos, start=1):
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Cidade: {endereco.get('cidade', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Bairro: {endereco.get('bairro', '[Não informado]')}")
            adicionar_paragrafo(doc, f"Estado: {endereco.get('estado', '[Não informado]')}")
            adicionar_paragrafo(doc, f"CEP: {endereco.get('cep', '[Não informado]')}")
            doc.add_paragraph("\n")

        # Corpo principal
            # Corpo principal
        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância proferida pela Coordenação de Atuação Administrativa e Julgamento das Infrações Sanitárias.", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo Sancionador nº: {numero_processo} ", negrito=True)
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

        # Interface Streamlit
        uploaded_file = st.file_uploader("Envie o arquivo PDF do processo", type="pdf")
        
        # Fechamento
        advogado_nome = info.get('socios_advogados', ["[Nome não informado]"])
        advogado_nome = advogado_nome[0] if advogado_nome else "[Nome não informado]"
        
        advogado_email = info.get('emails', ["[E-mail não informado]"])
        advogado_email = advogado_email[0] if advogado_email else "[E-mail não informado]"
        
        adicionar_paragrafo(doc, f"Por fim, esclarecemos que foi concedido aos autos por meio do Sistema Eletrônico de Informações (SEI), por 180 (cento e oitenta) dias, ao usuário: {advogado_nome} – E-mail: {advogado_email}")
        adicionar_paragrafo(doc, "Atenciosamente,", negrito=True)


        # Salva o documento
        doc.save(output_path)

        # Botão de download no Streamlit
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

            info = extract_information(text)
            addresses = extract_addresses(text)

            doc = gerar_documento_streamlit(info, addresses, numero_processo)

            if doc:
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="Baixar Notificação Gerada",
                    data=buffer,
                    file_name=f"Notificacao_Processo_Nº_{numero_processo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
