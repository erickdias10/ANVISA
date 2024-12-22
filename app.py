# ---------------------------
# Carregando o Modelo SpaCy
# ---------------------------
@st.cache_resource
def load_spacy_model():
    try:
        return spacy.load("pt_core_news_lg")
    except Exception as e:
        st.error(f"Erro ao carregar o modelo SpaCy: {e}")
        return None

nlp = load_spacy_model()

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
        print(f"Erro ao processar PDF {pdf_path}: {e}")
        return ''

# ---------------------------
# Funções de Extração com SpaCy
# ---------------------------
def extract_entities_with_spacy(text, entity_types):
    if nlp is None:
        return []

    doc = nlp(text)
    entities = {entity_type: [] for entity_type in entity_types}

    for ent in doc.ents:
        if ent.label_ in entity_types:
            entities[ent.label_].append(ent.text)

    return entities

def extract_information_with_spacy(text):
    entity_types = ["PERSON", "ORG", "EMAIL", "LOC", "GPE"]
    entities = extract_entities_with_spacy(text, entity_types)

    info = {
        "nome_autuado": entities.get("PERSON", [None])[0],
        "cnpj_cpf": None,  # CPF/CNPJ não é extraído diretamente por SpaCy
        "socios_advogados": entities.get("ORG", []),
        "emails": entities.get("EMAIL", [])
    }
    return info

# ---------------------------
# Função de Extração de Endereços
# ---------------------------
def extract_addresses_with_spacy(text):
    if nlp is None:
        return []

    doc = nlp(text)
    addresses = []
    seen_addresses = set()

    for ent in doc.ents:
        if ent.label_ in ["LOC", "GPE"]:  # Localidades e regiões geográficas
            address = ent.text.strip()
            if address and address not in seen_addresses:
                seen_addresses.add(address)
                addresses.append({"endereco": address})

    return addresses

# ---------------------------
# Função de Geração de Documento
# ---------------------------
def gerar_documento_docx(info, enderecos, numero_processo):
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
        for endereco in enderecos:
            adicionar_paragrafo(doc, f"Endereço: {endereco.get('endereco', '[Não informado]')}")
            doc.add_paragraph("\n")

        # Corpo principal
        adicionar_paragrafo(doc, "Assunto: Decisão de 1ª instância", negrito=True)
        adicionar_paragrafo(doc, f"Referência: Processo Administrativo nº: {numero_processo} ", negrito=True)
        doc.add_paragraph("\n")

        # Fechamento
        advogado_nome = info.get('socios_advogados', ["[Nome não informado]"])
        advogado_nome = advogado_nome[0] if advogado_nome else "[Nome não informado]"

        advogado_email = info.get('emails', ["[E-mail não informado]"])
        advogado_email = advogado_email[0] if advogado_email else "[E-mail não informado]"

        adicionar_paragrafo(doc, f"Por fim, esclarecemos que foi concedido ao usuário: {advogado_nome} – E-mail: {advogado_email}")
        adicionar_paragrafo(doc, "Atenciosamente", negrito=True)

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
st.title("Sistema de Extração e Geração de Notificações com SpaCy")

uploaded_file = st.file_uploader("Envie um arquivo PDF", type="pdf")

if uploaded_file:
    try:
        # Extrai o número do processo a partir do nome do arquivo
        file_name = uploaded_file.name
        numero_processo = extract_process_number(file_name)

        # Extrai o texto do PDF
        text = extract_text_with_pypdf2(uploaded_file)
        if text:
            st.success(f"Texto extraído com sucesso! Número do processo: {numero_processo}")

            # Extrai informações e endereços
            info = extract_information_with_spacy(text) or {}
            addresses = extract_addresses_with_spacy(text) or []

            # Gera o documento ao clicar no botão
            if st.button("Gerar Documento"):
                gerar_documento_docx(info, addresses, numero_processo)
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
