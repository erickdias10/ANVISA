# Definição de funções acima do bloco principal

# ---------------------------
# Função de Processamento do PDF e Integração com Streamlit
# ---------------------------
def processar_pdf(file):
    try:
        # Extrair texto do PDF
        texto_extraido = extract_text_with_pypdf2(file)

        # Validar se o texto foi extraído com sucesso
        if not texto_extraido:
            raise ValueError("Não foi possível extrair texto do arquivo PDF.")

        # Extrair informações do texto
        info_extraida = extract_information(texto_extraido)

        # Garantir que todas as chaves existem em `info_extraida`
        info_extraida = {
            "nome_autuado": info_extraida.get("nome_autuado", "[Nome não informado]"),
            "cnpj_cpf": info_extraida.get("cnpj_cpf", "[CNPJ/CPF não informado]"),
            "socios_advogados": info_extraida.get("socios_advogados", []),
            "emails": info_extraida.get("emails", [])
        }

        # Extrair endereços
        enderecos = extract_addresses(texto_extraido)

        # Validar endereços extraídos
        if not enderecos:
            enderecos = [{"endereco": "[Endereço não informado]", "cidade": "", "bairro": "", "estado": "", "cep": ""}]

        # Extrair número do processo
        numero_processo = extract_process_number(file.name)

        # Validar o número do processo
        if not numero_processo:
            raise ValueError("Número do processo não identificado no nome do arquivo.")

        # Gerar documento
        docx_path = gerar_documento_docx(info_extraida, enderecos, numero_processo)
        return docx_path
    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
        return None

# ---------------------------
# Interface Streamlit
# ---------------------------
if __name__ == "__main__":
    st.title("Sistema de Extração e Geração de Documentos")
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type=["pdf"])

    if uploaded_file is not None:
        st.write(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")

        # Processar o arquivo PDF
        docx_path = processar_pdf(uploaded_file)

        # Se o documento foi gerado com sucesso, oferece para download
        if docx_path:
            with open(docx_path, "rb") as f:
                st.download_button(
                    label="Baixar Documento Gerado",
                    data=f,
                    file_name=os.path.basename(docx_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.write("Erro ao gerar o documento!")
