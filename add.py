import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import tempfile
import os
import zipfile
import io
from datetime import datetime

# Função para formatar o documento
def formatar_documento(doc):
    # Estilo para todo o documento
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    
    # Configurar margens (em cm)
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
    
    return doc

# Função para criar um documento Word (Ofício)
def criar_oficio_word(nome, endereco, telefone, processo, assinatura, tipo, conteudo_personalizado):
    doc = Document()
    doc = formatar_documento(doc)
    
    # Data atual
    data_atual = datetime.now().strftime("%d/%m/%Y")
    
    # Cabeçalho do documento
    header = doc.add_paragraph()
    header.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    header.add_run(f"São Paulo, {data_atual}")
    doc.add_paragraph("\n")
    
    # Referência do ofício
    ref_paragraph = doc.add_paragraph()
    ref_run = ref_paragraph.add_run(f"OFÍCIO Nº {tipo}/{datetime.now().year}")
    ref_run.bold = True
    doc.add_paragraph("\n")
    
    # Destinatário
    doc.add_paragraph(f"Ao Sr(a).")
    doc.add_paragraph(f"{nome}")
    doc.add_paragraph(f"{endereco}")
    doc.add_paragraph(f"Tel: {telefone}")
    doc.add_paragraph("\n")
    
    # Assunto
    assunto_paragraph = doc.add_paragraph()
    assunto_run = assunto_paragraph.add_run(f"Assunto: Referente ao processo nº {processo}")
    assunto_run.bold = True
    doc.add_paragraph("\n")
    
    # Vocativo
    doc.add_paragraph("Prezado(a) Senhor(a),")
    doc.add_paragraph("\n")
    
    # Conteúdo do Ofício (personalizado por tipo)
    doc.add_paragraph(conteudo_personalizado)
    doc.add_paragraph("\n")
    
    # Fechamento
    doc.add_paragraph("Atenciosamente,")
    doc.add_paragraph("\n\n")
    despedida = doc.add_paragraph()
    despedida.add_run(assinatura).bold = True
    
    # Rodapé
    footer = doc.add_paragraph()
    footer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    footer_run = footer.add_run("Este é um documento oficial. Favor manter em seus registros.")
    footer_run.italic = True
    
    # Criar arquivo temporário
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    
    return temp_file.name

# Função para criar um arquivo ZIP com os ofícios
def criar_zip_oficios(arquivos):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for i, (nome_arquivo, caminho_arquivo) in enumerate(arquivos.items(), 1):
            # Adicionar cada arquivo ao ZIP
            with open(caminho_arquivo, "rb") as f:
                zip_file.writestr(f"Oficio_{i}.docx", f.read())
    
    zip_buffer.seek(0)
    return zip_buffer

# Interface do Streamlit
st.set_page_config(page_title="Gerador de Ofícios", layout="wide")

st.title("Gerador de Ofícios Automáticos")
st.write("Preencha os dados abaixo para gerar três ofícios para destinatários diferentes.")

# Número do processo comum a todos os ofícios
processo = st.text_input("Número do Processo (comum a todos os ofícios)")
assinatura = st.text_area("Assinatura (Nome e cargo do responsável)")

# Criar abas para cada ofício
tab1, tab2, tab3 = st.tabs(["Ofício 1", "Ofício 2", "Ofício 3"])

# Dicionário para armazenar os dados de cada ofício
dados_oficios = {}

with tab1:
    st.subheader("Dados do Ofício 1")
    with st.form("dados_oficio_1"):
        nome_1 = st.text_input("Nome do Destinatário 1")
        endereco_1 = st.text_input("Endereço 1")
        telefone_1 = st.text_input("Telefone 1")
        conteudo_1 = st.text_area("Conteúdo do Ofício 1", 
                                  "Venho por meio deste comunicar sobre o andamento do processo mencionado. Solicito seu comparecimento em nosso escritório para tratar de assuntos relacionados ao mesmo.")
        submit_1 = st.form_submit_button("Salvar Dados do Ofício 1")
    
    if submit_1:
        dados_oficios["oficio_1"] = {
            "nome": nome_1,
            "endereco": endereco_1,
            "telefone": telefone_1,
            "conteudo": conteudo_1
        }
        st.success("Dados do Ofício 1 salvos!")

with tab2:
    st.subheader("Dados do Ofício 2")
    with st.form("dados_oficio_2"):
        nome_2 = st.text_input("Nome do Destinatário 2")
        endereco_2 = st.text_input("Endereço 2")
        telefone_2 = st.text_input("Telefone 2")
        conteudo_2 = st.text_area("Conteúdo do Ofício 2", 
                                  "Informamos que o processo em questão requer documentação adicional. Favor providenciar os documentos solicitados em anexo no prazo de 10 dias úteis.")
        submit_2 = st.form_submit_button("Salvar Dados do Ofício 2")
    
    if submit_2:
        dados_oficios["oficio_2"] = {
            "nome": nome_2,
            "endereco": endereco_2,
            "telefone": telefone_2,
            "conteudo": conteudo_2
        }
        st.success("Dados do Ofício 2 salvos!")

with tab3:
    st.subheader("Dados do Ofício 3")
    with st.form("dados_oficio_3"):
        nome_3 = st.text_input("Nome do Destinatário 3")
        endereco_3 = st.text_input("Endereço 3")
        telefone_3 = st.text_input("Telefone 3")
        conteudo_3 = st.text_area("Conteúdo do Ofício 3", 
                                  "Notificamos que o prazo para manifestação no processo referido está se esgotando. Solicitamos sua atenção para o cumprimento dos prazos legais.")
        submit_3 = st.form_submit_button("Salvar Dados do Ofício 3")
    
    if submit_3:
        dados_oficios["oficio_3"] = {
            "nome": nome_3,
            "endereco": endereco_3,
            "telefone": telefone_3,
            "conteudo": conteudo_3
        }
        st.success("Dados do Ofício 3 salvos!")

# Verificar se temos dados para pelo menos um ofício
if st.session_state.get("dados_oficios"):
    dados_oficios = st.session_state.dados_oficios

# Botão para gerar todos os ofícios
if st.button("Gerar Todos os Ofícios"):
    if not processo or not assinatura:
        st.error("Preencha o número do processo e a assinatura antes de gerar os ofícios!")
    elif len(dados_oficios) < 3:
        st.warning("Preencha e salve os dados de todos os três ofícios antes de gerar!")
    else:
        st.success("Gerando ofícios...")
        
        # Criar os três documentos
        arquivos = {
            "Ofício 1": criar_oficio_word(
                dados_oficios["oficio_1"]["nome"],
                dados_oficios["oficio_1"]["endereco"],
                dados_oficios["oficio_1"]["telefone"],
                processo,
                assinatura,
                "001",
                dados_oficios["oficio_1"]["conteudo"]
            ),
            "Ofício 2": criar_oficio_word(
                dados_oficios["oficio_2"]["nome"],
                dados_oficios["oficio_2"]["endereco"],
                dados_oficios["oficio_2"]["telefone"],
                processo,
                assinatura,
                "002",
                dados_oficios["oficio_2"]["conteudo"]
            ),
            "Ofício 3": criar_oficio_word(
                dados_oficios["oficio_3"]["nome"],
                dados_oficios["oficio_3"]["endereco"],
                dados_oficios["oficio_3"]["telefone"],
                processo,
                assinatura,
                "003",
                dados_oficios["oficio_3"]["conteudo"]
            )
        }
        
        # Criar um ZIP com todos os arquivos
        zip_buffer = criar_zip_oficios(arquivos)
        
        # Botão para baixar o ZIP
        st.download_button(
            label="Baixar Todos os Ofícios (ZIP)",
            data=zip_buffer,
            file_name="Oficios.zip",
            mime="application/zip",
            key="download-zip"
        )
        
        # Opção para baixar individualmente também
        st.subheader("Ou baixe individualmente:")
        download_col1, download_col2, download_col3 = st.columns(3)
        
        with download_col1:
            with open(arquivos["Ofício 1"], "rb") as file:
                st.download_button(
                    label=f"Baixar Ofício 1 ({dados_oficios['oficio_1']['nome']})",
                    data=file,
                    file_name="Ofício_1.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-1"
                )
        
        with download_col2:
            with open(arquivos["Ofício 2"], "rb") as file:
                st.download_button(
                    label=f"Baixar Ofício 2 ({dados_oficios['oficio_2']['nome']})",
                    data=file,
                    file_name="Ofício_2.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-2"
                )
        
        with download_col3:
            with open(arquivos["Ofício 3"], "rb") as file:
                st.download_button(
                    label=f"Baixar Ofício 3 ({dados_oficios['oficio_3']['nome']})",
                    data=file,
                    file_name="Ofício_3.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-3"
                )
        
        # Remover arquivos temporários
        for caminho in arquivos.values():
            os.remove(caminho)

# Inicializar o estado da sessão para armazenar os dados dos ofícios
if "dados_oficios" not in st.session_state:
    st.session_state.dados_oficios = {}

# Atualizar o estado da sessão com os dados dos ofícios
if dados_oficios:
    st.session_state.dados_oficios.update(dados_oficios)
