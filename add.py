import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import tempfile
import os
import base64
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

# Função para criar um link de download
def get_download_link(file_path, label):
    with open(file_path, "rb") as file:
        bytes_data = file.read()
    b64 = base64.b64encode(bytes_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{label}.docx">{label}</a>'
    return href

# Interface do Streamlit
st.set_page_config(page_title="Gerador de Ofícios", layout="wide")

st.title("Gerador de Ofícios Automáticos")
st.write("Preencha os dados abaixo para gerar três ofícios com conteúdos diferentes.")

# Criando colunas para layout
col1, col2 = st.columns([2, 1])

with col1:
    # Formulário para entrada de dados
    with st.form("dados_oficio"):
        nome = st.text_input("Nome do Destinatário")
        endereco = st.text_input("Endereço")
        telefone = st.text_input("Telefone")
        processo = st.text_input("Número do Processo")
        
        # Conteúdo personalizado para cada ofício
        st.subheader("Conteúdo Personalizado dos Ofícios")
        conteudo_1 = st.text_area("Conteúdo do Ofício 1", 
                                  "Venho por meio deste comunicar sobre o andamento do processo mencionado. Solicito seu comparecimento em nosso escritório para tratar de assuntos relacionados ao mesmo.")
        conteudo_2 = st.text_area("Conteúdo do Ofício 2", 
                                  "Informamos que o processo em questão requer documentação adicional. Favor providenciar os documentos solicitados em anexo no prazo de 10 dias úteis.")
        conteudo_3 = st.text_area("Conteúdo do Ofício 3", 
                                  "Notificamos que o prazo para manifestação no processo referido está se esgotando. Solicitamos sua atenção para o cumprimento dos prazos legais.")
        
        assinatura = st.text_area("Assinatura (Nome e cargo do responsável)")
        submit = st.form_submit_button("Gerar Ofícios")

with col2:
    st.subheader("Instruções")
    st.info("""
    1. Preencha todos os campos obrigatórios
    2. Personalize o conteúdo de cada ofício
    3. Clique em "Gerar Ofícios"
    4. Faça o download dos documentos gerados
    """)
    
    # Exibir preview ou exemplo
    if st.checkbox("Mostrar exemplo de um ofício"):
        st.image("https://via.placeholder.com/400x500?text=Exemplo+de+Ofício", use_column_width=True)

# Se o usuário enviar os dados
if submit:
    if nome and endereco and telefone and processo and assinatura:
        st.success("Ofícios gerados com sucesso!")
        
        # Criar os três documentos
        arquivos = {
            "Ofício 1": criar_oficio_word(nome, endereco, telefone, processo, assinatura, "001", conteudo_1),
            "Ofício 2": criar_oficio_word(nome, endereco, telefone, processo, assinatura, "002", conteudo_2),
            "Ofício 3": criar_oficio_word(nome, endereco, telefone, processo, assinatura, "003", conteudo_3),
        }
        
        # Criar área para download
        st.subheader("Download dos Ofícios")
        download_col1, download_col2, download_col3 = st.columns(3)
        
        with download_col1:
            with open(arquivos["Ofício 1"], "rb") as file:
                st.download_button(
                    label="Baixar Ofício 1",
                    data=file,
                    file_name="Ofício_1.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        
        with download_col2:
            with open(arquivos["Ofício 2"], "rb") as file:
                st.download_button(
                    label="Baixar Ofício 2",
                    data=file,
                    file_name="Ofício_2.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        
        with download_col3:
            with open(arquivos["Ofício 3"], "rb") as file:
                st.download_button(
                    label="Baixar Ofício 3",
                    data=file,
                    file_name="Ofício_3.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        
        # Botão para baixar todos os ofícios
        st.markdown("#### Ou baixe todos os ofícios de uma vez:")
        
        # Criar um ZIP com todos os arquivos (simulado)
        st.info("Funcionalidade para download em ZIP será implementada em breve.")
        
        # Remover arquivos temporários
        for caminho in arquivos.values():
            os.remove(caminho)
    else:
        st.error("Preencha todos os campos antes de gerar os ofícios!")
