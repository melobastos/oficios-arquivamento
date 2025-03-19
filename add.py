import streamlit as st
from docx import Document
import tempfile
import os

# Função para criar um documento Word (Ofício)
def criar_oficio_word(nome, endereco, telefone, processo, assinatura, tipo):
    doc = Document()
    
    # Título
    doc.add_heading(f'Ofício {tipo}', level=1)
    doc.add_paragraph("\n")
    
    # Conteúdo do Ofício
    doc.add_paragraph(f"Prezado(a) {nome},")
    doc.add_paragraph("\n")
    doc.add_paragraph(f"Vimos por meio deste ofício informar sobre o processo de número {processo}.")
    doc.add_paragraph("Segue abaixo os dados:")
    doc.add_paragraph(f"Nome: {nome}")
    doc.add_paragraph(f"Endereço: {endereco}")
    doc.add_paragraph(f"Telefone: {telefone}")
    doc.add_paragraph("\n")
    doc.add_paragraph("Atenciosamente,")
    doc.add_paragraph(assinatura)

    # Criar arquivo temporário
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    
    return temp_file.name

# Interface do Streamlit
st.title("Gerador de Ofícios Automáticos (Word)")

# Formulário para entrada de dados
with st.form("dados_oficio"):
    nome = st.text_input("Nome")
    endereco = st.text_input("Endereço")
    telefone = st.text_input("Telefone")
    processo = st.text_input("Número do Processo")
    assinatura = st.text_area("Assinatura (Nome do responsável)")

    submit = st.form_submit_button("Gerar Ofícios")

# Se o usuário enviar os dados
if submit:
    if nome and endereco and telefone and processo and assinatura:
        # Criar os três documentos
        arquivos = {
            "Ofício 1": criar_oficio_word(nome, endereco, telefone, processo, assinatura, "1"),
            "Ofício 2": criar_oficio_word(nome, endereco, telefone, processo, assinatura, "2"),
            "Ofício 3": criar_oficio_word(nome, endereco, telefone, processo, assinatura, "3"),
        }

        # Oferecer os arquivos para download
        for titulo, caminho in arquivos.items():
            with open(caminho, "rb") as file:
                st.download_button(
                    label=f"Baixar {titulo}",
                    data=file,
                    file_name=f"{titulo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        # Remover arquivos temporários
        for caminho in arquivos.values():
            os.remove(caminho)
    else:
        st.error("Preencha todos os campos antes de gerar os ofícios!")
