import streamlit as st
from fpdf import FPDF
import tempfile
import os

# Função para criar o PDF do ofício
def criar_oficio(nome, endereco, telefone, processo, assinatura, tipo):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Título do Ofício
    pdf.set_font("Arial", style="B", size=14)
    pdf.cell(200, 10, f"Ofício {tipo}", ln=True, align="C")
    pdf.ln(10)

    # Corpo do Ofício
    pdf.set_font("Arial", size=12)
    texto = f"""
    Prezado(a) {nome},

    Vimos por meio deste ofício informar sobre o processo de número {processo}.
    Seguem os detalhes:

    Nome: {nome}
    Endereço: {endereco}
    Telefone: {telefone}

    Atenciosamente,

    {assinatura}
    """
    pdf.multi_cell(0, 10, texto)

    # Salvar arquivo temporário
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_file.name)
    return temp_file.name

# Interface Streamlit
st.title("Gerador de Ofícios Automáticos")

# Formulário
with st.form("dados_oficio"):
    nome = st.text_input("Nome")
    endereco = st.text_input("Endereço")
    telefone = st.text_input("Telefone")
    processo = st.text_input("Número do Processo")
    assinatura = st.text_area("Assinatura (Nome do responsável)")

    submit = st.form_submit_button("Gerar Ofícios")

# Se o usuário enviar o formulário
if submit:
    if nome and endereco and telefone and processo and assinatura:
        # Criar os três ofícios
        arquivos = {
            "Ofício 1": criar_oficio(nome, endereco, telefone, processo, assinatura, "1"),
            "Ofício 2": criar_oficio(nome, endereco, telefone, processo, assinatura, "2"),
            "Ofício 3": criar_oficio(nome, endereco, telefone, processo, assinatura, "3"),
        }

        # Exibir os links para download
        for titulo, caminho in arquivos.items():
            with open(caminho, "rb") as file:
                st.download_button(
                    label=f"Baixar {titulo}",
                    data=file,
                    file_name=f"{titulo}.pdf",
                    mime="application/pdf",
                )

        # Remover arquivos temporários após o download
        for caminho in arquivos.values():
            os.remove(caminho)
    else:
        st.error("Preencha todos os campos antes de gerar os ofícios!")
