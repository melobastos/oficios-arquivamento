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

# Função para criar ofício modelo de comunicação de arquivamento
def criar_oficio_arquivamento(numero_oficio, data, numero_idea):
    doc = Document()
    doc = formatar_documento(doc)
    
    # Número do ofício
    ref_paragraph = doc.add_paragraph()
    ref_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    ref_run = ref_paragraph.add_run(f"OFÍCIO Nº {numero_oficio}/2024/SP-FSA/25ªPJ")
    ref_run.bold = True
    
    # Referência IDEA
    idea_paragraph = doc.add_paragraph()
    idea_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    idea_run = idea_paragraph.add_run(f"(Ref.: IDEA nº {numero_idea}/2024)")
    idea_run.bold = True
    idea_run.italic = True
    
    # Local e data
    data_paragraph = doc.add_paragraph()
    data_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    data_paragraph.add_run(f"Feira de Santana, {data}")
    
    # Destinatário
    doc.add_paragraph("A Sua Excelência a Senhora")
    
    nome_paragraph = doc.add_paragraph()
    nome_run = nome_paragraph.add_run("MARIA CLÉCIA VASCONCELOS DE MORAIS FIRMINO COSTA")
    nome_run.bold = True
    
    doc.add_paragraph("Delegacia Especializada de Atendimento à Mulher de Feira de Santana --")
    doc.add_paragraph("DEAM")
    doc.add_paragraph("Avenida Maria Quitéria nº 1870, Centro")
    doc.add_paragraph("Feira de Santana -- Bahia, CEP: 44001-344")
    doc.add_paragraph("E-mail: deam.feiradesantana@pcivil.ba.gov.br")
    
    # Vocativo
    doc.add_paragraph("Excelentíssima Senhora,")
    
    # Conteúdo
    conteudo = (
        "Com os nossos cordiais cumprimentos, DE ORDEM DE DRA. NAYARA VALTÉRCIA "
        "GONÇALVES BARRETO, Promotora de Justiça titular da 25ª Promotoria de "
        "Justiça, sirvo-me do presente para, atendendo ao quanto disposto no art. "
        "28 do Código de Processo Penal, comunicar a Vossa Excelência o "
        f"ARQUIVAMENTO do Inquérito Policial IDEA nº {numero_idea}/2024, consoante "
        "Promoção anexa."
    )
    doc.add_paragraph(conteudo)
    
    # Despedida
    doc.add_paragraph("Cordialmente,")
    doc.add_paragraph("\n")
    doc.add_paragraph("(assinado eletronicamente)")
    doc.add_paragraph("\n")
    
    assinatura_paragraph = doc.add_paragraph()
    assinatura_run = assinatura_paragraph.add_run("Larissa Brandão de Carvalho e Carvalho")
    assinatura_run.bold = True
    
    cargo_paragraph = doc.add_paragraph()
    cargo_paragraph.add_run("Secretaria Processual")
    
    # Criar arquivo temporário
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    
    return temp_file.name

# Função para criar um documento Word (Ofício) para os modelos 2 e 3
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

# Inicializar o estado da sessão para armazenar os dados dos ofícios
if "dados_oficios" not in st.session_state:
    st.session_state.dados_oficios = {}

# Interface do Streamlit
st.set_page_config(page_title="Gerador de Ofícios", layout="wide")

st.title("Gerador de Ofícios Automáticos")
st.write("Preencha os dados abaixo para gerar três ofícios.")

# Número do processo comum aos ofícios 2 e 3
processo = st.text_input("Número do Processo (para ofícios 2 e 3)")
assinatura = st.text_area("Assinatura (Nome e cargo do responsável para ofícios 2 e 3)")

# Criar abas para cada ofício
tab1, tab2, tab3 = st.tabs(["Ofício 1 - Arquivamento", "Ofício 2", "Ofício 3"])

with tab1:
    st.subheader("Dados do Ofício 1 - Comunicação de Arquivamento")
    st.info("Este ofício segue o modelo padrão de comunicação de arquivamento. Apenas alguns campos podem ser alterados.")
    
    with st.form("dados_oficio_1"):
        numero_oficio = st.text_input("Número do Ofício", "4886")
        data_oficio = st.date_input("Data do Ofício").strftime("%d de %B de %Y")
        idea_numero = st.text_input("Número IDEA (número do processo)", "596.9.489799")
        
        submit_1 = st.form_submit_button("Salvar Dados do Ofício 1")
    
    if submit_1:
        st.session_state.dados_oficios["oficio_1"] = {
            "numero_oficio": numero_oficio,
            "data_oficio": data_oficio,
            "idea_numero": idea_numero
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
        st.session_state.dados_oficios["oficio_2"] = {
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
        st.session_state.dados_oficios["oficio_3"] = {
            "nome": nome_3,
            "endereco": endereco_3,
            "telefone": telefone_3,
            "conteudo": conteudo_3
        }
        st.success("Dados do Ofício 3 salvos!")

# Mostrar o status atual dos dados salvos
status_col1, status_col2, status_col3 = st.columns(3)
with status_col1:
    if "oficio_1" in st.session_state.dados_oficios:
        # Corrigido para verificar se a chave existe antes de tentar acessá-la
        st.info(f"Ofício 1: Dados do Ofício nº {st.session_state.dados_oficios['oficio_1'].get('numero_oficio', '')} salvos ✅")
    else:
        st.warning("Ofício 1: Dados não salvos ❌")
        
with status_col2:
    if "oficio_2" in st.session_state.dados_oficios:
        # Corrigido para verificar se a chave existe antes de tentar acessá-la
        st.info(f"Ofício 2: Dados de {st.session_state.dados_oficios['oficio_2'].get('nome', '')} salvos ✅")
    else:
        st.warning("Ofício 2: Dados não salvos ❌")
        
with status_col3:
    if "oficio_3" in st.session_state.dados_oficios:
        # Corrigido para verificar se a chave existe antes de tentar acessá-la
        st.info(f"Ofício 3: Dados de {st.session_state.dados_oficios['oficio_3'].get('nome', '')} salvos ✅")
    else:
        st.warning("Ofício 3: Dados não salvos ❌")

# Botão para gerar todos os ofícios
if st.button("Gerar Todos os Ofícios"):
    if not all(f"oficio_{i}" in st.session_state.dados_oficios for i in range(1, 4)):
        st.warning("Preencha e salve os dados de todos os três ofícios antes de gerar!")
        # Mostrar quais ofícios estão faltando
        for i in range(1, 4):
            if f"oficio_{i}" not in st.session_state.dados_oficios:
                st.warning(f"Ofício {i} não tem dados salvos!")
    elif not processo or not assinatura:
        # Verifique o processo e assinatura apenas para os ofícios 2 e 3
        st.error("Preencha o número do processo e a assinatura antes de gerar os ofícios 2 e 3!")
    else:
        st.success("Gerando ofícios...")
        
        # Criar os três documentos
        arquivos = {
            "Ofício 1": criar_oficio_arquivamento(
                st.session_state.dados_oficios["oficio_1"]["numero_oficio"],
                st.session_state.dados_oficios["oficio_1"]["data_oficio"],
                st.session_state.dados_oficios["oficio_1"]["idea_numero"]
            ),
            "Ofício 2": criar_oficio_word(
                st.session_state.dados_oficios["oficio_2"]["nome"],
                st.session_state.dados_oficios["oficio_2"]["endereco"],
                st.session_state.dados_oficios["oficio_2"]["telefone"],
                processo,
                assinatura,
                "002",
                st.session_state.dados_oficios["oficio_2"]["conteudo"]
            ),
            "Ofício 3": criar_oficio_word(
                st.session_state.dados_oficios["oficio_3"]["nome"],
                st.session_state.dados_oficios["oficio_3"]["endereco"],
                st.session_state.dados_oficios["oficio_3"]["telefone"],
                processo,
                assinatura,
                "003",
                st.session_state.dados_oficios["oficio_3"]["conteudo"]
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
                    label=f"Baixar Ofício 1 (Nº {st.session_state.dados_oficios['oficio_1']['numero_oficio']})",
                    data=file,
                    file_name="Ofício_1_Arquivamento.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-1"
                )
        
        with download_col2:
            with open(arquivos["Ofício 2"], "rb") as file:
                st.download_button(
                    label=f"Baixar Ofício 2 ({st.session_state.dados_oficios['oficio_2']['nome']})",
                    data=file,
                    file_name="Ofício_2.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-2"
                )
        
        with download_col3:
            with open(arquivos["Ofício 3"], "rb") as file:
                st.download_button(
                    label=f"Baixar Ofício 3 ({st.session_state.dados_oficios['oficio_3']['nome']})",
                    data=file,
                    file_name="Ofício_3.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download-3"
                )
        
        # Remover arquivos temporários após um tempo razoável para download
        for caminho in arquivos.values():
            try:
                os.remove(caminho)
            except:
                pass

# Adicionar um botão para limpar todos os dados
if st.button("Limpar Todos os Dados"):
    st.session_state.dados_oficios = {}
    st.success("Todos os dados foram limpos!")
    st.experimental_rerun()
