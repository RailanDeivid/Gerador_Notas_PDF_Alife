import streamlit as st
import pandas as pd
from fpdf import FPDF
import tempfile
import zipfile
import os
import shutil

# Variável global para o número da nota
numero_nota = 1

# Função para gerar o PDF
def gerar_pdf(dados):
    global numero_nota  # Usa a variável global
    
    # Criar um diretório temporário
    temp_dir = tempfile.mkdtemp()
    temp_pdf_path = os.path.join(temp_dir, f"NOTA_DÉBITO_{dados['LOJA']}.pdf")
    
    # Configuração da página
    pdf = FPDF(orientation='L', unit='mm', format='A4')  
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Título
    pdf.set_font("Arial", "B", 18)
    pdf.set_text_color(0, 0, 0)  
    pdf.set_fill_color(132, 183, 83)  
    pdf.cell(280, 10, "NOTA DE DÉBITO", ln=True, align='C', border=1, fill=True)
    pdf.ln(20)
    
    # Incrementar o número da nota
    numero_nota += 1
    
    # Salvar o PDF
    pdf.output(temp_pdf_path)
    return temp_pdf_path, temp_dir

# Função para gerar um arquivo ZIP contendo todos os PDFs
def gerar_zip_com_pdfs(df):
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, "notas_de_debito.zip")
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for index, row in df.iterrows():
            pdf_path, pdf_temp_dir = gerar_pdf(row)
            zipf.write(pdf_path, os.path.basename(pdf_path))
            shutil.rmtree(pdf_temp_dir)  # Remover pasta temporária do PDF
    
    return zip_path, temp_dir

# Interface Streamlit
st.set_page_config(page_title="Gerador de Notas de Débito em PDF",
                   page_icon="📄", layout="wide", initial_sidebar_state="expanded")
st.title("Gerador de Notas de Débito em PDF")

# Upload do arquivo Excel
arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

if arquivo:
    df = pd.read_excel(arquivo)
    if df.empty:
        st.error("O arquivo Excel está vazio. Verifique os dados.")
    else:
        st.write("Dados carregados:")
        st.dataframe(df)
        
        opcao = st.radio("Escolha uma opção:", ("Gerar todos os PDFs", "Escolher quais gerar"), horizontal=True)
        
        if opcao == "Escolher quais gerar":
            loja_selecionada = st.selectbox("Selecione a loja para gerar o PDF", df['LOJA'].tolist())
            numero_manual = st.radio("Deseja inserir manualmente o número da nota?", ("Não", "Sim"))
            if numero_manual == "Sim":
                numero_nota_input = st.text_input("Insira o número da nota (0000):", value=str(numero_nota), max_chars=4)
                if numero_nota_input:
                    numero_nota = int(numero_nota_input)

        if st.button("Gerar PDF"):
            if not df.empty:
                if opcao == "Gerar todos os PDFs":
                    zip_path, temp_zip_dir = gerar_zip_com_pdfs(df)
                    with open(zip_path, "rb") as f:
                        st.download_button("Baixar Todos os PDFs", data=f, file_name="notas_de_debito.zip", mime="application/zip")
                    shutil.rmtree(temp_zip_dir)  # Remover pasta temporária do ZIP
                elif opcao == "Escolher quais gerar":
                    row = df[df['LOJA'] == loja_selecionada].iloc[0]
                    pdf_path, temp_pdf_dir = gerar_pdf(row)
                    with open(pdf_path, "rb") as f:
                        st.download_button(f"Baixar PDF da {loja_selecionada}", data=f, file_name=f"NOTA_DEBITO_{loja_selecionada}.pdf", mime="application/pdf")
                    shutil.rmtree(temp_pdf_dir)  # Remover pasta temporária do PDF
            else:
                st.error("O arquivo Excel está vazio. Verifique os dados e tente novamente.")
