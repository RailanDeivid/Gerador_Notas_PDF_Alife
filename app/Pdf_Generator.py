import streamlit as st
import pandas as pd
from fpdf import FPDF
import tempfile
import zipfile
import os
import shutil  

# Função para gerar o PDF
def gerar_pdf(dados, nome_arquivo):
    # Garantir que o diretório temp_pdfs exista
    os.makedirs('temp_pdfs', exist_ok=True)
    
    # Configuração da página
    pdf = FPDF(orientation='L', unit='mm', format='A4')  
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Definindo a fonte do título
    pdf.set_font("Arial", "B", 18)
    # Título
    pdf.set_text_color(0, 0, 0)  
    pdf.set_fill_color(132, 183, 83)  
    pdf.cell(280, 10, "NOTA DE DÉBITO", ln=True, align='C', border=1, fill=True)
    pdf.ln(20)
    
    # Fonte do texto
    pdf.set_font("Arial", "", 10)
    
    # Dados do Destinatário - Alinhados à esquerda
    pdf.set_xy(10, 30) 
    pdf.cell(100, 5, f"{dados['LOJA']}", ln=True)
    pdf.ln(2)
    pdf.cell(100, 5, f"ENDEREÇO: {dados['ENDEREÇO']}", ln=True)
    pdf.cell(100, 5, f"CEP: {dados['CEP']}", ln=True)
    pdf.cell(100, 5, f"BAIRRO: {dados['BAIRRO']}", ln=True)
    pdf.cell(100, 5, f"CNPJ: {dados['CNPJ']}", ln=True)
    pdf.cell(100, 5, f"E-mail: {dados['EMAIL']}", ln=True)
    
    # Ajustando a posição para o lado direito
    pdf.set_xy(120, 45) 
    pdf.cell(200, 10, f"CNPJ: 11513881000160", ln=True)
    pdf.set_xy(120, 50)
    pdf.cell(200, 10, f"Rua AUGUSTA, 3000", ln=True)
    pdf.set_xy(120, 55)
    pdf.cell(200, 10, f"CERQUEIRA CESAR - São Paulo / SP", ln=True)
    pdf.set_xy(120, 60)
    pdf.cell(200, 10, f"CEP: 01.412-100", ln=True)
    pdf.set_xy(120, 65)
    pdf.cell(200, 10, f"E-mail: laiane.costa@alifegroup.com.br", ln=True)
    
    # Cabeçalho acima da tabela
    pdf.set_xy(230, 25)   
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(36, 38, 68)
    pdf.cell(60, 5, 'Nº', ln=True, align='C', border=1, fill=True)
    
    # Cabeçalho da tabela 
    pdf.set_xy(190, 30) 
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(40, 10, 'Nota de Débito', border=1, align='C')
    pdf.set_xy(190, 40) 
    pdf.cell(40, 10, 'Emissão', border=1, align='C')
    
    # Valores da tabela
    pdf.set_font("Arial", "", 11)
    pdf.set_xy(230, 30) 
    pdf.cell(60, 10, "0001", border=1, ln=True, align='C')
    pdf.set_xy(230, 40) 
    pdf.cell(60, 10, "23/03/2025", border=1, ln=True, align='C')
    
    # Salvar o PDF temporariamente
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir="temp_pdfs")
    pdf.output(temp_pdf.name)
    return temp_pdf.name

# ------------------------- Função para gerar um arquivo ZIP contendo todos os PDFs
def gerar_zip_com_pdfs(df):
    zip_name = tempfile.mktemp(suffix=".zip")
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        # Criar uma pasta temporária para armazenar os PDFs
        os.makedirs('temp_pdfs', exist_ok=True)
        
        for index, row in df.iterrows():
            pdf_path = gerar_pdf(row, f"NOTA_DÉBITO_{row['LOJA']}")
            # Adicionar cada PDF ao arquivo zip com o nome desejado
            zipf.write(pdf_path, f"NOTA DÉBITO - {row['LOJA']}.pdf")
            os.remove(pdf_path)  # Remover o arquivo PDF após adicionar ao ZIP
        
        # Após criar o zip, exclui a pasta temporária e seus arquivos
        shutil.rmtree('temp_pdfs')
    
    return zip_name

# -------------------------------------------------------------------- Interface Streamlit -------------------------------------------------------------------- #
# ----------------------- Configuração da página
st.set_page_config(page_title="Gerador de Notas de Débito em PDF",
                   page_icon="📄", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
    <style>
        .text {
            text-align: center; /* Centraliza o texto */
            font-size: 10px;
            color: hsl(0, 0%, 35%);
            
        }
        .linkedin-icon {
            width: 15px;
            vertical-align: center;
            margin-left: 5px;
        }
    </style>
    """, unsafe_allow_html=True)
st.markdown("<p class='text'>💡 Desenvolvido por Railan Deivid<br>", unsafe_allow_html=True)
st.title("Gerador de Notas de Débito em PDF")
st.markdown("<br>", unsafe_allow_html=True)


# ------------------- Upload do arquivo Excel
arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

# ------------------- Verifica se o arquivo foi enviado
if arquivo:
    df = pd.read_excel(arquivo)
    st.write("Dados carregados:")
    st.dataframe(df)
    
    # Escolher se quer gerar todos ou selecionar manualmente
    opcao = st.radio("Escolha uma opção:", ("Gerar todos os PDFs", "Escolher quais gerar"),horizontal=True)
    
    if opcao == "Escolher quais gerar":
        # Checkbox para selecionar os PDFs individuais
        cols = st.columns(1)
        with cols[0]:
            selecao = st.multiselect("Selecione as lojas para gerar PDF", df['LOJA'].tolist())
    
    if st.button("Gerar PDF"):
        if not df.empty:
            if opcao == "Gerar todos os PDFs":
                zip_path = gerar_zip_com_pdfs(df)
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="Baixar Todos os PDFs",
                        data=f,
                        file_name="notas_de_debito.zip",
                        mime="application/zip"
                    )
            elif opcao == "Escolher quais gerar":
                for loja in selecao:
                    row = df[df['LOJA'] == loja].iloc[0]
                    pdf_path = gerar_pdf(row, f"NOTA_DÉBITO_{row['LOJA']}")
                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            label=f"Baixar Nota de Débito - {loja}",
                            data=f,
                            file_name=f"NOTA DÉBITO - {loja}.pdf",  
                            mime="application/pdf"
                        )
        else:
            st.error("O arquivo Excel está vazio. Verifique os dados e tente novamente.")
