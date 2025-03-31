import streamlit as st
import pandas as pd
from fpdf import FPDF
from streamlit_option_menu import option_menu
import tempfile
import zipfile
import os
import shutil 


# Variável global para o número da nota
numero_nota = 1

# Função para gerar o PDF
def gerar_pdf(dados, nome_arquivo):
    
    global numero_nota  # Usa a variável global
    
    # Garantir que o diretório temp_pdfs exista
    os.makedirs('temp_pdfs', exist_ok=True)
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
    

    
    # Dados casa pagantet e
    pdf.set_font("Arial", "", 10)
    pdf.set_xy(10, 30) 
    pdf.cell(100, 5, f"{dados['LOJA']}", ln=True)
    pdf.ln(2)
    pdf.cell(100, 5, f"ENDEREÇO: {dados['ENDEREÇO']}", ln=True)
    pdf.cell(100, 5, f"CEP: {dados['CEP']}", ln=True)
    pdf.cell(100, 5, f"BAIRRO: {dados['BAIRRO']}", ln=True)
    pdf.cell(100, 5, f"CNPJ: {dados['CNPJ']}", ln=True)
    pdf.cell(100, 5, f"E-mail: {dados['EMAIL']}", ln=True)
    
    # Dados Destinatário 
    pdf.set_xy(118, 45) 
    pdf.cell(200, 10, f"CNPJ: 11513881000160", ln=True)
    pdf.set_xy(118, 50)
    pdf.cell(200, 10, f"Rua AUGUSTA, 3000", ln=True)
    pdf.set_xy(118, 55)
    pdf.cell(200, 10, f"CERQUEIRA CESAR - São Paulo / SP", ln=True)
    pdf.set_xy(118, 60)
    pdf.cell(200, 10, f"CEP: 01.412-100", ln=True)
    pdf.set_xy(118, 65)
    pdf.cell(200, 10, f"E-mail: laiane.costa@alifegroup.com.br", ln=True)
    
    # Cabeçalho acima da tabela
    pdf.set_xy(230, 25)   
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(36, 38, 68)
    pdf.cell(60, 5, 'Nº', ln=True, align='C', border=1, fill=True)
    
    # Cabeçalho da tabela 
    pdf.set_xy(187, 30) 
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(43, 10, 'Nota de Débito', border=1, align='C')
    pdf.set_xy(187, 40) 
    pdf.cell(43, 10, 'Emissão', border=1, align='C')
    pdf.set_xy(187, 50) 
    pdf.cell(43, 10, 'Vencimento', border=1, align='C')
    pdf.set_xy(187, 60) 
    pdf.cell(43, 10, 'Finalidade', border=1, align='C')
    pdf.set_xy(187, 70) 
    pdf.cell(43, 10, 'Forma de pagamento', border=1, align='C')
    
    # Valores da tabela
    pdf.set_font("Arial", "", 11)
    pdf.set_xy(230, 30) 
    numero_nota_str = str(numero_nota).zfill(4) 
    pdf.cell(60, 10, numero_nota_str, border=1, ln=True, align='C')
    pdf.set_xy(230, 40) 
    pdf.cell(60, 10, f"{dados['DATA DE EMISSÃO'].strftime('%d/%m/%Y')}", border=1, ln=True, align='C')
    pdf.set_xy(230, 50) 
    pdf.cell(60, 10, f"{dados['DATA DE PAGAMENTO'].strftime('%d/%m/%Y')}", border=1, ln=True, align='C')
    pdf.set_xy(230, 60) 
    pdf.cell(60, 10, "Extras", border=1, ln=True, align='C')
    pdf.set_xy(230, 70) 
    pdf.cell(60, 10, "À Vista", border=1, ln=True, align='C')
    pdf.ln(20)
    
    # Incrementar o número da nota após gerar o PDF
    numero_nota += 1
    
    # Titulo parte inferior
    pdf.set_text_color(0, 0, 0)  
    pdf.set_fill_color(132, 183, 83) 
    pdf.set_font("Arial", "B", 14) 
    pdf.cell(280, 10, "DESTINATÁRIO", ln=True, align='L', border=1, fill=True)

    # Dados Destinatário parte inferior 
    pdf.set_font("Arial", "B", 11)
    pdf.cell(200, 10, f"Alife Group", ln=True)
    pdf.set_font("Arial", "", 11)
    pdf.cell(200, 5, f"CNPJ: 11513881000160", ln=True)
    pdf.cell(200, 5, f"Rua AUGUSTA, 3000", ln=True)
    pdf.cell(200, 5, f"CERQUEIRA CESAR - São Paulo / SP", ln=True)
    pdf.cell(200, 5, f"CEP: 01.412-100", ln=True)
    pdf.cell(200, 5, f"E-mail: laiane.costa@alifegroup.com.br", ln=True)
    pdf.ln(10)
    
    
    # Titulo tabela parte inferior
    pdf.set_text_color(0, 0, 0)  
    pdf.set_fill_color(132, 183, 83) 
    pdf.set_font("Arial", "B", 14) 
    pdf.cell(280, 8, "DESCRIÇÃO", ln=True, align='C', border=1, fill=True)
    pdf.set_font("Arial", "B", 12) 
    pdf.set_fill_color(183, 225, 205) 
    pdf.cell(280, 5, "NOTA DE DEBITO", ln=True, align='C', border=1, fill=True)
    pdf.set_font("Arial", "B", 14) 
    pdf.set_fill_color(132, 183, 83)
    pdf.cell(280, 8, "PRESTAÇÃO DE CONTAS", ln=True, align='L', border=1, fill=True)
    pdf.set_font("Arial", "B", 12)
    pdf.set_fill_color(218, 255, 183) 
    
    # Cabeçalhos da tabela Inferior
    
    pdf.cell(30, 5, "Item", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(40, 176) 
    pdf.cell(40, 5, "Descrição", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(80, 176) 
    pdf.cell(30, 5, "Quant.", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(110, 176) 
    pdf.cell(30, 5, "Unit.", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(140, 176) 
    pdf.cell(50, 5, "Prestador de Serviço", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(190, 176) 
    pdf.cell(40, 5, "Valor.", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(230, 176) 
    pdf.cell(60, 5, "Observações", ln=True, align='C', border=1, fill=True)
    
    # Valores tabela Inferior
    pdf.set_font("Arial", "", 11)
    pdf.cell(30, 5, "1", ln=True, align='C', border=1)
    pdf.set_xy(40, 181) 
    pdf.cell(40, 5, "Extras", ln=True, align='C', border=1)
    pdf.set_xy(80, 181) 
    pdf.cell(30, 5, "1", ln=True, align='C', border=1)
    pdf.set_xy(110, 181) 
    pdf.cell(30, 5, "1", ln=True, align='C', border=1)
    pdf.set_xy(140, 181) 
    pdf.cell(50, 5, "Extras", ln=True, align='C', border=1)
    pdf.set_xy(190, 181)
    pdf.cell(40, 5, f"{dados['VALOR']:.2f}".replace('.', ','), ln=True, align='C', border=1)
    pdf.set_xy(230, 181) 
    pdf.cell(60, 5, "", ln=True, align='C', border=1)
    
    
    pdf.set_text_color(0, 0, 0)  
    pdf.set_fill_color(132, 183, 83) 
    pdf.set_font("Arial", "B", 14) 
    pdf.cell(280, 8, "", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(101, 188)
    pdf.cell(20, 5, "TOTAL", ln=True, align='C')
    pdf.set_xy(190, 188)
    pdf.cell(40, 5, f"{dados['VALOR']:.2f}".replace('.', ','), ln=True, align='C')

    
    
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
# Setar título
st.markdown("""
    <style>
        .rounded-title {
            text-align: center; 
            font-size: 40px;
            
        }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 class='rounded-title'>Gerador de Notas de Débito em PDF</h1><br>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)



# ------------------- Upload do arquivo Excel

arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])



# ------------------- Verifica se o arquivo foi enviado
if arquivo:
    
    df = pd.read_excel(arquivo)
    if df.empty:
            
            st.error("O arquivo Excel está vazio. Por favor, verifique os dados.")
    else:
        st.write("Dados carregados:")
        st.dataframe(df)    
        
        # Escolher se quer gerar todos ou selecionar manualmente
        opcao = st.radio("Escolha uma opção:", ("Gerar todos os PDFs", "Escolher quais gerar"),horizontal=True)
        
        if opcao == "Escolher quais gerar":
            # Checkbox para selecionar os PDFs individuais
            cols = st.columns(3)
            with cols[0]:
                loja_selecionada = st.selectbox("Selecione a loja para gerar o PDF", df['LOJA'].tolist())
                # selecao = st.multiselect("Selecione as lojas para gerar PDF", df['LOJA'].tolist())

            # Perguntar se deseja inserir manualmente o número da nota
            numero_manual = st.radio("Deseja inserir manualmente o número da nota?", ("Não", "Sim"))

            if numero_manual == "Sim":
                # Se escolher "Sim", permitir que o usuário insira o número manualmente
                cols = st.columns(3)
                with cols[0]:
                    numero_nota_input = st.text_input("Insira o número da nota (0000):", value=str(numero_nota), max_chars=4)
                if numero_nota_input:
                    numero_nota = int(numero_nota_input)  # Atualiza a variável global com o número manual inserido

    # Agora, exibe o botão "Gerar PDF" após o número ser inserido ou caso não queira inserir
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
            # elif opcao == "Escolher quais gerar":
            #     for loja in selecao:
            #         row = df[df['LOJA'] == loja].iloc[0]
            #         pdf_path = gerar_pdf(row, f"NOTA_DÉBITO_{row['LOJA']}")
            #         with open(pdf_path, "rb") as f:
            #             st.download_button(
            #                 label=f"Baixar Nota de Débito - {loja}",
            #                 data=f,
            #                 file_name=f"NOTA DÉBITO - {loja}.pdf",  
            #                 mime="application/pdf"
            #             )
            elif opcao == "Escolher quais gerar":
                # Selecionar a linha da loja escolhida
                row = df[df['LOJA'] == loja_selecionada].iloc[0]
                pdf_path = gerar_pdf(row, f"NOTA_DÉBITO_{row['LOJA']}")
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        label=f"Baixar PDF da {loja_selecionada}",
                        data=f,
                        file_name=f"NOTA_DEBITO_{loja_selecionada}.pdf",
                        mime="application/pdf"
                    )
        else:
            st.error("O arquivo Excel está vazio. Verifique os dados e tente novamente.")
