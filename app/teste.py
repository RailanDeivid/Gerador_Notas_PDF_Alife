import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog
from fpdf import FPDF
import pandas as pd
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

def gerar_zip_com_pdfs(df):
    zip_name = tempfile.mktemp(suffix=".zip")
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        os.makedirs('temp_pdfs', exist_ok=True)

        for index, row in df.iterrows():
            pdf_path = gerar_pdf(row, f"NOTA_DÉBITO_{row['LOJA']}")
            zipf.write(pdf_path, f"NOTA DÉBITO - {row['LOJA']}.pdf")
            os.remove(pdf_path)

        shutil.rmtree('temp_pdfs')

    return zip_name


    
    
# Função para carregar o arquivo Excel
def carregar_arquivo():
    global df_global
    arquivo = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if arquivo:
        df_global = pd.read_excel(arquivo)

        if df_global.empty:
            messagebox.showerror("Erro", "O arquivo Excel está vazio. Verifique os dados.")
        else:
            colunas_necessarias = ["LOJA", "RAZAO SOCIAL", "CNPJ", "ENDEREÇO", "CEP", "BAIRRO", "EMAIL", "VALOR", "DATA DE EMISSÃO", "DATA DE PAGAMENTO"]
            colunas_faltando = [col for col in colunas_necessarias if col not in df_global.columns]
            if colunas_faltando:
                messagebox.showerror("Erro", f"Faltam as seguintes colunas: {', '.join(colunas_faltando)}.")
            else:
                messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")

# Função para gerar PDFs
def gerar_pdfs():
    global df_global
    if df_global is not None:
        zip_path = gerar_zip_com_pdfs(df_global)
        with open(zip_path, "rb") as f:
            zip_file_name = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP Files", "*.zip")])
            if zip_file_name:
                with open(zip_file_name, "wb") as output_file:
                    output_file.write(f.read())
                messagebox.showinfo("Sucesso", "PDFs gerados com sucesso!")
            os.remove(zip_path)
    else:
        messagebox.showerror("Erro", "Por favor, carregue um arquivo Excel primeiro.")

# Função para pedir o número da nota ao usuário
def pedir_numero_nota():
    global numero_nota
    novo_numero = simpledialog.askinteger("Número da Nota", "Digite o número da nota :")
    if novo_numero:
        numero_nota = novo_numero

# Função para selecionar PDF específico ou todos
def selecionar_tipo_geracao():
    global df_global
    if df_global is None:
        messagebox.showerror("Erro", "Por favor, carregue um arquivo Excel primeiro.")
        return

    opcao_gerar_todos = messagebox.askyesno("Gerar Todos os PDFs", "Deseja gerar todos os PDFs de uma vez?")
    
    if opcao_gerar_todos:  # Gerar todos os PDFs em um ZIP
        gerar_pdfs()
    else:  # Gerar PDF de uma loja específica
        lojas_disponiveis = df_global['LOJA'].unique()

        def loja_selecionada_callback():
            loja_selecionada = loja_var.get()
            if loja_selecionada:
                loja_df = df_global[df_global['LOJA'] == loja_selecionada]
                if loja_df.empty:
                    messagebox.showerror("Erro", "Loja não encontrada.")
                else:
                    # Perguntar se deseja alterar o número da nota
                    alterar_nota = messagebox.askyesno("Alterar Nota Fiscal", "Deseja definir um novo número de nota fiscal?")
                    
                    if alterar_nota:
                        pedir_numero_nota()
                    
                    # Gerar o PDF individualmente
                    pdf_path = gerar_pdf(loja_df.iloc[0], f"NOTA_DÉBITO_{loja_selecionada}")

                    # Solicitar onde salvar o PDF
                    pdf_file_name = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                  filetypes=[("PDF Files", "*.pdf")],
                                                                  initialfile=f"NOTA DÉBITO - {loja_selecionada}.pdf")
                    if pdf_file_name:
                        shutil.move(pdf_path, pdf_file_name)
                        messagebox.showinfo("Sucesso", "PDF gerado com sucesso!")
                    
                    # Fechar a janela de seleção da loja
                    loja_window.destroy()

        # Criar a janela da lista suspensa
        loja_window = tk.Toplevel(tk_root)
        loja_window.title("Escolher Loja")
        loja_window.geometry("350x200")
        loja_window.configure(bg="#f0f0f0")
        
        # Criar um título
        titulo_label = tk.Label(loja_window, text="Selecione a Loja", font=("Arial", 14, "bold"), bg="#f0f0f0")
        titulo_label.pack(pady=10)
        
        loja_var = tk.StringVar(loja_window)
        loja_var.set(lojas_disponiveis[0])  # Define o valor inicial como a primeira loja

        # Criar o OptionMenu com as opções de loja
        loja_menu = tk.OptionMenu(loja_window, loja_var, *lojas_disponiveis)
        loja_menu.config(font=("Arial", 12))
        loja_menu.pack(pady=10)

        # Botão para confirmar a escolha
        confirmar_button = tk.Button(loja_window, text="Confirmar", command=loja_selecionada_callback, 
                                     font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", padx=10, pady=5)
        confirmar_button.pack(pady=10)
        
        # Centralizar a janela na tela
        loja_window.update_idletasks()
        largura_janela = loja_window.winfo_width()
        altura_janela = loja_window.winfo_height()
        largura_tela = loja_window.winfo_screenwidth()
        altura_tela = loja_window.winfo_screenheight()
        x = (largura_tela - largura_janela) // 2
        y = (altura_tela - altura_janela) // 2
        loja_window.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")


# Configuração da interface gráfica principal
tk_root = tk.Tk()
tk_root.title("Gerador de Notas de Débito em PDF")
tk_root.geometry("600x400")

# Centralizar a janela principal
tk_root.update_idletasks()
largura_janela = tk_root.winfo_width()
altura_janela = tk_root.winfo_height()
largura_tela = tk_root.winfo_screenwidth()
altura_tela = tk_root.winfo_screenheight()
x = (largura_tela - largura_janela) // 2
y = (altura_tela - altura_janela) // 2
tk_root.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")


# Carregar logo
logo = tk.PhotoImage(file="logo Alife.png")  
logo_label = tk.Label(tk_root, image=logo)
logo_label.pack(pady=10)

# Configurar a janela
tk_root.configure(bg="#E8EBEF")  
frame = tk.Frame(tk_root, bg="#E8EBEF")
frame.pack(padx=10, pady=10)

# Título
titulo_label = tk.Label(frame, text="Gerador de Notas de Débito", font=("Arial", 16, "bold"), bg="#E8EBEF")
titulo_label.pack(pady=10)

# Botões da interface com estilo
botao_carregar = tk.Button(frame, text="Carregar Arquivo Excel", command=carregar_arquivo, font=("Arial", 12), bg="#4CAF50", fg="white", relief="raised", bd=2, padx=10, pady=5)
botao_carregar.pack(pady=10)

botao_gerar_pdfs = tk.Button(frame, text="Gerar PDFs", command=selecionar_tipo_geracao, font=("Arial", 12), bg="#008CBA", fg="white", relief="raised", bd=2, padx=10, pady=5)
botao_gerar_pdfs.pack(pady=10)

# Executar a interface gráfica
tk_root.mainloop()
