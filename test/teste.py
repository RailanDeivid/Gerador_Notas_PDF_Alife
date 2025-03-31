import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog
from fpdf import FPDF
import pandas as pd
import numpy as np
import tempfile
import zipfile
import os
import shutil
import os

# Caminho do arquivo onde o número da nota será armazenado
ARQUIVO_NUMERO_NOTA = "numero_nota.txt"

# Função para ler o número da nota do arquivo (ou iniciar em 1 caso não exista)
def carregar_numero_nota():
    if os.path.exists(ARQUIVO_NUMERO_NOTA):
        with open(ARQUIVO_NUMERO_NOTA, "r") as f:
            return int(f.read().strip())  
    return 1  

# Função para salvar o número da nota no arquivo
def salvar_numero_nota(numero):
    with open(ARQUIVO_NUMERO_NOTA, "w") as f:
        f.write(str(numero))

# Carrega o número da nota ao iniciar
numero_nota = carregar_numero_nota()

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
    

    
    # Dados casa pagante
    pdf.set_font("Arial", "", 10)
    pdf.set_xy(10, 30) 
    pdf.cell(100, 5, f"{dados['LOJA']}", ln=True)
    pdf.ln(2)
    pdf.cell(100, 5, f"ENDEREÇO: {dados['ENDEREÇO']}", ln=True)
    pdf.cell(100, 5, f"CEP: {dados['CEP']}", ln=True)
    pdf.cell(100, 5, f"CNPJ: {dados['CNPJ']}", ln=True)
    if dados.get('EMAIL') is not np.nan:  
        pdf.cell(100, 5, f"E-mail: {dados['EMAIL']}", ln=True)
    
    
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
    salvar_numero_nota(numero_nota) 
    
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
    valor_formatado = f"R$ {dados['VALOR']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    pdf.cell(40, 5, valor_formatado, ln=True, align='C', border=1)
    pdf.set_xy(230, 181) 
    pdf.cell(60, 5, "", ln=True, align='C', border=1)
    
    
    pdf.set_text_color(0, 0, 0)  
    pdf.set_fill_color(132, 183, 83) 
    pdf.set_font("Arial", "B", 14) 
    pdf.cell(280, 8, "", ln=True, align='C', border=1, fill=True)
    pdf.set_xy(101, 188)
    pdf.cell(20, 5, "TOTAL", ln=True, align='C')
    pdf.set_xy(190, 188)
    pdf.cell(40, 5, valor_formatado, ln=True, align='C')

    
    
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
            colunas_necessarias = ["LOJA", "CNPJ", "ENDEREÇO", "CEP", "EMAIL", "VALOR", "DATA DE EMISSÃO", "DATA DE PAGAMENTO"]
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
    
    if opcao_gerar_todos:  
        gerar_pdfs()
    else:  
        lojas_disponiveis = df_global['LOJA'].unique()

        def toggle_selecionar_todas():
            """Marcar ou desmarcar todas as checkboxes."""
            marcar = var_selecionar_todas.get()
            for var in checkboxes.values():
                var.set(marcar)

        def lojas_selecionadas_callback():
            """Gerar PDFs com base nas lojas selecionadas."""
            lojas_selecionadas = [loja for loja, var in checkboxes.items() if var.get()]

            if not lojas_selecionadas:
                messagebox.showerror("Erro", "Nenhuma loja foi selecionada.")
                return

            if len(lojas_selecionadas) == 1:
                loja = lojas_selecionadas[0]
                loja_df = df_global[df_global['LOJA'] == loja]
                
                if loja_df.empty:
                    messagebox.showerror("Erro", f"Dados não encontrados para a loja {loja}.")
                else:
                    pdf_path = gerar_pdf(loja_df.iloc[0], f"NOTA_DÉBITO_{loja}")
                    pdf_file_name = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                  filetypes=[("PDF Files", "*.pdf")],
                                                                  initialfile=f"NOTA DÉBITO - {loja}.pdf")
                    if pdf_file_name:
                        shutil.move(pdf_path, pdf_file_name)
                        messagebox.showinfo("Sucesso", "PDF gerado com sucesso!")

            else:
                pasta_temp = "pdf_temp"
                os.makedirs(pasta_temp, exist_ok=True)
                
                for loja in lojas_selecionadas:
                    loja_df = df_global[df_global['LOJA'] == loja]
                    if loja_df.empty:
                        messagebox.showerror("Erro", f"Dados não encontrados para a loja {loja}.")
                        continue
                    
                    pdf_path = gerar_pdf(loja_df.iloc[0], f"NOTA_DÉBITO_{loja}")
                    shutil.move(pdf_path, os.path.join(pasta_temp, f"NOTA DÉBITO - {loja}.pdf"))

                zip_file_name = filedialog.asksaveasfilename(defaultextension=".zip",
                                                             filetypes=[("ZIP Files", "*.zip")],
                                                             initialfile="NOTAS_DEBITO.zip")
                if zip_file_name:
                    with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for pdf_file in os.listdir(pasta_temp):
                            zipf.write(os.path.join(pasta_temp, pdf_file), pdf_file)
                    
                    shutil.rmtree(pasta_temp)
                    messagebox.showinfo("Sucesso", "ZIP gerado com sucesso!")

            loja_window.destroy()

        # Criar a janela de seleção
        loja_window = tk.Toplevel(tk_root)
        loja_window.title("Escolher Lojas")
        loja_window.geometry("400x500")
        loja_window.configure(bg="#f0f0f0")

        titulo_label = tk.Label(loja_window, text="Selecione as Lojas", font=("Arial", 14, "bold"), bg="#f0f0f0")
        titulo_label.pack(pady=10)

        checkbox_frame = tk.Frame(loja_window, bg="#f0f0f0")
        checkbox_frame.pack(pady=10, fill="both", expand=True)

        canvas = tk.Canvas(checkbox_frame, bg="#f0f0f0")
        scrollbar = tk.Scrollbar(checkbox_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="#f0f0f0")

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        checkboxes = {}
        for loja in lojas_disponiveis:
            var = tk.BooleanVar()
            checkbox = tk.Checkbutton(scroll_frame, text=loja, variable=var, font=("Arial", 12), bg="#f0f0f0")
            checkbox.pack(anchor="w")
            checkboxes[loja] = var

        var_selecionar_todas = tk.BooleanVar()
        selecionar_todas_btn = tk.Checkbutton(loja_window, text="Selecionar Todas", variable=var_selecionar_todas, 
                                              font=("Arial", 12, "bold"), bg="#f0f0f0", command=toggle_selecionar_todas)
        selecionar_todas_btn.pack(pady=5)

        confirmar_button = tk.Button(loja_window, text="Confirmar", command=lojas_selecionadas_callback, 
                                     font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", padx=10, pady=5)
        confirmar_button.pack(pady=10)

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
