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
from datetime import datetime, timedelta

mes_ano =  (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%m%Y")
# Caminho do arquivo onde o número da nota será armazenado

# Definir o caminho correto para armazenar o arquivo
APP_DIR = os.path.join(os.environ["USERPROFILE"], "Gerador Notas")
os.makedirs(APP_DIR, exist_ok=True)  # Criar pasta se não existir

ARQUIVO_NUMERO_NOTA = os.path.join(APP_DIR, "NFNumber.txt")

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

# Modo escuro
tema_fundo = "#2C2F33"
tema_texto = "#E0E0E0"
cor_botao = "#1E88E5"
cor_botao_hover = "#1565C0"

# Função para gerar o PDF
def gerar_pdf(dados, nome_arquivo):
    
    global numero_nota  # Usa a variável global
    
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
    numero_nota_str = str(numero_nota).zfill(5) 
    pdf.cell(60, 10, numero_nota_str, border=1, ln=True, align='C')
    pdf.set_xy(230, 40) 
    pdf.cell(60, 10, f"{dados['DATA DE EMISSÃO'].strftime('%d/%m/%Y')}", border=1, ln=True, align='C')
    pdf.set_xy(230, 50) 
    pdf.cell(60, 10, f"{dados['DATA DE PAGAMENTO'].strftime('%d/%m/%Y')}", border=1, ln=True, align='C')
    pdf.set_xy(230, 60) 
    pdf.cell(60, 10, f"{dados['DESCRIÇÃO']}", border=1, ln=True, align='C')
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
    pdf.cell(40, 5, f"{dados['DESCRIÇÃO']}", ln=True, align='C', border=1)
    pdf.set_xy(80, 181) 
    pdf.cell(30, 5, "1", ln=True, align='C', border=1)
    pdf.set_xy(110, 181) 
    pdf.cell(30, 5, "1", ln=True, align='C', border=1)
    pdf.set_xy(140, 181) 
    pdf.cell(50, 5, f"{dados['PRESTADOR DE SERVIÇO']}", ln=True, align='C', border=1)
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
    # pdf.cell(40, 5, f"R$ {dados['VALOR']:.2f}".replace('.', ','), ln=True, align='C')
    pdf.cell(40, 5, valor_formatado, ln=True, align='C')

    
    
    # Salvar o PDF temporariamente
    # Definir o diretório de trabalho no AppData
    APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Gerador_de_Notas_de_Debitos")
    TEMP_DIR = os.path.join(APP_DIR, "temp_pdfs")
    # Criar as pastas, se não existirem
    os.makedirs(TEMP_DIR, exist_ok=True)
    # Salvar o PDF temporariamente na pasta correta
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir=TEMP_DIR)
    pdf.output(temp_pdf.name)
    return temp_pdf.name

# ------------------------- Função para gerar um arquivo ZIP contendo todos os PDFs
def gerar_zip_com_pdfs(df):
    # Criar uma pasta temporária segura no diretório do usuário
    APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Gerador_de_Notas_de_Debitos")
    TEMP_DIR = os.path.join(APP_DIR, "temp_pdfs")
    os.makedirs(TEMP_DIR, exist_ok=True)  # Garantir que a pasta exista

    zip_name = tempfile.mktemp(suffix=".zip", dir=APP_DIR)
    
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for index, row in df.iterrows():
            pdf_path = gerar_pdf(row, os.path.join(TEMP_DIR, f"NOTA_DÉBITO_{row['LOJA']} {numero_nota-1}_{mes_ano}.pdf"))
            
            # Adicionar PDF ao ZIP
            zipf.write(pdf_path, f"NOTA DÉBITO - {row['LOJA']} {numero_nota-1}_{mes_ano}.pdf")
            os.remove(pdf_path)  # Remover o arquivo PDF após adicioná-lo ao ZIP
        
        # Remover a pasta temporária após a criação do ZIP
        shutil.rmtree(TEMP_DIR)
    
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
            colunas_necessarias = ["LOJA", "CNPJ", "ENDEREÇO", "CEP", "EMAIL", "VALOR", "DATA DE EMISSÃO", "DATA DE PAGAMENTO", 'PRESTADOR DE SERVIÇO','DESCRIÇÃO'] 
            colunas_necessarias_com_dados = ["LOJA", "CNPJ", "ENDEREÇO", "CEP", "VALOR", "DATA DE EMISSÃO", "DATA DE PAGAMENTO",'PRESTADOR DE SERVIÇO','DESCRIÇÃO'] 
            
            colunas_faltando = [col for col in colunas_necessarias if col not in df_global.columns]
            if colunas_faltando:
                messagebox.showerror("Erro", f"Faltam as seguintes colunas: {', '.join(colunas_faltando)}. Corrija os dados e reenvie o arquivo.")
            else:
                # Verifica se há valores nulos nas colunas obrigatórias
                colunas_com_nulos = [col for col in colunas_necessarias_com_dados if df_global[col].isna().sum() > 0]
                if colunas_com_nulos:
                    messagebox.showerror("Erro", f"As seguintes colunas possuem valores ausentes: {', '.join(colunas_com_nulos)}. Corrija os dados e reenvie o arquivo.")
                else:
                    messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")


# Função para gerar PDFs
def gerar_pdfs():
    global df_global
    if df_global is not None:
        zip_path = gerar_zip_com_pdfs(df_global)
        with open(zip_path, "rb") as f:
            zip_file_name = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP Files", "*.zip")],initialfile=f"NOTAS_DEBITO_{mes_ano}.zip")
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

            TEMP_DIR = os.path.join(tempfile.gettempdir(), "Gerador_Notas_Temp")
            os.makedirs(TEMP_DIR, exist_ok=True)  # Cria pasta temporária

            if len(lojas_selecionadas) == 1:
                loja = lojas_selecionadas[0]
                loja_df = df_global[df_global['LOJA'] == loja]
                
                if loja_df.empty:
                    messagebox.showerror("Erro", f"Dados não encontrados para a loja {loja}.")
                else:
                    pdf_path = gerar_pdf(loja_df.iloc[0], f"NOTA_DÉBITO_{loja}")
                    pdf_file_name = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        filetypes=[("PDF Files", "*.pdf")],
                        initialfile=f"NOTA DÉBITO - {loja} {numero_nota-1}_{mes_ano}.pdf"
                    )
                    if pdf_file_name:
                        shutil.move(pdf_path, pdf_file_name)
                        messagebox.showinfo("Sucesso", "PDF gerado com sucesso!")

            else:
                for loja in lojas_selecionadas:
                    loja_df = df_global[df_global['LOJA'] == loja]
                    if loja_df.empty:
                        messagebox.showerror("Erro", f"Dados não encontrados para a loja {loja}.")
                        continue
                    
                    pdf_path = gerar_pdf(loja_df.iloc[0], f"NOTA_DÉBITO_{loja}")
                    shutil.move(pdf_path, os.path.join(TEMP_DIR, f"NOTA DÉBITO - {loja} {numero_nota-1}_{mes_ano}.pdf"))

                zip_file_name = filedialog.asksaveasfilename(
                    defaultextension=".zip",
                    filetypes=[("ZIP Files", "*.zip")],
                    initialfile=f"NOTAS_DEBITO_{mes_ano}.zip"
                )

                if zip_file_name:
                    with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for pdf_file in os.listdir(TEMP_DIR):
                            zipf.write(os.path.join(TEMP_DIR, pdf_file), pdf_file)
                    
                    shutil.rmtree(TEMP_DIR)  # Remover arquivos temporários
                    messagebox.showinfo("Sucesso", "ZIP gerado com sucesso!")

            loja_window.destroy()

        # Criar a janela de seleção
        loja_window = tk.Toplevel(tk_root)
        loja_window.title("Escolher Lojas")
        loja_window.geometry("400x500")
        loja_window.configure(bg="#2C2F33")

        # Título
        titulo_label = tk.Label(loja_window, text="Selecione as Lojas", font=("Arial", 14, "bold"), bg="#2C2F33", fg="white")
        titulo_label.pack(pady=10)

        # Frame para checkboxes
        checkbox_frame = tk.Frame(loja_window, bg="#2C2F33")
        checkbox_frame.pack(pady=10, fill="both", expand=True)

        canvas = tk.Canvas(checkbox_frame, bg="#2C2F33")
        scrollbar = tk.Scrollbar(checkbox_frame, orient="vertical", command=canvas.yview, bg="#2C2F33")
        scroll_frame = tk.Frame(canvas, bg="#2C2F33")

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Checkbuttons para lojas
        checkboxes = {}
        for loja in lojas_disponiveis:
            var = tk.BooleanVar()
            checkbox = tk.Checkbutton(scroll_frame, text=loja, variable=var, font=("Arial", 12), bg="#2C2F33", fg="white", selectcolor="#2C2F33")
            checkbox.pack(anchor="w")
            checkboxes[loja] = var

        # Selecionar todas as lojas
        var_selecionar_todas = tk.BooleanVar()
        selecionar_todas_btn = tk.Checkbutton(loja_window, text="Selecionar Todas", variable=var_selecionar_todas,
                                            font=("Arial", 12, "bold"), bg="#2C2F33", fg="white", selectcolor="#2C2F33", command=toggle_selecionar_todas)
        selecionar_todas_btn.pack(pady=5)

        # Botão Confirmar
        confirmar_button = tk.Button(loja_window, text="Confirmar", command=lojas_selecionadas_callback,
                                    font=("Arial", 12, "bold"), bg="#1E88E5", fg="white", relief="raised", bd=2, padx=10, pady=5)
        confirmar_button.pack(pady=10)

    # Centralizar a janela de seleção


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
tk_root.geometry("400x250")
# # Carregar logo
# logo = tk.PhotoImage(file="logo Alife.png")  
# logo_label = tk.Label(tk_root, image=logo, bg=tema_fundo)
# logo_label.pack(pady=10)

# Centralizar a janela principal
tk_root.update_idletasks()
largura_janela = tk_root.winfo_width()
altura_janela = tk_root.winfo_height()
largura_tela = tk_root.winfo_screenwidth()
altura_tela = tk_root.winfo_screenheight()
x = (largura_tela - largura_janela) // 2
y = (altura_tela - altura_janela) // 2
tk_root.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")



# Configurar a janela
tk_root.configure(bg=tema_fundo)
frame = tk.Frame(tk_root, bg=tema_fundo)
frame.pack(padx=10, pady=10)

# Título
titulo_label = tk.Label(frame, text="Gerador de Notas de Débito", font=("Arial", 16, "bold"), bg=tema_fundo, fg=tema_texto)
titulo_label.pack(pady=10)

# Botões estilizados
def on_enter(e):
    e.widget.config(bg=cor_botao_hover)

def on_leave(e):
    e.widget.config(bg=cor_botao)

botao_carregar = tk.Button(frame, text="Carregar Arquivo Excel", command=carregar_arquivo, font=("Arial", 12), bg=cor_botao, fg="white", relief="raised", bd=2, padx=10, pady=5)
botao_carregar.pack(pady=10)
botao_carregar.bind("<Enter>", on_enter)
botao_carregar.bind("<Leave>", on_leave)

botao_gerar_pdfs = tk.Button(frame, text="Gerar PDFs", command=selecionar_tipo_geracao, font=("Arial", 12), bg=cor_botao, fg="white", relief="raised", bd=2, padx=10, pady=5)
botao_gerar_pdfs.pack(pady=10)
botao_gerar_pdfs.bind("<Enter>", on_enter)
botao_gerar_pdfs.bind("<Leave>", on_leave)

# Executar a interface gráfica
tk_root.mainloop()