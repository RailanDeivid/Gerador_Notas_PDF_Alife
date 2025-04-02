from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox, QInputDialog)
from fpdf import FPDF
import pandas as pd
import numpy as np
import tempfile
import zipfile
import os
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QPalette

        
# Caminho do arquivo onde o número da nota será armazenado
ARQUIVO_NUMERO_NOTA = "NFNumber.txt"

def carregar_numero_nota():
    if os.path.exists(ARQUIVO_NUMERO_NOTA):
        with open(ARQUIVO_NUMERO_NOTA, "r") as f:
            return int(f.read().strip())
    return 1

def salvar_numero_nota(numero):
    with open(ARQUIVO_NUMERO_NOTA, "w") as f:
        f.write(str(numero))

numero_nota = carregar_numero_nota()

class NotaDebitoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Selecione um arquivo Excel contendo os dados:")
        layout.addWidget(self.label)

        self.btn_selecionar = QPushButton("Selecionar Arquivo")
        self.btn_selecionar.clicked.connect(self.selecionar_arquivo)
        layout.addWidget(self.btn_selecionar)

        self.btn_gerar = QPushButton("Gerar Notas de Débito (Todas)")
        self.btn_gerar.clicked.connect(self.gerar_notas)
        layout.addWidget(self.btn_gerar)

        self.btn_gerar_individual = QPushButton("Gerar Nota Individual")
        self.btn_gerar_individual.clicked.connect(self.gerar_nota_individual)
        layout.addWidget(self.btn_gerar_individual)

        self.setLayout(layout)
        self.setWindowTitle("Gerador de Notas de Débito")
        self.setGeometry(100, 100, 400, 200)

    def selecionar_arquivo(self):
        options = QFileDialog.Options()
        arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo Excel", "", "Excel Files (*.xlsx;*.xls)", options=options)
        if arquivo:
            self.label.setText(f"{os.path.basename(arquivo)} Selecionado.")
            self.arquivo_excel = arquivo
            
    def gerar_notas(self):
        """Gera todas as notas de débito e compacta em um ZIP, permitindo ao usuário escolher o local de salvamento."""
        if not hasattr(self, 'arquivo_excel'):
            QMessageBox.warning(self, "Erro", "Nenhum arquivo selecionado!")
            return

        df = pd.read_excel(self.arquivo_excel)

        # Pergunta ao usuário onde salvar o ZIP
        options = QFileDialog.Options()
        zip_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar Arquivo ZIP", "", "Arquivos ZIP (*.zip);;Todos os Arquivos (*)", options=options
        )

        if not zip_path:  # Se o usuário cancelar, não faz nada
            return

        if not zip_path.endswith(".zip"):  # Garante que o arquivo tenha a extensão .zip
            zip_path += ".zip"

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for index, row in df.iterrows():
                pdf_path = self.gerar_pdf(row)
                zipf.write(pdf_path, f"Nota de Debito - {row['LOJA']}.pdf")
                os.remove(pdf_path)

        QMessageBox.information(self, "Sucesso", f"Notas de débito salvas em:\n{zip_path}")

    def gerar_nota_individual(self):
        """Gera uma ou mais notas de débito individualmente, permitindo escolher onde salvar."""
        if not hasattr(self, 'arquivo_excel'):
            QMessageBox.warning(self, "Erro", "Nenhum arquivo selecionado!")
            return

        df = pd.read_excel(self.arquivo_excel)

        # Lista de lojas únicas
        lojas = df["LOJA"].dropna().unique().tolist()

        if not lojas:
            QMessageBox.warning(self, "Erro", "Nenhuma loja encontrada no arquivo!")
            return

        # Perguntar ao usuário qual loja selecionar
        loja, ok = QInputDialog.getItem(self, "Escolher Loja", "Selecione uma loja:", lojas, 0, False)

        if ok and loja:
            # Filtra os dados para a loja escolhida
            df_loja = df[df["LOJA"] == loja]
            
            # Perguntar onde salvar os arquivos
            pasta_destino = QFileDialog.getExistingDirectory(self, "Escolher Diretório para Salvar a Nota")
            if not pasta_destino:
                return
            
            # Garantir que o caminho da pasta está correto
            pasta_destino = os.path.abspath(pasta_destino)

            if df_loja.shape[0] > 1:  # Se houver mais de uma linha para a mesma loja
                zip_path = os.path.join(pasta_destino, f"Notas de Debito - {loja}.zip")
                
                try:
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for _, row in df_loja.iterrows():
                            pdf_path = os.path.join(pasta_destino, f"Nota de Debito - {loja}.pdf")
                            self.gerar_pdf(row, pdf_path)  # Passa o caminho para a função gerar_pdf
                            
                            if os.path.exists(pdf_path):  # Verifica se o PDF foi realmente gerado
                                zipf.write(pdf_path, os.path.basename(pdf_path))
                                os.remove(pdf_path)  # Remove o PDF após adicionar ao ZIP

                    QMessageBox.information(self, "Sucesso", f"Arquivo ZIP gerado em: {zip_path}")
                except Exception as e:
                    QMessageBox.critical(self, "Erro", f"Erro ao criar o ZIP: {str(e)}")
            else:
                pdf_path = os.path.join(pasta_destino, f"Nota de Debito - {loja}.pdf")
                
                try:
                    self.gerar_pdf(df_loja.iloc[0], pdf_path)  # Passa o caminho para a função gerar_pdf
                    if os.path.exists(pdf_path):
                        QMessageBox.information(self, "Sucesso", f"Nota gerada em: {pdf_path}")
                    else:
                        QMessageBox.warning(self, "Erro", "O PDF não foi gerado corretamente!")
                except Exception as e:
                    QMessageBox.critical(self, "Erro", f"Erro ao gerar o PDF: {str(e)}")

    def gerar_pdf(self, dados, pdf_path):
        
        global numero_nota  # Usa a variável global
         # Criar a pasta de saída se não existir
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
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
        numero_nota_str = str(numero_nota).zfill(5) 
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
        # pdf.cell(40, 5, f"R$ {dados['VALOR']:.2f}".replace('.', ','), ln=True, align='C')
        pdf.cell(40, 5, valor_formatado, ln=True, align='C')

        # Salvando o PDF
        try:
            pdf.output(pdf_path)
            print(f"PDF gerado com sucesso: {pdf_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o PDF: {str(e)}")

        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf.output(temp_pdf.name)
        return temp_pdf.name



if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = NotaDebitoApp()
    window.show()
    sys.exit(app.exec_())
    
    