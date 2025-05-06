import os
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from lxml import etree 
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle


def sanitizar_nome(nome):
    caracteres_invalidos = r'[<>:"/\\|?*]'
    return re.sub(caracteres_invalidos, '_', nome).strip()

# Função para substituir variáveis em parágrafos, tabelas e caixas de texto
def substituir_variaveis(doc, dados_aluno):
    substituicoes_feitas = 0
    
    for paragraph in doc.paragraphs:
        texto_original = paragraph.text
        for chave, valor in dados_aluno.items():
            marcador = f"$VARIÁVEL {chave}"
            if marcador in texto_original:
                paragraph.text = texto_original.replace(marcador, str(valor))
                substituicoes_feitas += 1
                print(f"Substituído '{marcador}' por '{valor}' em parágrafo do corpo")

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto_original = cell.text
                for chave, valor in dados_aluno.items():
                    marcador = f"$VARIÁVEL {chave}"
                    if marcador in texto_original:
                        cell.text = texto_original.replace(marcador, str(valor))
                        substituicoes_feitas += 1
                        print(f"Substituído '{marcador}' por '{valor}' em tabela")

    for txbx in doc.element.body.findall('.//w:txbxContent', namespaces=doc.element.nsmap):
        for paragraph in txbx.findall('.//w:p', namespaces=doc.element.nsmap):
            texto_original = ''
            runs = paragraph.findall('.//w:r', namespaces=doc.element.nsmap)
            for run in runs:
                texto_original += ''.join(t.text for t in run.findall('.//w:t', namespaces=doc.element.nsmap))
            
            for chave, valor in dados_aluno.items():
                marcador = f"$VARIÁVEL {chave}"
                if marcador in texto_original:
                    novo_texto = texto_original.replace(marcador, str(valor))
                    for run in runs:
                        for t in run.findall('.//w:t', namespaces=doc.element.nsmap):
                            t.text = ''
                    if runs:
                        first_run = runs[0]
                        t_elements = first_run.findall('.//w:t', namespaces=doc.element.nsmap)
                        if t_elements:
                            t_elements[0].text = novo_texto
                        else:
                            t = first_run.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                            t.text = novo_texto
                            first_run.append(t)
                    else:
                        new_run = paragraph.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                        t = new_run.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                        t.text = novo_texto
                        new_run.append(t)
                        paragraph.append(new_run)
                    substituicoes_feitas += 1
                    print(f"Substituído '{marcador}' por '{valor}' em caixa de texto")

    if substituicoes_feitas == 0:
        print("Nenhuma substituição foi feita. Verifique os marcadores no modelo!")
    else:
        print(f"Total de substituições realizadas: {substituicoes_feitas}")
    
    return doc


def criar_lista_presenca(escola, turma, alunos, turma_dir):
    """
    Gera um PDF com a lista de presença para uma turma específica.

    Args:
        escola (str): Nome da escola.
        turma (str): Nome da turma.
        alunos (list): Lista de nomes dos alunos.
        turma_dir (str): Diretório onde o PDF será salvo.
    """

    alunos = sorted(alunos)

    caminho_arquivo = os.path.join(turma_dir, f"lista_presenca_{turma}.pdf")
    c = canvas.Canvas(caminho_arquivo, pagesize=A4)
    largura, altura = A4

  
    margem_esquerda = 40
    margem_direita = largura - 40
    margem_superior = altura - 40
    margem_inferior = 100  

    col_width_num = 28  
    col_width_nome = 228.35  
    col_width_assinatura = 250 
    col_widths = [col_width_num, col_width_nome, col_width_assinatura]
    row_height = 30  
    font_size = 12 
    max_rows_per_page = int((margem_superior - margem_inferior - 70) / row_height)

  
    def desenhar_cabecalho_rodape():
        c.setFont("Helvetica-Bold", font_size + 8)
        c.setFillColor(colors.HexColor("#2C3E50"))  
        c.drawCentredString(largura / 2, margem_superior, "Lista de Presença")
        c.setFont("Helvetica", font_size)
        c.setFillColor(colors.HexColor("#34495E"))  
        c.drawCentredString(largura / 2, margem_superior - 20, f"Escola: {escola} | Turma: {turma}")
        c.setFont("Helvetica", font_size)
        c.drawString(margem_esquerda, margem_superior - 40, "Data: ____________")
        c.setLineWidth(1)
        c.setStrokeColor(colors.HexColor("#ee91ed"))  
        c.line(margem_esquerda, margem_superior - 50, margem_direita, margem_superior - 50)
        c.setFont("Helvetica", font_size)
        c.setFillColor(colors.HexColor("#000000")) 
        c.drawString(margem_esquerda, margem_inferior + 20, f"Total de alunos na turma: {len(alunos)}")
        c.drawString(margem_esquerda, margem_inferior, "Total de alunos presentes: ________")
        c.drawCentredString(largura / 2, margem_inferior - 20, "SECRETARIA MUNICIPAL DE EDUCAÇÃO")

    # Dividir os alunos em páginas
    for page, start in enumerate(range(0, len(alunos), max_rows_per_page)):
        if page > 0:
            c.showPage()  
        desenhar_cabecalho_rodape()

        # Dados da tabela para a página 
        data = [["Nº", "Nome do Aluno", "Assinatura"]] + [
            [str(i + 1 + start), aluno, ""] for i, aluno in enumerate(alunos[start:start + max_rows_per_page])
        ]

      
        table = Table(data, colWidths=col_widths, rowHeights=row_height)
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#ee91ed")),  
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#FFFFFF")),  
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#ECF0F1")),  
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor("#000000")), 
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), 
            ('FONTSIZE', (0, 0), (-1, -1), 9),  
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),  
            ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor("#000000")),  
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  
        ])
        table.setStyle(style)

    
        altura_tabela = margem_superior - 70
        table.wrapOn(c, largura, altura)
        table.drawOn(c, margem_esquerda, altura_tabela - len(data) * row_height)

  
    c.save()
    print(f"Lista de presença salva: {caminho_arquivo}")


def criar_gabaritos(csv_path, modelo_path, output_dir, config):
    try:
       
        df = pd.read_csv(csv_path, sep=';', encoding='utf-8')

       
        colunas_necessarias = {'ESCOLA', 'TURMA', 'NOME DO ALUNO', 'PROFESSOR REGENTE'}
        if not colunas_necessarias.issubset(df.columns):
            raise ValueError(f"O arquivo CSV deve conter as colunas: {', '.join(colunas_necessarias)}")


        grouped = df.groupby(['ESCOLA', 'TURMA'])
        
        for (escola, turma), group in grouped:
            escola = sanitizar_nome(escola)
            turma = sanitizar_nome(turma)
            
         
            escola_dir = os.path.join(output_dir, escola)
            os.makedirs(escola_dir, exist_ok=True)
            
            
            turma_dir = os.path.join(escola_dir, turma)
            os.makedirs(turma_dir, exist_ok=True)
            
            
            alunos = group['NOME DO ALUNO'].tolist()
            professor_regente = group['PROFESSOR REGENTE'].iloc[0] 
            
            criar_lista_presenca(escola, turma, alunos, turma_dir)
            
            
            for aluno in alunos:
                aluno = sanitizar_nome(aluno)
                
                doc = Document(modelo_path)
                
                dados_aluno = {
                    'NOME DO ALUNO': aluno,
                    'NOME DO ALUNO 2': aluno, 
                    'ESCOLA': escola,
                    'TURMA': turma,
                    'PROFESSOR REGENTE': professor_regente
                }
                
                doc = substituir_variaveis(doc, dados_aluno)
                
                nome_arquivo = f"{aluno}_gabarito.docx"
                caminho_arquivo = os.path.join(turma_dir, nome_arquivo)
                doc.save(caminho_arquivo)
                print(f"Arquivo salvo: {caminho_arquivo}")
        
        return True
    except Exception as e:
        print(f"Erro ao criar gabaritos: {str(e)}")
        return False

# Interface gráfica
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Gabaritos e Lista de Presença")
        self.root.geometry("400x350")  
        
        self.csv_path = None
        self.modelo_path = None
        self.output_dir = None
        self.usar_somente_aluno1 = tk.BooleanVar(value=False)  
        
        tk.Label(root, text="Gerador de Gabaritos e Lista de Presença", font=("Arial", 14)).pack(pady=10)
        
        tk.Button(root, text="Selecionar CSV", command=self.selecionar_csv).pack(pady=5)
        self.csv_label = tk.Label(root, text="Nenhum CSV selecionado")
        self.csv_label.pack()
        
        tk.Button(root, text="Selecionar Modelo Word", command=self.selecionar_modelo).pack(pady=5)
        self.modelo_label = tk.Label(root, text="Nenhum modelo selecionado")
        self.modelo_label.pack()
        
        tk.Button(root, text="Selecionar Pasta de Saída", command=self.selecionar_output).pack(pady=5)
        self.output_label = tk.Label(root, text="Nenhuma pasta selecionada")
        self.output_label.pack()
        
        # Checkbox para ativar/desativar o uso de apenas o aluno 1
        tk.Checkbutton(root, text="Usar apenas o Aluno 1", variable=self.usar_somente_aluno1).pack(pady=10)
        
        tk.Button(root, text="Gerar Documentos", command=self.gerar).pack(pady=20)
    
    def selecionar_csv(self):
        self.csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.csv_path:
            self.csv_label.config(text=os.path.basename(self.csv_path))
    
    def selecionar_modelo(self):
        self.modelo_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if self.modelo_path:
            self.modelo_label.config(text=os.path.basename(self.modelo_path))
    
    def selecionar_output(self):
        self.output_dir = filedialog.askdirectory()
        if self.output_dir:
            self.output_label.config(text=self.output_dir)
    
    def gerar(self):
        if not self.csv_path or not self.modelo_path or not self.output_dir:
            messagebox.showerror("Erro", "Selecione todos os arquivos e a pasta de saída!")
            return
        
        try:
            
            config = Configuracao(usar_somente_aluno1=self.usar_somente_aluno1.get())
            sucesso = criar_gabaritos(self.csv_path, self.modelo_path, self.output_dir, config)
            if sucesso:
                messagebox.showinfo("Sucesso", "Gabaritos gerados com sucesso!")
            else:
                messagebox.showerror("Erro", "Falha ao gerar os documentos. Veja o console.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

class Configuracao:
    def __init__(self, usar_somente_aluno1=False):
        self.usar_somente_aluno1 = usar_somente_aluno1

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()