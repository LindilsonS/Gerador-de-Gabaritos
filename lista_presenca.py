from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import os

def criar_lista_presenca(escola, turma, alunos, output_dir):
    """
    Gera um PDF com a lista de presença para uma turma específica.

    Args:
        escola (str): Nome da escola.
        turma (str): Nome da turma.
        alunos (list): Lista de nomes dos alunos.
        output_dir (str): Diretório onde o PDF será salvo.
    """
    # Configurar o caminho do arquivo
    caminho_arquivo = os.path.join(output_dir, f"lista_presenca_{turma}.pdf")
    c = canvas.Canvas(caminho_arquivo, pagesize=A4)
    largura, altura = A4

    # Definir margens
    margem_esquerda = 50
    margem_direita = largura - 50
    margem_superior = altura - 50
    margem_inferior = 50

    # Cabeçalho estilizado
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(largura / 2, margem_superior - 20, "SECRETARIA MUNICIPAL DE EDUCAÇÃO")
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura / 2, margem_superior - 50, f"Escola: {escola}")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(largura / 2, margem_superior - 80, f"Turma: {turma}")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(largura / 2, margem_superior - 110, "LISTA DE PRESENÇA")

    # Linha horizontal estilizada
    c.setLineWidth(2)
    c.setStrokeColor(colors.HexColor("#4CAF50"))  # Cor verde elegante
    c.line(margem_esquerda, margem_superior - 130, margem_direita, margem_superior - 130)

    # Configuração da tabela
    col_widths = [328.35, 200]  # Aumentar largura da coluna "Nome do Aluno" em 1 cm
    row_height = 40  # Altura das linhas
    data = [["Nome do Aluno", "Assinatura"]] + [[aluno, ""] for aluno in alunos]

    # Criar tabela
    table = Table(data, colWidths=col_widths, rowHeights=row_height)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E8F5E9")),  # Fundo do cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Cor do texto do cabeçalho
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinhamento centralizado
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Fonte do cabeçalho
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Fonte das linhas
        ('FONTSIZE', (0, 0), (-1, -1), 9),  # Ajustar o tamanho da fonte para 9
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Espaçamento inferior no cabeçalho
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),  # Grade da tabela
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Alinhamento vertical
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor("#4CAF50")),  # Linha abaixo do cabeçalho
    ])
    table.setStyle(style)

    # Calcular posição inicial da tabela
    altura_tabela = margem_superior - 150
    table.wrapOn(c, largura, altura)
    table.drawOn(c, margem_esquerda, altura_tabela - len(data) * row_height)

    # Salvar PDF
    c.save()
    print(f"Lista de presença salva: {caminho_arquivo}")
