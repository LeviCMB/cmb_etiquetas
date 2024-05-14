from streamlit_gsheets import GSheetsConnection
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import portrait
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
import streamlit as st
import pandas as pd
import pdfplumber
import textwrap
import PyPDF2
import shutil
import math
import re
import io
import os

# Registrar as fontes
pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Bold', 'Arialbd.ttf'))

st.set_page_config(layout="wide", page_title="CMB Etiquetas")

itens_pedido = []
def extrair_cliente(conteudo_pdf):
    cliente = None
    nome_fantasia = None
    for linha in conteudo_pdf.split('\n'):
        if 'Cliente:' in linha:
            cliente = linha.split(':')[1].split('Nome Fantasia:')[0].strip()
            return cliente
        elif linha.startswith('Nome Fantasia:'):
            nome_fantasia = linha.split(':')[1].strip()
            if cliente and nome_fantasia:
                return f"{cliente} / {nome_fantasia}"
    return None

def extrair_itens_pedido(conteudo_pdf, pacote_dict):
    itens_pedido = []
    padrao_item = r"(\d+)\s*-\s*([\w\s&.,-/]+?)\s*(\d+)\s*(?:g|G)?\s*(?:UN|Un|Und|und)"
    for match in re.finditer(padrao_item, conteudo_pdf):
        produto = match.group(2).strip()
        quantidade = int(match.group(3))
        if produto in pacote_dict:
            valor_pacote = pacote_dict[produto]
            etiquetas_necessarias = math.ceil(quantidade / valor_pacote)
            itens_pedido.append({'produto': produto, 'quantidade': etiquetas_necessarias})
        else:
            itens_pedido.append({'produto': produto, 'quantidade': quantidade})
    return itens_pedido


# Interface

url = "https://docs.google.com/spreadsheets/d/10xH-WrGzH3efBqlrrUvX4kHotmL-sX19RN3_dn5YqyA/edit?usp=sharing"

conn = st.connection("gsheets", type=GSheetsConnection)

with st.sidebar:
    st.header("GERADOR DE ETIQUETAS CMB")
    arquivo_pedido = st.file_uploader(label="Arraste ou Selecione o Arquivo em PDF do Pedido:", type=['pdf'])

if arquivo_pedido:
    arquivo_pedido_bytes = io.BytesIO(arquivo_pedido.read())
    with pdfplumber.open(arquivo_pedido_bytes) as pdf:
        conteudo_pdf = ""
        for pagina in pdf.pages:
            conteudo_pdf += pagina.extract_text()

        cliente = extrair_cliente(conteudo_pdf)
        if cliente:
            st.success(f"Cliente identificado: {cliente}")
        else:
            st.error("Nenhum cliente identificado no PDF.")

        # Carregar base de dados dos produtos e extrair itens do pedido
        df_excel = conn.read(spreadsheet=url)
        pacote_dict = dict(zip(df_excel["Produto"], df_excel["ProdutoPacote"]))
        itens_pedido = extrair_itens_pedido(conteudo_pdf, pacote_dict)

# Diretório para salvar os PDFs
pasta_destino = "pedidos"

# Criar o diretório se não existir
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)
else:
    # Limpar a pasta "pedidos" se já existir
    shutil.rmtree(pasta_destino)
    os.makedirs(pasta_destino)

# Diretório para salvar os PDFs
# pasta_destino = "pedidos"

# Criar o diretório se não existir
# if not os.path.exists(pasta_destino):
 #   os.makedirs(pasta_destino)

# Tamanho da página em pontos (9.8cm de largura x 2.5cm de altura)
largura_cm = 9.8
altura_cm = 2.5
largura_pagina = largura_cm * 28.35  # 1cm = 28.35 pontos
altura_pagina = altura_cm * 28.35
pagesize = (largura_pagina, altura_pagina)

if itens_pedido: 
    st.write("Iniciando a geração de etiquetas...")
    # Gerar PDFs para cada item do pedido
    for item in itens_pedido:
        produto = item["produto"]
        quantidade = item["quantidade"]
        
        for i in range(quantidade):
            # Delimitando a largura do texto
            linhas_produto = textwrap.wrap(produto, width=15)  # Ajuste o valor de width conforme necessário
            
            # Definindo o nome do arquivo
            nome_arquivo = f"{cliente}_{produto}_{i+1}.pdf".replace('/', '-').replace(' ', '_')
            caminho_completo = os.path.join(pasta_destino, nome_arquivo)
            
            # Criar o arquivo PDF com a página configurada
            c = canvas.Canvas(caminho_completo, pagesize=pagesize)
            c.setFont("Arial", 10)  # Fonte e tamanho ajustáveis conforme necessário

            # Verificar se o produto existe no DataFrame
            if produto in df_excel["Produto"].values:
                descricao_produto = df_excel[df_excel["Produto"] == produto]["Descricao"].values[0]
            else:
                # Usar o nome do produto como descrição padrão
                descricao_produto = produto

            # Concatenar o nome do produto com a descrição
            texto_completo = f"{produto} - {descricao_produto}"

            # Converter o texto completo em uma string
            texto_completo = str(texto_completo)

            # Dividir o texto completo em linhas com quebra de linha a partir da largura
            linhas_texto = textwrap.wrap(texto_completo, width=25)  # Ajuste o valor de width conforme necessário

            # Calcular a altura total do texto
            altura_total_texto = len(linhas_produto) * 1  # Considerando 10 pontos de altura para cada linha de texto
            
            # Posicionar o texto verticalmente ao centro da página
            posicao_y = (altura_pagina - altura_total_texto) / 2 + (len(linhas_produto) - 1) * 5

            # Desenhar o texto completo
            for linha in linhas_texto:
                largura_texto = c.stringWidth(linha, "Arial", 10)
                posicao_x = (largura_pagina - largura_texto) / 2
                c.drawString(posicao_x, posicao_y, linha)
                posicao_y -= 10  # Descer para a próxima linha
            
            c.save()
    st.write("PDF DE CADA ETIQUETA GERADO COM SUCESSO!")    


merger = PyPDF2.PdfMerger()

lista_arquivos = os.listdir("pedidos")
lista_arquivos.sort()
for arquivo in lista_arquivos:
    if ".pdf" in arquivo:
        caminho_arquivo = os.path.join("pedidos", arquivo)
        if os.path.isfile(caminho_arquivo):  # Verifica se é um arquivo válido
            merger.append(caminho_arquivo)

# Diretório para salvar o PDF combinado
pasta_destino_combinados = "pedidos_combinados"

# Criar o diretório se não existir
if not os.path.exists(pasta_destino_combinados):
    os.makedirs(pasta_destino_combinados)

# Definir o caminho completo para o arquivo PDF combinado
arquivo_combinado = os.path.join(pasta_destino_combinados, f"{cliente}_etiquetas.pdf".replace('/', '-').replace(' ', '_'))

# Escrever o PDF combinado em um novo arquivo
merger.write(arquivo_combinado)
merger.close()  # Fechar o arquivo após a escrita

# Fazer o download do arquivo
if st.button(label="Preparar o Download"):
    if os.path.exists(arquivo_combinado):  # Verifica se o arquivo combinado existe
        with open(arquivo_combinado, "rb") as file:
            bytes = file.read()
            st.download_button(label="Clique aqui para baixar o PDF gerado", data=bytes, file_name=arquivo_combinado)
    else:
        st.error("O arquivo combinado não foi gerado corretamente.")
st.text("")
st.text("")

# Adicionar botão para apagar as pastas após o processo
if st.button("Finalizar Processos"):
    shutil.rmtree(pasta_destino)
    shutil.rmtree(pasta_destino_combinados)
    st.success("Processos Finalizados com Sucesso!")

           