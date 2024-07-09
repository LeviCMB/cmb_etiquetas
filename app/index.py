from streamlit_gsheets import GSheetsConnection
from reportlab.pdfgen import canvas
import streamlit as st
import pdfplumber
import PyPDF2
import shutil
import math
import re
import io
import os

st.set_page_config(layout="wide", page_title="CMB Etiquetas")

data_fabricacao = st.date_input(label="Data de Fabricaçao dos Produtos:", format="DD/MM/YYYY")


itens_pedido = []
cliente = None
def extrair_cliente(conteudo_pdf):
    global cliente
    for linha in conteudo_pdf.split('\n'):
        if 'Cliente:' in linha:
            cliente = linha.split(':')[1].strip()
            return cliente
    return None


def extrair_itens_pedido(conteudo_pdf, pacote_dict):
    itens_pedido = []
    padrao_item = r"(\d+)\s*-\s*([\w\s&.,-/]+?)\s*(\d+)\s*(?:g'|G)?\s*(?:UN|Un|Und|und)"
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
    data_fabricacao = data_fabricacao.strftime("%d/%m/%Y")
    st.success(f"Data de Fabricaçao dos Produtos: :blue[{data_fabricacao}]") 
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

        # Tamanho da página em pontos (9.8cm de largura x 2.5cm de altura)

        if itens_pedido:
            st.write("Iniciando a geração de etiquetas...")
            # Gerar PDFs para cada item do pedido
            for item in itens_pedido:
                produto = item["produto"]
                quantidade = item["quantidade"]

                for i in range(quantidade):
                    fileName = f"{cliente}_{produto}_{i+1}.pdf".replace('/', '-').replace(' ', '_')
                    documentTitle = cliente
                    title = produto
                    subTitle = 'etiquetas'  
                    caminho_completo = os.path.join(pasta_destino, fileName)

                    pdf = canvas.Canvas(caminho_completo)

                    pdf.setPageSize((278, 71))
                    pdf.setTitle(documentTitle)
                    pdf.setTitle(title)
                    
                    
                    # Verificar se o produto existe no DataFrame
                    if produto in df_excel["Produto"].values:
                        descricao_produto = df_excel[df_excel ["Produto"] == produto]["Descricao"].values[0]
                    else:
                        # Usar o nome do produto como descrição padrão
                        descricao_produto = produto

                    regex = r"(?m)^(.*?)(?::|\.)\s*(.*?)(?::|\.)\s*(.*?)$"

                    match = re.search(regex, descricao_produto)
                    if match:
                        ingredientes = match.group(1).strip()
                        descricao = match.group(2).strip()
                        validade = match.group(3).strip()
                    else:
                        ingredientes = descricao_produto
                        descricao = ""
                        validade = ""

                    if not any(char.isdigit() for char in validade):
                        validade = 'Consumo Diário.'

                    if descricao == "Informações na Embalagem" or descricao == "":
                        pdf.setFont("Helvetica-Bold", 10)
                        pdf.drawCentredString(140, 50, title)

                        pdf.setFont("Helvetica", 10)
                        pdf.drawString(5, 5, f"{validade}")
                        pdf.setFont("Helvetica-Bold", 10)
                        pdf.drawString(200, 5, f"Fab: {data_fabricacao}")
                    else:
                        # Dividir a descrição em partes para o PDF
                        parte1 = descricao[:70].strip()
                        parte2 = descricao[70:].strip()
                    
                        pdf.setFont("Helvetica-Bold", 10)
                        pdf.drawCentredString(140, 60, title)
                        # Desenhar as partes do texto no PDF
                        pdf.setFont("Helvetica", 10)
                        pdf.drawCentredString(140, 45, f"{ingredientes}:")
                        
                        pdf.setFont("Helvetica", 8)
                        pdf.drawCentredString(140, 35, parte1)
                        pdf.drawCentredString(140, 25, parte2)

                        pdf.setFont("Helvetica", 10)
                        pdf.drawString(5, 5, f"{validade}")
                        pdf.setFont("Helvetica-Bold", 10)
                        pdf.drawString(200, 5, f"Fab: {data_fabricacao}")
                    
                    pdf.save()
            st.write("PDF DE CADA ETIQUETA GERADO COM SUCESSO!")    
        # Ordenar a lista de arquivos antes de combinar
        lista_arquivos = sorted(os.listdir(pasta_destino))
        merger = PyPDF2.PdfMerger()
        
        lista_arquivos = os.listdir("pedidos")
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
                    st.download_button(label="Clique aqui para baixar o PDF gerado", data=bytes, file_name=f"{cliente}_etiquetas.pdf".replace('/', '-').replace(' ', '_'))
                    
            else:
                st.error("O arquivo combinado não foi gerado corretamente.")
        st.text("")
        st.text("")

    # Adicionar botão para apagar as pastas após o processo
    if st.button("Finalizar Processos"):
                shutil.rmtree(pasta_destino)
                shutil.rmtree(pasta_destino_combinados)
                st.success("Processos Finalizados com Sucesso!")

st.write("##")
st.write("Desenvolvido por CMB Capital")
st.write("© 2024 CMB Capital. Todos os direitos reservados.")                
           