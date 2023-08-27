import zipfile
import xml.etree.ElementTree as ET
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Inicializar a janela Tkinter
root = Tk()
root.withdraw()

# Abrir a janela de diálogo para selecionar o arquivo zip
arquivo_zip = askopenfilename(filetypes=[("Arquivos ZIP", "*.zip")])

# Verificar se o usuário selecionou um arquivo
if arquivo_zip:
    # Abrir o arquivo zip
    with zipfile.ZipFile(arquivo_zip, 'r') as arquivo_zip:
        # Iterar pelos arquivos no arquivo zip
        for nome_arquivo in arquivo_zip.namelist():
            # Verificar se é um arquivo XML
            if nome_arquivo.endswith('.xml'):
                # Extrair o arquivo do zip
                arquivo_xml = arquivo_zip.open(nome_arquivo)

                # Ler o conteúdo do arquivo XML
                conteudo_xml = arquivo_xml.read()

                # Analisar o arquivo XML
                raiz = ET.fromstring(conteudo_xml)

                # Dicionário para armazenar as informações desejadas
                informacoes = {
                    "{http://www.portalfiscal.inf.br/nfe}xNome": [],
                    "{http://www.portalfiscal.inf.br/cte}xNome": [],
                    "{http://www.portalfiscal.inf.br/nfe}IE": [],
                    "{http://www.portalfiscal.inf.br/cte}IE": []
                }

                # Localizar as tags e obter as informações correspondentes
                for tag, valor in informacoes.items():
                    elementos = raiz.findall(".//" + tag)
                    for elemento in elementos:
                        if elemento is not None and elemento.text:
                            informacoes[tag].append(elemento.text)

                # Exibir as informações no prompt de comando
                for tag, valores in informacoes.items():
                    print(tag + ":")
                    for valor in valores:
                        print(valor)
                    print("---")
else:
    print("Nenhum arquivo zip selecionado.")
