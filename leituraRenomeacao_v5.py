import zipfile
import xml.etree.ElementTree as ET
from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename, askdirectory
import os
import glob
from tkinter import messagebox

# Inicializar a janela Tkinter
root = Tk()
root.withdraw()
root.geometry('250x100')

# Solicitar ao usuário para inserir o nome a ser pesquisado
nome_pesquisado = simpledialog.askstring("Pesquisa", "Insira o nome a ser pesquisado")

# Verifique se o usuário clicou em 'Cancelar'
if nome_pesquisado is None:
    messagebox.showerror("Erro", "Processo interrompido pelo usuário.")
    exit()

# Abrir a janela de diálogo para selecionar o arquivo zip
arquivo_zip = askopenfilename(filetypes=[("Arquivos ZIP", "*.zip")])

# Verifique se o usuário clicou em 'Cancelar'
if not arquivo_zip:
    messagebox.showerror("Erro", "Processo interrompido pelo usuário.")
    exit()

# Função auxiliar para renomear o arquivo XML
def renomear_arquivo_xml(nome_arquivo):
    # Verificar se o arquivo possui a extensão .xml
    if nome_arquivo.endswith('.xml'):
        # Remover a extensão .xml
        nome_arquivo = nome_arquivo[:-4]
        if "-" in nome_arquivo:
            # Renomear o arquivo a partir do último "-"
            novo_nome = nome_arquivo.rsplit("-", 1)[1]
            # Adicionar novamente a extensão .xml
            return novo_nome + ".xml"
    return nome_arquivo

# Obter o diretório de saída para os arquivos de texto
diretorio_saida = askdirectory(title="Selecione o diretório de saída")

# Verificar se o diretório de saída foi selecionado
if not diretorio_saida:
    messagebox.showerror("Erro", "Processo interrompido pelo usuário.")
    exit()

# Abrir o arquivo zip
with zipfile.ZipFile(arquivo_zip, 'r') as arquivo_zip:
    # Dicionário para armazenar os nomes IE que correspondem ao nome pesquisado
    correspondentes = {}


    def buscar_toma_e_pai(elemento):
        # Verificar a presença de toma3 ou toma4 e pegar o valor da tag <toma> dentro delas
        toma_tag = None
        toma_value = None

        if elemento.find("toma3/toma") is not None:
            toma_tag = "toma3"
            toma_value = elemento.find("toma3/toma").text
        elif elemento.find("toma4/toma") is not None:
            toma_tag = "toma4"
            toma_value = elemento.find("toma4/toma").text

        if toma_value is not None and toma_value.isdigit():
            numero_toma = int(toma_value)

            # Dicionário para mapear o número à tag correspondente
            mapeamento_toma = {
                0: "rem",
                1: "exped",
                2: "receb",
                3: "dest",
                4: "outros"
            }

            # Retorna a tag pai (toma3 ou toma4) e a tag correspondente ao número encontrado em toma
            tag_correspondente = mapeamento_toma.get(numero_toma, None)
            pai, tag_correspondente = buscar_toma_e_pai(elemento)
            if pai and tag_correspondente:
                print(f"Tag pai: {pai}, Tag correspondente: {tag_correspondente}")
            return toma_tag, tag_correspondente

        return None, None


    # Função para pesquisar o nome em todas as tags de um elemento
    def buscar_nome(elemento, nome):
        for filho in elemento:
            if filho.text is not None and nome.lower() in filho.text.lower():

                tag_pai = buscar_toma_e_pai(filho)

                # Verificar se a tag pai é aquela em que estamos interessados
                if tag_pai == "dest":
                    ie_value = buscar_ie(filho)
                    if ie_value:
                        # Criar pasta "Cliente - Tomador"
                        nome_diretorio = "Cliente - Tomador"
                        caminho_diretorio = os.path.join(diretorio_saida, nome_diretorio)
                        os.makedirs(caminho_diretorio, exist_ok=True)

                        # Armazene o nome do arquivo XML para posterior extração
                        return caminho_diretorio
            resultado = buscar_nome(filho, nome)
            if resultado is not None:
                return resultado
        return None

    # Função para buscar a tag IE de forma recursiva
    def buscar_ie(elemento):
        for filho in elemento:
            if filho.tag.endswith('IE'):  # Procura por qualquer tag que termine com 'IE'
                return filho.text
            resultado = buscar_ie(filho)
            if resultado is not None:
                return resultado
        return None

    # Primeiro, coletar todos os nomes IE correspondentes
    for nome_arquivo in arquivo_zip.namelist():
        if nome_arquivo.endswith('.xml'):
            arquivo_xml = arquivo_zip.open(nome_arquivo)
            conteudo_xml = arquivo_xml.read()
            raiz = ET.fromstring(conteudo_xml)

            grupo = buscar_nome(raiz, nome_pesquisado)
            if grupo is not None:
                IE = buscar_ie(grupo)  # Procura a tag IE dentro do grupo de tags
                if IE is not None:
                    correspondentes[IE] = True

    def buscar_ies_no_elemento(elemento, correspondentes):
        ies_encontrados = []
        for filho in elemento:
            if filho.tag.endswith('IE'):
                if filho.text in correspondentes:
                    ies_encontrados.append(filho.text)
            ies_encontrados.extend(buscar_ies_no_elemento(filho, correspondentes))
        return ies_encontrados

    # Em seguida, extrair todos os arquivos XML que correspondem aos nomes IE coletados
    for nome_arquivo in arquivo_zip.namelist():
        if nome_arquivo.endswith('.xml'):
            arquivo_xml = arquivo_zip.open(nome_arquivo)
            conteudo_xml = arquivo_xml.read()
            raiz = ET.fromstring(conteudo_xml)

            ies_encontrados = buscar_ies_no_elemento(raiz, correspondentes)
            grupo = buscar_nome(raiz, nome_pesquisado)

            if ies_encontrados:
                # Exemplo de manipulação quando os IE's são encontrados.
                # Você pode ajustar isso de acordo com suas necessidades.
                for IE in ies_encontrados:
                    # Criar o diretório de saída com o nome IE se não existir
                    diretorio_saida_ie = os.path.join(diretorio_saida, IE)
                    os.makedirs(diretorio_saida_ie, exist_ok=True)

                    # Extrair o arquivo XML para o diretório de saída
                    nome_arquivo_saida = os.path.join(diretorio_saida_ie, nome_arquivo)
                    nome_arquivo_saida = os.path.normpath(nome_arquivo_saida)

                    # Criar os diretórios necessários, se eles não existirem
                    os.makedirs(os.path.dirname(nome_arquivo_saida), exist_ok=True)

                    # Criar um arquivo XML para escrever o conteúdo
                    with open(nome_arquivo_saida, 'wb') as arquivo_saida:
                        arquivo_saida.write(conteudo_xml)
            elif grupo is None:
                # Caso não encontre a pesquisa realizada, criar uma pasta chamada "Sem identificação" e salvar o arquivo XML nela
                diretorio_saida_sem_identificacao = os.path.join(diretorio_saida, "Sem identificação")
                os.makedirs(diretorio_saida_sem_identificacao, exist_ok=True)

                nome_arquivo_saida = os.path.join(diretorio_saida_sem_identificacao, nome_arquivo)
                nome_arquivo_saida = os.path.normpath(nome_arquivo_saida)

                # Criar os diretórios necessários, se eles não existirem
                os.makedirs(os.path.dirname(nome_arquivo_saida), exist_ok=True)

                # Criar um arquivo XML para escrever o conteúdo
                with open(nome_arquivo_saida, 'wb') as arquivo_saida:
                    arquivo_saida.write(conteudo_xml)
            else:
                # Caso não encontre IE's, criar uma pasta chamada "NFe sem IE" e salvar o arquivo XML nela
                diretorio_saida_sem_ie = os.path.join(diretorio_saida, "NF sem IE")
                os.makedirs(diretorio_saida_sem_ie, exist_ok=True)

                nome_arquivo_saida = os.path.join(diretorio_saida_sem_ie, nome_arquivo)
                nome_arquivo_saida = os.path.normpath(nome_arquivo_saida)

                # Criar os diretórios necessários, se eles não existirem
                os.makedirs(os.path.dirname(nome_arquivo_saida), exist_ok=True)
                # Criar um arquivo XML para escrever o conteúdo
                with open(nome_arquivo_saida, 'wb') as arquivo_saida:
                    arquivo_saida.write(conteudo_xml)

print("Arquivos XML extraídos com sucesso para o diretório:", diretorio_saida)

# Após salvar todos os arquivos XML, percorrer todos os diretórios dentro da pasta de saída e renomear cada arquivo XML
for nome_arquivo in glob.glob(diretorio_saida + "/**/*.xml", recursive=True):
    novo_nome_arquivo = renomear_arquivo_xml(os.path.basename(nome_arquivo))
    novo_caminho_arquivo = os.path.join(os.path.dirname(nome_arquivo), novo_nome_arquivo)
    try:
        os.rename(nome_arquivo, novo_caminho_arquivo)
    except Exception as e:
        print(f"Não foi possível renomear o arquivo {nome_arquivo} para {novo_nome_arquivo}. Erro: {str(e)}")

# Apresentar uma janela informando que o processo foi concluído
messagebox.showinfo("Processo Concluído", "O processo da automação foi concluído com sucesso!")
