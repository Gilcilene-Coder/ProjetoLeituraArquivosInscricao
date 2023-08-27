import zipfile
import xml.etree.ElementTree as ET
from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, askdirectory
import os
import glob

def solicitar_entrada(titulo, mensagem):
    valor = simpledialog.askstring(titulo, mensagem)
    if valor is None:
        messagebox.showerror("Erro", "Processo interrompido pelo usuário.")
        exit()
    return valor

def salvar_arquivo_em_diretorio(conteudo, diretorio, nome_arquivo):
    nome_arquivo_saida = os.path.join(diretorio, nome_arquivo)
    os.makedirs(os.path.dirname(nome_arquivo_saida), exist_ok=True)
    with open(nome_arquivo_saida, 'wb') as arquivo_saida:
        arquivo_saida.write(conteudo)

def buscar_tags_toma(raiz_xml):
    namespaces = {"ns": "http://www.portalfiscal.inf.br/cte"}
    tags_toma = ["toma0", "toma1", "toma2", "toma3", "toma4"]
    for tag in tags_toma:
        elemento = raiz_xml.find(f".//ns:{tag}/ns:toma", namespaces=namespaces)
        if elemento is not None:
            return tag
    return None

def buscar_nome(elemento, nome):
    for filho in elemento:
        if filho.text is not None and nome.lower() in filho.text.lower():
            return elemento  # retorna o elemento pai
        resultado = buscar_nome(filho, nome)
        if resultado is not None:
            return resultado
    return None

def buscar_ie(elemento):
    for filho in elemento:
        if filho.tag.endswith('IE'):  # Procura por qualquer tag que termine com 'IE'
            return filho.text
        resultado = buscar_ie(filho)
        if resultado is not None:
            return resultado
    return None

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

def buscar_ies_no_elemento(elemento, correspondentes):
    ies_encontrados = []
    for filho in elemento:
        if filho.tag.endswith('IE'):
            if filho.text in correspondentes:
                ies_encontrados.append(filho.text)
        ies_encontrados.extend(buscar_ies_no_elemento(filho, correspondentes))
    return ies_encontrados

def main():
    root = Tk()
    root.withdraw()

    nome_pesquisado = solicitar_entrada("Pesquisa", "Insira o nome a ser pesquisado")
    arquivo_zip = askopenfilename(filetypes=[("Arquivos ZIP", "*.zip")])
    if not arquivo_zip:
        messagebox.showerror("Erro", "Processo interrompido pelo usuário.")
        exit()

    diretorio_saida = askdirectory(title="Selecione o diretório de saída")
    if not diretorio_saida:
        messagebox.showerror("Erro", "Processo interrompido pelo usuário.")
        exit()

    with zipfile.ZipFile(arquivo_zip, 'r') as arquivo_zip:
        correspondentes = {}
        for nome_arquivo in arquivo_zip.namelist():
            if nome_arquivo.endswith('.xml'):
                arquivo_xml = arquivo_zip.open(nome_arquivo)
                conteudo_xml = arquivo_xml.read()
                raiz = ET.fromstring(conteudo_xml)

                tag_encontrada = buscar_tags_toma(raiz)

                if tag_encontrada in ["toma0", "toma1", "toma2", "toma3"]:
                    diretorio_cliente_tomador = os.path.join(diretorio_saida, "Cliente - Tomador")
                    salvar_arquivo_em_diretorio(conteudo_xml, diretorio_cliente_tomador, nome_arquivo)
                elif tag_encontrada == "toma4":
                    diretorio_cliente_outros = os.path.join(diretorio_saida, "Cliente - Outros")
                    salvar_arquivo_em_diretorio(conteudo_xml, diretorio_cliente_outros, nome_arquivo)

                grupo = buscar_nome(raiz, nome_pesquisado)
                if grupo is not None:
                    IE = buscar_ie(grupo)
                    if IE is not None:
                        correspondentes[IE] = True

                ies_encontrados = buscar_ies_no_elemento(raiz, correspondentes)

                if ies_encontrados:
                    for IE in ies_encontrados:
                        diretorio_saida_ie = os.path.join(diretorio_saida, IE)
                        salvar_arquivo_em_diretorio(conteudo_xml, diretorio_saida_ie, nome_arquivo)
                elif grupo is None:
                    diretorio_sem_identificacao = os.path.join(diretorio_saida, "Sem identificação")
                    salvar_arquivo_em_diretorio(conteudo_xml, diretorio_sem_identificacao, nome_arquivo)
                else:
                    diretorio_sem_ie = os.path.join(diretorio_saida, "NF sem IE")
                    salvar_arquivo_em_diretorio(conteudo_xml, diretorio_sem_ie, nome_arquivo)

    print("Arquivos XML extraídos com sucesso para o diretório:", diretorio_saida)

    for nome_arquivo in glob.glob(diretorio_saida + "/**/*.xml", recursive=True):
        novo_nome_arquivo = renomear_arquivo_xml(os.path.basename(nome_arquivo))
        novo_caminho_arquivo = os.path.join(os.path.dirname(nome_arquivo), novo_nome_arquivo)
        try:
            os.rename(nome_arquivo, novo_caminho_arquivo)
        except Exception as e:
            print(f"Não foi possível renomear o arquivo {nome_arquivo} para {novo_nome_arquivo}. Erro: {str(e)}")

    messagebox.showinfo("Processo Concluído", "O processo da automação foi concluído com sucesso!")

if __name__ == '__main__':
    main()
