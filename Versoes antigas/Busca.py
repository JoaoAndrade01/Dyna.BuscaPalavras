import os
from docx import Document

# Função para buscar palavra em um arquivo .docx
def buscar_palavra_em_docx(arquivo, palavra):
    doc = Document(arquivo)
    for paragrafo in doc.paragraphs:
        if palavra.lower() in paragrafo.text.lower():
            return True
    return False

# Função principal para buscar palavras na pasta
def buscar_palavras_na_pasta(pasta, palavra):
    arquivos_encontrados = []
    
    # Itera sobre os arquivos da pasta
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".docx"):  # Verifica se o arquivo é .docx
            caminho_completo = os.path.join(pasta, arquivo)
            if buscar_palavra_em_docx(caminho_completo, palavra):
                arquivos_encontrados.append(arquivo)
    
    return arquivos_encontrados

# Função para realizar a segunda busca dentro dos arquivos já filtrados
def nova_busca_nos_arquivos(pasta, arquivos, nova_palavra):
    arquivos_encontrados_novamente = []
    
    # Itera sobre os arquivos que passaram pela primeira busca
    for arquivo in arquivos:
        caminho_completo = os.path.join(pasta, arquivo)
        if buscar_palavra_em_docx(caminho_completo, nova_palavra):
            arquivos_encontrados_novamente.append(arquivo)
    
    return arquivos_encontrados_novamente

# Referenciar a pasta
pasta = r"G:\Drives compartilhados\Dynatest RIO\02_Operações\02. Programa CECILIA\Projeto CYBORG\BD\Atestados"

# Pedir a primeira palavra para busca
palavra = input("Digite a primeira palavra a ser buscada: ")

# Realizar a primeira busca
arquivos_com_palavra = buscar_palavras_na_pasta(pasta, palavra)

# Exibir os arquivos encontrados na primeira busca
if arquivos_com_palavra:
    print("A palavra '{}' foi encontrada nos seguintes arquivos:".format(palavra))
    for arquivo in arquivos_com_palavra:
        print(arquivo)
    
    # Perguntar ao usuário se deseja realizar uma segunda busca
    nova_busca = input("Deseja realizar uma nova busca nos arquivos encontrados? (s/n): ").lower()
    
    if nova_busca == 's':
        nova_palavra = input("Digite a nova palavra a ser buscada: ")
        
        # Realizar a segunda busca apenas nos arquivos já filtrados
        arquivos_com_nova_palavra = nova_busca_nos_arquivos(pasta, arquivos_com_palavra, nova_palavra)
        
        # Exibir os arquivos encontrados na segunda busca
        if arquivos_com_nova_palavra:
            print("A nova palavra '{}' foi encontrada nos seguintes arquivos:".format(nova_palavra))
            for arquivo in arquivos_com_nova_palavra:
                print(arquivo)
        else:
            print("Nenhum dos arquivos contém a palavra '{}'.".format(nova_palavra))
    else:
        print("Busca encerrada.")
else:
    print("Nenhum arquivo contém a palavra '{}'.".format(palavra))
