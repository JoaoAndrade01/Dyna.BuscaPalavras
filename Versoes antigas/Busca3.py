import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit, QTextEdit, QFileDialog, QCheckBox)
from docx import Document

# Função para buscar palavra em um arquivo .docx
def buscar_palavra_em_docx(arquivo, palavra):
    doc = Document(arquivo)
    for paragrafo in doc.paragraphs:
        if palavra.lower() in paragrafo.text.lower():
            return True
    return False

# Função para realizar a busca em arquivos .docx na pasta
def buscar_palavras_na_pasta(pasta, palavra):
    arquivos_encontrados = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".docx"):
            caminho_completo = os.path.join(pasta, arquivo)
            if buscar_palavra_em_docx(caminho_completo, palavra):
                arquivos_encontrados.append(arquivo)
    return arquivos_encontrados

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        
        self.setWindowTitle("Busca de Palavras em Arquivos .docx")
        
        self.layout = QVBoxLayout()
        
        # Caixa de seleção para o caminho sugerido
        self.checkbox_caminho_sugerido = QCheckBox("Usar caminho sugerido")
        self.checkbox_caminho_sugerido.setChecked(True)
        self.layout.addWidget(self.checkbox_caminho_sugerido)
        
        self.caminho_sugerido = r"G:\Drives compartilhados\Dynatest RIO\02_Operações\02. Programa CECILIA\Projeto CYBORG\BD\Atestados"
        
        # Campo para o caminho da pasta
        self.caminho_pasta_label = QLabel("Caminho da pasta:")
        self.layout.addWidget(self.caminho_pasta_label)
        
        self.caminho_pasta_input = QLineEdit(self)
        self.caminho_pasta_input.setText(self.caminho_sugerido)
        self.layout.addWidget(self.caminho_pasta_input)
        
        # Botão para selecionar pasta
        self.selecionar_pasta_btn = QPushButton("Selecionar Pasta")
        self.selecionar_pasta_btn.clicked.connect(self.selecionar_pasta)
        self.layout.addWidget(self.selecionar_pasta_btn)
        
        # Campo para a palavra a ser buscada
        self.palavra_label = QLabel("Digite a palavra a ser buscada:")
        self.layout.addWidget(self.palavra_label)
        
        self.palavra_input = QLineEdit(self)
        self.layout.addWidget(self.palavra_input)
        
        # Botão para buscar palavra
        self.buscar_btn = QPushButton("Buscar")
        self.buscar_btn.clicked.connect(self.buscar_palavra)
        self.layout.addWidget(self.buscar_btn)
        
        # Botão para nova pesquisa
        self.nova_pesquisa_btn = QPushButton("Nova Pesquisa")
        self.nova_pesquisa_btn.clicked.connect(self.nova_pesquisa)
        self.layout.addWidget(self.nova_pesquisa_btn)
        
        # Campo para exibir os resultados
        self.resultados_text = QTextEdit(self)
        self.resultados_text.setReadOnly(True)
        self.layout.addWidget(self.resultados_text)
        
        # Campo para a segunda palavra
        self.palavra_input_2 = QLineEdit(self)
        self.palavra_input_2.setPlaceholderText("Digite uma nova palavra para buscar nos resultados")
        self.layout.addWidget(self.palavra_input_2)
        
        # Botão para buscar segunda palavra
        self.buscar_novamente_btn = QPushButton("Buscar Novamente")
        self.buscar_novamente_btn.clicked.connect(self.buscar_novamente)
        self.layout.addWidget(self.buscar_novamente_btn)

        # Configuração da janela principal
        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

    # Função para selecionar uma pasta
    def selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecione a pasta")
        if pasta:
            self.caminho_pasta_input.setText(pasta)

    # Função para buscar a palavra
    def buscar_palavra(self):
        self.resultados_text.clear()
        caminho_pasta = self.caminho_pasta_input.text()
        palavra = self.palavra_input.text()

        if not os.path.exists(caminho_pasta):
            self.resultados_text.setText("Caminho da pasta não existe.")
            return

        arquivos_encontrados = buscar_palavras_na_pasta(caminho_pasta, palavra)
        
        if arquivos_encontrados:
            self.resultados_text.append(f"A palavra '{palavra}' foi encontrada nos seguintes arquivos:\n")
            for arquivo in arquivos_encontrados:
                self.resultados_text.append(arquivo)
        else:
            self.resultados_text.setText(f"Nenhum arquivo contém a palavra '{palavra}'.")

    # Função para nova pesquisa
    def nova_pesquisa(self):
        self.palavra_input.clear()
        self.palavra_input_2.clear()
        self.resultados_text.clear()

    # Função para buscar segunda palavra nos arquivos já encontrados
    def buscar_novamente(self):
        segunda_palavra = self.palavra_input_2.text()
        arquivos_encontrados = self.resultados_text.toPlainText().splitlines()[1:]  # Pular a primeira linha de cabeçalho
        
        if not segunda_palavra:
            self.resultados_text.append("Por favor, digite uma nova palavra para buscar.")
            return
        
        arquivos_filtrados = []
        for arquivo in arquivos_encontrados:
            arquivo = arquivo.strip()
            if buscar_palavra_em_docx(os.path.join(self.caminho_pasta_input.text(), arquivo), segunda_palavra):
                arquivos_filtrados.append(arquivo)
        
        if arquivos_filtrados:
            self.resultados_text.append(f"\nA palavra '{segunda_palavra}' foi encontrada nos seguintes arquivos:\n")
            for arquivo in arquivos_filtrados:
                self.resultados_text.append(arquivo)
        else:
            self.resultados_text.append(f"\nNenhum arquivo contém a palavra '{segunda_palavra}' nos resultados anteriores.")

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())
