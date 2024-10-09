import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLabel, QFileDialog, QLineEdit, QTextEdit, QMessageBox)
from docx import Document

# Função para buscar palavra em um arquivo .docx
def buscar_palavra_em_docx(arquivo, palavra):
    try:
        doc = Document(arquivo)
    except Exception as e:
        print(f"Erro ao abrir {arquivo}: {e}")
        return False
    for paragrafo in doc.paragraphs:
        if palavra.lower() in paragrafo.text.lower():
            return True
    return False

# Função para buscar palavras na lista de arquivos
def buscar_palavras_nos_arquivos(pasta, arquivos, palavra):
    arquivos_encontrados = []
    for arquivo in arquivos:
        caminho_completo = os.path.join(pasta, arquivo)
        if buscar_palavra_em_docx(caminho_completo, palavra):
            arquivos_encontrados.append(arquivo)
    return arquivos_encontrados

# Classe principal da interface gráfica
class BuscaPalavrasApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Busca de Palavras em .docx")
        self.setGeometry(300, 300, 600, 500)

        # Layout principal
        layout = QVBoxLayout()

        # Botão para selecionar pasta
        self.btn_selecionar_pasta = QPushButton("Selecionar Pasta", self)
        self.btn_selecionar_pasta.clicked.connect(self.selecionar_pasta)
        layout.addWidget(self.btn_selecionar_pasta)

        # Campo para inserir primeira palavra
        self.input_palavra1 = QLineEdit(self)
        self.input_palavra1.setPlaceholderText("Digite a primeira palavra para busca")
        layout.addWidget(self.input_palavra1)

        # Botão para iniciar a primeira busca
        self.btn_buscar1 = QPushButton("Buscar Primeira Palavra", self)
        self.btn_buscar1.clicked.connect(self.iniciar_primeira_busca)
        layout.addWidget(self.btn_buscar1)

        # Campo para inserir segunda palavra
        self.input_palavra2 = QLineEdit(self)
        self.input_palavra2.setPlaceholderText("Digite a segunda palavra para busca (opcional)")
        layout.addWidget(self.input_palavra2)

        # Botão para iniciar a segunda busca
        self.btn_buscar2 = QPushButton("Buscar Segunda Palavra", self)
        self.btn_buscar2.clicked.connect(self.iniciar_segunda_busca)
        self.btn_buscar2.setEnabled(False)  # Inicialmente desabilitado
        layout.addWidget(self.btn_buscar2)

        # Área de texto para exibir resultados
        self.resultado_busca = QTextEdit(self)
        self.resultado_busca.setReadOnly(True)
        layout.addWidget(self.resultado_busca)

        # Widget central
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Variáveis para armazenar o caminho da pasta e os resultados
        self.pasta = ""
        self.arquivos_primeira_busca = []
        self.arquivos_segunda_busca = []

    # Função para selecionar a pasta
    def selecionar_pasta(self):
        self.pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if self.pasta:
            self.resultado_busca.setText(f"Pasta selecionada: {self.pasta}\n")
            # Limpar resultados anteriores
            self.arquivos_primeira_busca = []
            self.arquivos_segunda_busca = []
            self.btn_buscar2.setEnabled(False)
            self.input_palavra2.clear()

    # Função para iniciar a primeira busca
    def iniciar_primeira_busca(self):
        if not self.pasta:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione uma pasta.")
            return
        palavra1 = self.input_palavra1.text().strip()
        if not palavra1:
            QMessageBox.warning(self, "Aviso", "Por favor, insira a primeira palavra para busca.")
            return

        self.resultado_busca.append(f"\nBuscando a palavra '{palavra1}'...\n")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()

        # Realiza a primeira busca
        self.arquivos_primeira_busca = buscar_palavras_nos_arquivos(self.pasta, os.listdir(self.pasta), palavra1)

        QApplication.restoreOverrideCursor()

        # Exibir os arquivos encontrados na primeira busca
        if self.arquivos_primeira_busca:
            self.resultado_busca.append(f"A palavra '{palavra1}' foi encontrada nos seguintes arquivos:")
            for arquivo in self.arquivos_primeira_busca:
                self.resultado_busca.append(f" - {arquivo}")
            self.btn_buscar2.setEnabled(True)  # Habilitar botão para segunda busca
        else:
            self.resultado_busca.append(f"Nenhum arquivo contém a palavra '{palavra1}'.")
            self.btn_buscar2.setEnabled(False)

    # Função para iniciar a segunda busca
    def iniciar_segunda_busca(self):
        if not self.arquivos_primeira_busca:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo encontrado na primeira busca.")
            return
        palavra2 = self.input_palavra2.text().strip()
        if not palavra2:
            QMessageBox.warning(self, "Aviso", "Por favor, insira a segunda palavra para busca.")
            return

        self.resultado_busca.append(f"\nBuscando a palavra '{palavra2}' nos arquivos filtrados...\n")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()

        # Realiza a segunda busca nos arquivos filtrados
        self.arquivos_segunda_busca = buscar_palavras_nos_arquivos(self.pasta, self.arquivos_primeira_busca, palavra2)

        QApplication.restoreOverrideCursor()

        # Exibir os arquivos encontrados na segunda busca
        if self.arquivos_segunda_busca:
            self.resultado_busca.append(f"A palavra '{palavra2}' foi encontrada nos seguintes arquivos:")
            for arquivo in self.arquivos_segunda_busca:
                self.resultado_busca.append(f" - {arquivo}")
        else:
            self.resultado_busca.append(f"Nenhum dos arquivos filtrados contém a palavra '{palavra2}'.")

# Execução do aplicativo
if __name__ == "__main__":
    from PySide6.QtCore import Qt  # Import necessário para o cursor de espera
    app = QApplication([])
    janela = BuscaPalavrasApp()
    janela.show()
    app.exec()
