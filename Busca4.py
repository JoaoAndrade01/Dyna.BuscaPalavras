import os
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget,
    QPushButton, QFileDialog, QLineEdit, QTextEdit, QMessageBox,
    QProgressBar, QTableWidget, QTableWidgetItem
)
from docx import Document

# Função para obter a lista de arquivos .docx em uma pasta ordenado pelo nome
def obter_arquivos_docx(pasta):
    arquivos = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".docx"):
            arquivos.append(arquivo)
    arquivos.sort()  # Ordena os arquivos pelo nome
    return arquivos

# Função para buscar uma palavra em um arquivo .docx
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

# Função para buscar uma palavra na lista de arquivos .docx
def buscar_palavra_nos_arquivos(pasta, arquivos, palavra, progress_callback=None):
    arquivos_encontrados = []
    total_arquivos = len(arquivos)
    
    for index, arquivo in enumerate(arquivos):
        if buscar_palavra_em_docx(os.path.join(pasta, arquivo), palavra):
            arquivos_encontrados.append(arquivo)
        
        # Atualiza a barra de progresso se o callback foi fornecido
        if progress_callback:
            progress_callback(index + 1, total_arquivos)

    return arquivos_encontrados

# Função para criar um dicionário de uma lista de arquivos .docx
def criar_dicionario_arquivos(arquivos):
    dicionario = {}
    for arquivo in arquivos:
        try:
            codigo_str, empresa, status = arquivo.split("-")
            codigo = codigo_str.strip()
            dicionario[arquivo] = {
                "Código": codigo,
                "Empresa": empresa.strip(),
                "Status": status.strip().replace(".docx", "")
            }
        except ValueError:
            # Caso o nome do arquivo não tenha o formato esperado ou a conversão falhe, exibe uma mensagem de erro
            print(f"Formato inesperado ou erro ao converter o código no arquivo: {arquivo}")
    return dicionario


# Classe principal da interface gráfica
class BuscaPalavrasApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Busca de Palavras em .docx")
        self.setGeometry(300, 300, 600, 500)

        layout = QVBoxLayout()

        self.btn_selecionar_pasta = QPushButton("Selecionar Pasta", self)
        self.btn_selecionar_pasta.clicked.connect(self.selecionar_pasta)
        layout.addWidget(self.btn_selecionar_pasta)

        self.input_palavra1 = QLineEdit(self)
        self.input_palavra1.setPlaceholderText("Digite a primeira palavra para busca")
        layout.addWidget(self.input_palavra1)

        self.btn_buscar1 = QPushButton("Buscar Primeira Palavra", self)
        self.btn_buscar1.clicked.connect(self.iniciar_primeira_busca)
        layout.addWidget(self.btn_buscar1)

        self.input_palavra2 = QLineEdit(self)
        self.input_palavra2.setPlaceholderText("Digite a segunda palavra para busca (opcional)")
        layout.addWidget(self.input_palavra2)

        self.btn_buscar2 = QPushButton("Buscar Segunda Palavra", self)
        self.btn_buscar2.clicked.connect(self.iniciar_segunda_busca)
        self.btn_buscar2.setEnabled(False)
        layout.addWidget(self.btn_buscar2)

        self.resultado_busca = QTextEdit(self)
        self.resultado_busca.setReadOnly(True)
        layout.addWidget(self.resultado_busca)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Adicionando o widget para a tabela
        self.tabela_resultados = QTableWidget(self)
        self.tabela_resultados.setColumnCount(3)
        self.tabela_resultados.setHorizontalHeaderLabels(["Código", "Empresa", "Status"])
        layout.addWidget(self.tabela_resultados)

        self.btn_nova_pesquisa = QPushButton("Nova Pesquisa", self)
        self.btn_nova_pesquisa.clicked.connect(self.limpar_pesquisa)
        layout.addWidget(self.btn_nova_pesquisa)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.pasta = ""
        self.arquivos_primeira_busca = []
        self.arquivos_segunda_busca = []
        self.colunas_adicionais = 0  # Contador para as colunas adicionais de buscas subsequentes

    def selecionar_pasta(self):
        self.pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if self.pasta:
            self.resultado_busca.setText(f"Pasta selecionada: {self.pasta}\n")
            self.arquivos_primeira_busca = []
            self.arquivos_segunda_busca = []
            self.btn_buscar2.setEnabled(False)
            self.input_palavra2.clear()

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

        arquivos = obter_arquivos_docx(self.pasta)
        self.progress_bar.setMaximum(len(arquivos))
        self.arquivos_primeira_busca = buscar_palavra_nos_arquivos(
            self.pasta, arquivos, palavra1, self.atualizar_progresso
        )

        dicionario_arquivos = criar_dicionario_arquivos(self.arquivos_primeira_busca)

        QApplication.restoreOverrideCursor()

        if self.arquivos_primeira_busca:
            self.resultado_busca.append(f"A palavra '{palavra1}' foi encontrada nos seguintes arquivos:")

            self.tabela_resultados.setRowCount(len(dicionario_arquivos))
            for row, dados in enumerate(dicionario_arquivos.values()):
                # Preenche as colunas Código, Empresa e Status
                self.tabela_resultados.setItem(row, 0, QTableWidgetItem(dados['Código']))
                self.tabela_resultados.setItem(row, 1, QTableWidgetItem(dados['Empresa']))
                self.tabela_resultados.setItem(row, 2, QTableWidgetItem(dados['Status']))

                # Torna as células não editáveis
                for col in range(3):
                    item = self.tabela_resultados.item(row, col)
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            self.resultado_busca.append("")  
            self.btn_buscar2.setEnabled(True)
        else:
            self.resultado_busca.append(f"Nenhum arquivo contém a palavra '{palavra1}'.")
            self.btn_buscar2.setEnabled(False)

        self.progress_bar.setValue(0)

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

        # Adicionar nova coluna para a palavra buscada
        self.colunas_adicionais += 1
        nova_coluna_index = self.tabela_resultados.columnCount()
        self.tabela_resultados.setColumnCount(nova_coluna_index + 1)
        self.tabela_resultados.setHorizontalHeaderItem(nova_coluna_index, QTableWidgetItem(palavra2))

        # Realiza a segunda busca
        self.progress_bar.setMaximum(len(self.arquivos_primeira_busca))
        for row, arquivo in enumerate(self.arquivos_primeira_busca):
            encontrado = buscar_palavra_em_docx(os.path.join(self.pasta, arquivo), palavra2)
            resultado = "Sim" if encontrado else "Não"
            item = QTableWidgetItem(resultado)

            # Torna a nova célula da segunda busca não editável
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.tabela_resultados.setItem(row, nova_coluna_index, item)

        QApplication.restoreOverrideCursor()
        self.resultado_busca.append(f"Busca pela palavra '{palavra2}' concluída.\n")
        self.progress_bar.setValue(0)

        # Ordenar a tabela com base na nova coluna
        self.ordenar_tabela_por_coluna(nova_coluna_index)

    def ordenar_tabela_por_coluna(self, coluna):
        # Ativa a ordenação na tabela e ordena pela coluna fornecida (onde foi feita a segunda busca)
        self.tabela_resultados.sortItems(coluna, Qt.DescendingOrder)


    def limpar_pesquisa(self):
        self.arquivos_primeira_busca = []
        self.arquivos_segunda_busca = []
        self.input_palavra1.clear()
        self.input_palavra2.clear()
        self.resultado_busca.clear()
        self.tabela_resultados.clearContents()
        self.tabela_resultados.setColumnCount(3)
        self.tabela_resultados.setHorizontalHeaderLabels(["Código", "Empresa", "Status"])
        self.colunas_adicionais = 0  # Reiniciar contador de colunas adicionais
        self.btn_buscar2.setEnabled(False)
        self.resultado_busca.append("Pronto para uma nova pesquisa!\n")
        self.progress_bar.setValue(0)

    def atualizar_progresso(self, valor, total):
        self.progress_bar.setValue(valor)

# Execução do aplicativo
if __name__ == "__main__":
    from PySide6.QtCore import Qt
    app = QApplication([])
    janela = BuscaPalavrasApp()
    janela.show()
    app.exec()
