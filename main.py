# main.py
import sys
import sqlite3
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
import PyPDF2
import docx
import os
from datetime import datetime
import re


class Database:
    def __init__(self):
        self.conn = sqlite3.connect('produtos.db')
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()
        
        # Tabela de produtos
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                descricao TEXT,
                palavras_positivas TEXT,
                palavras_negativas TEXT,
                data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Tabela para histórico de análises
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS analises (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                arquivo_nome TEXT,
                data_analise DATETIME DEFAULT CURRENT_TIMESTAMP,
                resultado TEXT
            )
        ''')
        
        self.conn.commit()

    def get_produtos(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM produtos ORDER BY nome")
        return cursor.fetchall()

    def get_produto(self, produto_id):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM produtos WHERE id = ?", (produto_id,))
        return cursor.fetchone()

    def add_produto(self, nome, descricao, palavras_positivas, palavras_negativas):
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO produtos (nome, descricao, palavras_positivas, palavras_negativas)
            VALUES (?, ?, ?, ?)
        ''', (nome, descricao, palavras_positivas, palavras_negativas))
        self.conn.commit()
        return cursor.lastrowid

    def update_produto(self, produto_id, nome, descricao, palavras_positivas, palavras_negativas):
        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE produtos 
            SET nome = ?, descricao = ?, palavras_positivas = ?, palavras_negativas = ?
            WHERE id = ?
        ''', (nome, descricao, palavras_positivas, palavras_negativas, produto_id))
        self.conn.commit()

    def delete_produto(self, produto_id):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM produtos WHERE id = ?", (produto_id,))
        self.conn.commit()

    def salvar_analise(self, arquivo_nome, resultado):
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO analises (arquivo_nome, resultado)
            VALUES (?, ?)
        ''', (arquivo_nome, str(resultado)))
        self.conn.commit()


class DocumentAnalyzer:
    @staticmethod
    def extract_text_from_file(filepath):
        """Extrai texto de diferentes tipos de arquivos"""
        text = ""
        ext = os.path.splitext(filepath)[1].lower()
        
        try:
            if ext == '.pdf':
                with open(filepath, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for page in pdf_reader.pages:
                        text += page.extract_text() + "\n"
            
            elif ext in ['.doc', '.docx']:
                doc = docx.Document(filepath)
                for paragraph in doc.paragraphs:
                    text += paragraph.text + "\n"
            
            elif ext in ['.txt', '.rtf']:
                with open(filepath, 'r', encoding='utf-8') as file:
                    text = file.read()
            
            else:
                # Tentar ler como texto plano
                try:
                    with open(filepath, 'r', encoding='utf-8') as file:
                        text = file.read()
                except:
                    text = ""
        
        except Exception as e:
            print(f"Erro ao ler arquivo: {e}")
            return ""
        
        return text

    @staticmethod
    def normalize_text(text):
        """Normaliza o texto para análise"""
        # Converte para minúsculas
        text = text.lower()
        
        # Remove caracteres especiais, mantendo letras, números, espaços e acentos
        # Esta regex mantém letras (incluindo acentuadas), números e espaços
        text = re.sub(r'[^\w\sáàâãéèêíìîóòôõúùûçÁÀÂÃÉÈÊÍÌÎÓÒÔÕÚÙÛÇ]', ' ', text)
        
        # Remove múltiplos espaços
        text = re.sub(r'\s+', ' ', text)
        
        return text.strip()

    @staticmethod
    def count_occurrences(text, word_list):
        """Conta ocorrências de forma inteligente usando expressões regulares"""
        total_count = 0
        
        for word in word_list:
            word = word.strip()
            if not word:
                continue
                
            # Se a palavra tem espaço, cria padrão flexível
            if ' ' in word:
                # Remove espaços extras
                word = word.strip()
                
                # Cria padrão regex que aceita:
                # 1. A palavra exata com espaços
                # 2. A palavra sem espaços
                # 3. Variações com diferentes quantidades de espaços
                pattern = r'\b' + re.escape(word) + r'\b'
                pattern_alt = r'\b' + re.escape(word.replace(' ', '')) + r'\b'
                pattern_flex = r'\b' + re.escape(word.replace(' ', r'\s+')) + r'\b'
                
                # Conta ocorrências para cada padrão
                count = len(re.findall(pattern, text))
                count_alt = len(re.findall(pattern_alt, text))
                count_flex = len(re.findall(pattern_flex, text))
                
                # Usa o maior valor
                total_count += max(count, count_alt, count_flex)
                
            else:
                # Para palavras simples, conta normalmente com regex para palavras inteiras
                pattern = r'\b' + re.escape(word) + r'\b'
                total_count += len(re.findall(pattern, text))
        
        return total_count

    @staticmethod
    def calculate_index(produto, text, palavras_positivas, palavras_negativas):
        """Calcula índice baseado nas palavras positivas e negativas"""
        text_normalized = DocumentAnalyzer.normalize_text(text)
        
        # Converte strings de palavras para listas
        pos_list = [p.strip().lower() for p in palavras_positivas.split(',') if p.strip()]
        neg_list = [p.strip().lower() for p in palavras_negativas.split(',') if p.strip()]
        
        # Conta ocorrências usando o método inteligente
        pos_count = DocumentAnalyzer.count_occurrences(text_normalized, pos_list)
        neg_count = DocumentAnalyzer.count_occurrences(text_normalized, neg_list)
        
        # Calcula índice
        total_ocorrencias = pos_count + neg_count
        if total_ocorrencias == 0:
            return 0, 0, 0
        
        # Fórmula: (positivas - negativas) / total * 100
        index = ((pos_count - neg_count) / total_ocorrencias) * 100
        
        # Garante que o índice não seja negativo
        if index < 0:
            index = 0
        
        return round(index, 2), pos_count, neg_count

    @staticmethod
    def find_products_in_text(text, produtos):
        """Encontra produtos no texto e calcula seus índices"""
        resultados = []
        text_normalized = DocumentAnalyzer.normalize_text(text)
        
        for produto in produtos:
            produto_id, nome, descricao, palavras_positivas, palavras_negativas, _ = produto
            
            # Normaliza o nome do produto para busca
            nome_normalized = DocumentAnalyzer.normalize_text(nome)
            
            # Verifica se o nome do produto aparece no texto
            # Usando regex para busca de palavras inteiras
            pattern = r'\b' + re.escape(nome_normalized) + r'\b'
            
            if re.search(pattern, text_normalized):
                index, pos_count, neg_count = DocumentAnalyzer.calculate_index(
                    nome, text, palavras_positivas, palavras_negativas
                )
                
                # Adiciona informações detalhadas sobre as palavras encontradas
                palavras_encontradas = DocumentAnalyzer.get_palavras_encontradas(
                    text_normalized, palavras_positivas, palavras_negativas
                )
                
                resultados.append({
                    'id': produto_id,
                    'nome': nome,
                    'descricao': descricao,
                    'indice': index,
                    'positivas_encontradas': pos_count,
                    'negativas_encontradas': neg_count,
                    'palavras_positivas': palavras_positivas,
                    'palavras_negativas': palavras_negativas,
                    'palavras_encontradas_lista': palavras_encontradas
                })
        
        # Ordena por índice (maior primeiro)
        resultados.sort(key=lambda x: x['indice'], reverse=True)
        return resultados

    @staticmethod
    def get_palavras_encontradas(text, palavras_positivas, palavras_negativas):
        """Retorna lista das palavras específicas encontradas no texto"""
        encontradas = []
        
        # Processa palavras positivas
        pos_list = [p.strip().lower() for p in palavras_positivas.split(',') if p.strip()]
        for palavra in pos_list:
            palavra = palavra.strip()
            if ' ' in palavra:
                # Para frases, verifica variações
                pattern = r'\b' + re.escape(palavra) + r'\b'
                pattern_alt = r'\b' + re.escape(palavra.replace(' ', '')) + r'\b'
                pattern_flex = r'\b' + re.escape(palavra.replace(' ', r'\s+')) + r'\b'
                
                if (re.search(pattern, text) or 
                    re.search(pattern_alt, text) or 
                    re.search(pattern_flex, text)):
                    encontradas.append(f"✓ {palavra}")
            else:
                pattern = r'\b' + re.escape(palavra) + r'\b'
                if re.search(pattern, text):
                    encontradas.append(f"✓ {palavra}")
        
        # Processa palavras negativas
        neg_list = [p.strip().lower() for p in palavras_negativas.split(',') if p.strip()]
        for palavra in neg_list:
            palavra = palavra.strip()
            if ' ' in palavra:
                # Para frases, verifica variações
                pattern = r'\b' + re.escape(palavra) + r'\b'
                pattern_alt = r'\b' + re.escape(palavra.replace(' ', '')) + r'\b'
                pattern_flex = r'\b' + re.escape(palavra.replace(' ', r'\s+')) + r'\b'
                
                if (re.search(pattern, text) or 
                    re.search(pattern_alt, text) or 
                    re.search(pattern_flex, text)):
                    encontradas.append(f"✗ {palavra}")
            else:
                pattern = r'\b' + re.escape(palavra) + r'\b'
                if re.search(pattern, text):
                    encontradas.append(f"✗ {palavra}")
        
        return encontradas

    @staticmethod
    def extract_product_context(text, product_name, context_words=10):
        """Extrai o contexto onde o produto é mencionado"""
        text_normalized = DocumentAnalyzer.normalize_text(text)
        product_normalized = DocumentAnalyzer.normalize_text(product_name)
        
        # Encontra todas as ocorrências do produto
        pattern = r'\b' + re.escape(product_normalized) + r'\b'
        matches = list(re.finditer(pattern, text_normalized))
        
        contexts = []
        for match in matches:
            start = max(0, match.start() - (context_words * 10))
            end = min(len(text_normalized), match.end() + (context_words * 10))
            
            context = text_normalized[start:end]
            # Adiciona reticências se cortou texto
            if start > 0:
                context = "..." + context
            if end < len(text_normalized):
                context = context + "..."
            
            contexts.append(context)
        
        return contexts[:3]  # Retorna até 3 contextos

class ProdutoDialog(QDialog):
    def __init__(self, parent=None, produto=None):
        super().__init__(parent)
        self.produto = produto
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Editar Produto" if self.produto else "Novo Produto")
        self.setModal(True)
        self.resize(500, 400)

        layout = QVBoxLayout(self)

        # Nome
        layout.addWidget(QLabel("Nome do Produto:"))
        self.nome_edit = QLineEdit()
        if self.produto:
            self.nome_edit.setText(self.produto[1])
        layout.addWidget(self.nome_edit)

        # Descrição
        layout.addWidget(QLabel("Descrição:"))
        self.descricao_edit = QTextEdit()
        if self.produto:
            self.descricao_edit.setText(self.produto[2])
        layout.addWidget(self.descricao_edit)

        # Palavras Positivas
        layout.addWidget(QLabel("Palavras Positivas (separadas por vírgula):"))
        self.pos_edit = QTextEdit()
        if self.produto:
            self.pos_edit.setText(self.produto[3])
        layout.addWidget(self.pos_edit)

        # Palavras Negativas
        layout.addWidget(QLabel("Palavras Negativas (separadas por vírgula):"))
        self.neg_edit = QTextEdit()
        if self.produto:
            self.neg_edit.setText(self.produto[4])
        layout.addWidget(self.neg_edit)

        # Botões
        button_layout = QHBoxLayout()
        self.salvar_btn = QPushButton("Salvar")
        self.cancelar_btn = QPushButton("Cancelar")
        
        self.salvar_btn.clicked.connect(self.accept)
        self.cancelar_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(self.salvar_btn)
        button_layout.addWidget(self.cancelar_btn)
        layout.addLayout(button_layout)

    def get_data(self):
        return {
            'nome': self.nome_edit.text().strip(),
            'descricao': self.descricao_edit.toPlainText().strip(),
            'palavras_positivas': self.pos_edit.toPlainText().strip(),
            'palavras_negativas': self.neg_edit.toPlainText().strip()
        }


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.analyzer = DocumentAnalyzer()
        self.current_file = None
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Analisador de Licitações")
        self.setGeometry(100, 100, 1200, 700)

        # Widget central e layout principal
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Painel esquerdo - Cadastro de Produtos
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        left_layout.addWidget(QLabel("<h3>Cadastro de Produtos</h3>"))
        
        # Botões do CRUD
        button_layout = QHBoxLayout()
        self.novo_btn = QPushButton("Novo Produto")
        self.editar_btn = QPushButton("Editar")
        self.excluir_btn = QPushButton("Excluir")
        self.atualizar_btn = QPushButton("Atualizar Lista")
        
        self.novo_btn.clicked.connect(self.novo_produto)
        self.editar_btn.clicked.connect(self.editar_produto)
        self.excluir_btn.clicked.connect(self.excluir_produto)
        self.atualizar_btn.clicked.connect(self.carregar_produtos)
        
        button_layout.addWidget(self.novo_btn)
        button_layout.addWidget(self.editar_btn)
        button_layout.addWidget(self.excluir_btn)
        button_layout.addWidget(self.atualizar_btn)
        left_layout.addLayout(button_layout)
        
        # Lista de produtos
        self.produtos_table = QTableWidget()
        self.produtos_table.setColumnCount(3)
        self.produtos_table.setHorizontalHeaderLabels(['ID', 'Nome', 'Palavras Positivas'])
        self.produtos_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.produtos_table.setEditTriggers(QTableWidget.NoEditTriggers)
        left_layout.addWidget(self.produtos_table)
        
        main_layout.addWidget(left_panel, 1)

        # Painel direito - Análise de Documentos
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        right_layout.addWidget(QLabel("<h3>Análise de Documentos</h3>"))
        
        # Seleção de arquivo
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Selecione um arquivo...")
        self.browse_btn = QPushButton("Procurar")
        self.analisar_btn = QPushButton("Analisar Documento")
        
        self.browse_btn.clicked.connect(self.selecionar_arquivo)
        self.analisar_btn.clicked.connect(self.analisar_documento)
        
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_btn)
        file_layout.addWidget(self.analisar_btn)
        right_layout.addLayout(file_layout)
        
        # Resultados da análise
        right_layout.addWidget(QLabel("<h4>Resultados da Análise:</h4>"))
        self.resultados_table = QTableWidget()
        self.resultados_table.setColumnCount(5)
        self.resultados_table.setHorizontalHeaderLabels([
            'Produto', 'Índice', 'Positivas', 'Negativas', 'Status'
        ])
        self.resultados_table.setEditTriggers(QTableWidget.NoEditTriggers)
        right_layout.addWidget(self.resultados_table)
        
        # Detalhes do resultado selecionado
        right_layout.addWidget(QLabel("<h4>Detalhes:</h4>"))
        self.detalhes_text = QTextEdit()
        self.detalhes_text.setReadOnly(True)
        right_layout.addWidget(self.detalhes_text)
        
        main_layout.addWidget(right_panel, 2)
        
        # Conectar sinais
        self.resultados_table.itemSelectionChanged.connect(self.mostrar_detalhes)
        
        # Carregar dados iniciais
        self.carregar_produtos()

    def carregar_produtos(self):
        produtos = self.db.get_produtos()
        self.produtos_table.setRowCount(len(produtos))
        
        for row, produto in enumerate(produtos):
            self.produtos_table.setItem(row, 0, QTableWidgetItem(str(produto[0])))
            self.produtos_table.setItem(row, 1, QTableWidgetItem(produto[1]))
            self.produtos_table.setItem(row, 2, QTableWidgetItem(produto[3][:50] + "..." if len(produto[3]) > 50 else produto[3]))
        
        self.produtos_table.resizeColumnsToContents()

    def novo_produto(self):
        dialog = ProdutoDialog(self)
        if dialog.exec():
            data = dialog.get_data()
            if data['nome']:
                self.db.add_produto(
                    data['nome'],
                    data['descricao'],
                    data['palavras_positivas'],
                    data['palavras_negativas']
                )
                self.carregar_produtos()

    def editar_produto(self):
        selected = self.produtos_table.currentRow()
        if selected >= 0:
            produto_id = int(self.produtos_table.item(selected, 0).text())
            produto = self.db.get_produto(produto_id)
            
            dialog = ProdutoDialog(self, produto)
            if dialog.exec():
                data = dialog.get_data()
                if data['nome']:
                    self.db.update_produto(
                        produto_id,
                        data['nome'],
                        data['descricao'],
                        data['palavras_positivas'],
                        data['palavras_negativas']
                    )
                    self.carregar_produtos()

    def excluir_produto(self):
        selected = self.produtos_table.currentRow()
        if selected >= 0:
            produto_id = int(self.produtos_table.item(selected, 0).text())
            
            reply = QMessageBox.question(
                self, 'Confirmar Exclusão',
                f'Tem certeza que deseja excluir este produto?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.db.delete_produto(produto_id)
                self.carregar_produtos()

    def selecionar_arquivo(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Documento",
            "",
            "Documentos (*.pdf *.doc *.docx *.txt *.rtf);;Todos os arquivos (*)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            self.current_file = file_path

    def analisar_documento(self):
        if not self.current_file or not os.path.exists(self.current_file):
            QMessageBox.warning(self, "Aviso", "Selecione um arquivo válido!")
            return
        
        try:
            # Alterar cursor para ampulheta (wait cursor)
            QApplication.setOverrideCursor(Qt.WaitCursor)

            # Desabilitar botão enquanto processa
            self.analisar_btn.setEnabled(False)
            self.analisar_btn.setText("Analisando...")
            
            # Forçar atualização da interface
            QApplication.processEvents()

            # Extrair texto do documento
            texto = self.analyzer.extract_text_from_file(self.current_file)
            
            if not texto.strip():
                QMessageBox.warning(self, "Aviso", "Não foi possível extrair texto do documento!")
                return
            
            # Obter produtos do banco de dados
            produtos = self.db.get_produtos()
            
            if not produtos:
                QMessageBox.warning(self, "Aviso", "Cadastre produtos primeiro!")
                return
            
            # Analisar documento
            resultados = self.analyzer.find_products_in_text(texto, produtos)
            
            # Exibir resultados
            self.exibir_resultados(resultados)
            
            # Salvar análise no histórico
            self.db.salvar_analise(os.path.basename(self.current_file), resultados)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao analisar documento: {str(e)}")
        finally:
            # Restaurar cursor normal
            QApplication.restoreOverrideCursor()
        
            # Reabilitar botão
            self.analisar_btn.setEnabled(True)
            self.analisar_btn.setText("Analisar Documento")

    def exibir_resultados(self, resultados):
        self.resultados_table.setRowCount(len(resultados))
        
        for row, resultado in enumerate(resultados):
            # Produto
            self.resultados_table.setItem(row, 0, QTableWidgetItem(resultado['nome']))
            
            # Índice
            index_item = QTableWidgetItem(str(resultado['indice']))
            # Colorir baseado no índice
            if resultado['indice'] >= 70:
                index_item.setBackground(QColor(144, 238, 144))  # Verde claro
            elif resultado['indice'] >= 30:
                index_item.setBackground(QColor(255, 255, 224))  # Amarelo claro
            else:
                index_item.setBackground(QColor(255, 182, 193))  # Vermelho claro
            self.resultados_table.setItem(row, 1, index_item)
            
            # Contagem positiva
            self.resultados_table.setItem(row, 2, QTableWidgetItem(str(resultado['positivas_encontradas'])))
            
            # Contagem negativa
            self.resultados_table.setItem(row, 3, QTableWidgetItem(str(resultado['negativas_encontradas'])))
            
            # Status
            status = "Ótimo" if resultado['indice'] >= 70 else \
                     "Regular" if resultado['indice'] >= 30 else "Ruim"
            self.resultados_table.setItem(row, 4, QTableWidgetItem(status))
        
        self.resultados_table.resizeColumnsToContents()

    def mostrar_detalhes(self):
        selected = self.resultados_table.currentRow()
        if selected >= 0:
            detalhes = f"""
            <b>Produto:</b> {self.resultados_table.item(selected, 0).text()}<br>
            <b>Índice:</b> {self.resultados_table.item(selected, 1).text()}<br>
            <b>Palavras Positivas Encontradas:</b> {self.resultados_table.item(selected, 2).text()}<br>
            <b>Palavras Negativas Encontradas:</b> {self.resultados_table.item(selected, 3).text()}<br>
            <b>Status:</b> {self.resultados_table.item(selected, 4).text()}<br>
            <hr>
            <i>Nota: Índice calculado com base na relação entre palavras positivas e negativas encontradas na descrição.</i>
            """
            self.detalhes_text.setHtml(detalhes)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Estilo básico
    app.setStyleSheet("""
        QMainWindow {
            background-color: #292828;
        }
        QPushButton {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 8px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #45a049;
        }
        QTableWidget {
            background-color: #363434;
            border: 1px solid #ddd;
            color: white;
        }
        QLineEdit, QTextEdit {
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 4px;
        }
        QLabel[text*="<h3>"],
        QLabel[text*="<h4>"] {
            color: #ebe8e8;
        }
    """)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()