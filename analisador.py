import sys
import os
import tempfile
import json
import decimal  # Adicionar importação para decimal
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog, 
                            QLabel, QVBoxLayout, QHBoxLayout, QWidget, QScrollArea,
                            QListWidget, QListWidgetItem, QMessageBox, QInputDialog,
                            QLineEdit, QTabWidget, QTableWidget, QTableWidgetItem,
                            QHeaderView)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt, QSize
import PyPDF2
import fitz  # PyMuPDF

class PDFAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_file = os.path.join(os.path.expanduser("~"), ".pdf_analyzer_config.json")
        self.load_config()  # Carregar configuração salva
        self.initUI()
        self.current_pdf_path = None
        self.page_images = []
        # Configurar log
        self.setup_logging()

    def setup_logging(self):
        """Configurar o sistema de log simples"""
        import logging
        self.logger = logging.getLogger("PDFAnalyzer")
        self.logger.setLevel(logging.INFO)
        
        if not self.logger.handlers:
            # Criar handler para console
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            formatter = logging.Formatter('%(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

    def initUI(self):
        self.setWindowTitle('Analisador de PDF')
        self.setGeometry(100, 100, 1200, 700)

        # Layout principal
        main_layout = QHBoxLayout()
        
        # Painel esquerdo para controles e informações básicas
        left_panel = QVBoxLayout()
        
        # Botão para configurar Poppler (opcional)
        self.poppler_btn = QPushButton('Configurar Poppler (Opcional)', self)
        if self.poppler_path:
            self.poppler_btn.setText(f'Poppler: {self.poppler_path}')
        self.poppler_btn.clicked.connect(self.set_poppler_path)
        left_panel.addWidget(self.poppler_btn)
        
        # Botão para upload de PDF
        self.upload_btn = QPushButton('Selecionar PDF', self)
        self.upload_btn.clicked.connect(self.upload_pdf)
        left_panel.addWidget(self.upload_btn)
        
        # Informações do PDF
        self.info_label = QLabel('Nenhum arquivo selecionado', self)
        self.info_label.setWordWrap(True)
        left_panel.addWidget(self.info_label)
        
        # Lista de páginas e formatos
        self.page_list = QListWidget(self)
        self.page_list.currentRowChanged.connect(self.on_page_selected)
        left_panel.addWidget(QLabel('Páginas e Formatos:'))
        left_panel.addWidget(self.page_list)
        
        # Alerta de formatos diferentes
        self.format_alert = QLabel('', self)
        self.format_alert.setStyleSheet("color: red;")
        left_panel.addWidget(self.format_alert)
        
        # Painel central para informações detalhadas
        center_panel = QVBoxLayout()
        
        # Criar TabWidget para organizar tipos de informação
        self.tabs = QTabWidget()
        
        # Tab para informações de Boxes
        self.box_tab = QWidget()
        self.box_layout = QVBoxLayout(self.box_tab)
        
        # Tabela para exibir informações de boxes
        self.box_table = QTableWidget()
        self.box_table.setColumnCount(5)  # Box Type, Width, Height, X, Y
        self.box_table.setHorizontalHeaderLabels(['Tipo de Box', 'Largura (mm)', 'Altura (mm)', 'X (mm)', 'Y (mm)'])
        self.box_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.box_layout.addWidget(self.box_table)
        
        self.tabs.addTab(self.box_tab, "Boxes de Página")
        
        # Tab para erros e avisos
        self.log_tab = QWidget()
        self.log_layout = QVBoxLayout(self.log_tab)
        self.log_text = QLabel("Nenhum erro ou aviso registrado.")
        self.log_text.setWordWrap(True)
        self.log_layout.addWidget(self.log_text)
        
        self.tabs.addTab(self.log_tab, "Erros e Avisos")
        
        # Adicionar TabWidget ao painel central
        center_panel.addWidget(self.tabs)
        
        # Painel direito para preview
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel('Visualização:'))
        
        # Área de scroll para preview
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        
        # Container para as imagens de preview
        self.preview_container = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_container)
        
        self.scroll_area.setWidget(self.preview_container)
        right_panel.addWidget(self.scroll_area)
        
        # Adicionar os painéis ao layout principal
        main_layout.addLayout(left_panel, 1)    # 1 parte para o painel esquerdo
        main_layout.addLayout(center_panel, 1)  # 1 parte para o painel central
        main_layout.addLayout(right_panel, 2)   # 2 partes para o painel direito (maior)

        # Widget central para conter o layout principal
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        # Variáveis para armazenar dados das páginas
        self.page_data = []
        self.log_messages = []

    def load_config(self):
        """Carrega a configuração salva do arquivo"""
        self.poppler_path = None
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.poppler_path = config.get('poppler_path')
        except Exception as e:
            print(f"Erro ao carregar configuração: {str(e)}")

    def save_config(self):
        """Salva a configuração atual em um arquivo"""
        try:
            config = {
                'poppler_path': self.poppler_path
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"Erro ao salvar configuração: {str(e)}")
            QMessageBox.warning(self, "Aviso", f"Não foi possível salvar a configuração: {str(e)}")
    
    def add_log_message(self, message, level="INFO"):
        """Adiciona uma mensagem ao log interno do aplicativo"""
        self.log_messages.append(f"{level}: {message}")
        log_text = "<br/>".join(self.log_messages)
        self.log_text.setText(f"<html><body>{log_text}</body></html>")
        
        # Também envia para o logger do sistema
        if hasattr(self, 'logger'):
            if level == "ERROR":
                self.logger.error(message)
            elif level == "WARNING":
                self.logger.warning(message)
            else:
                self.logger.info(message)
        
    def set_poppler_path(self):
        """Configura e salva o caminho do Poppler"""
        path, ok = QInputDialog.getText(
            self, 'Configurar Poppler', 
            'Digite o caminho completo para a pasta bin do Poppler:\n'
            'Ex: C:\\Poppler\\bin ou /usr/local/bin',
            QLineEdit.Normal,
            self.poppler_path or ""
        )
        
        if ok:
            if path.strip():
                self.poppler_path = path.strip()
                self.poppler_btn.setText(f'Poppler: {self.poppler_path}')
                self.save_config()
                QMessageBox.information(self, "Configuração Salva", 
                                    f"Caminho do Poppler configurado para: {path}")
            else:
                # Se o caminho estiver em branco, limpar a configuração
                self.poppler_path = None
                self.poppler_btn.setText('Configurar Poppler (Opcional)')
                self.save_config()
                QMessageBox.information(self, "Configuração Removida", 
                                    "Caminho do Poppler removido")

    def upload_pdf(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'Selecionar PDF', '', 'Arquivos PDF (*.pdf)', options=options
        )
        
        if file_path:
            self.current_pdf_path = file_path
            self.analyze_pdf(file_path)
    
    def analyze_pdf(self, pdf_path):
        try:
            # Limpar visualizações e logs anteriores
            self.page_list.clear()
            self.clear_preview()
            self.box_table.setRowCount(0)
            self.page_data = []
            self.log_messages = []
            self.log_text.setText("Nenhum erro ou aviso registrado.")
            
            # Adicionar primeira mensagem de log
            self.add_log_message(f"Analisando arquivo: {os.path.basename(pdf_path)}")
            
            # Abrir o PDF e obter informações básicas
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                num_pages = len(pdf_reader.pages)
                
                # Atualizar informações básicas
                file_name = os.path.basename(pdf_path)
                self.info_label.setText(f'Arquivo: {file_name}\nTotal de páginas: {num_pages}')
                
                # Análise adicional com PyMuPDF para formatos problemáticos
                pymupdf_doc = fitz.open(pdf_path)
                
                # Analisar cada página
                page_formats = []
                
                for i in range(num_pages):
                    page = pdf_reader.pages[i]
                    page_info = self.analyze_page_boxes(page, i)
                    
                    # Usar PyMuPDF como backup para obter tamanho da página
                    if not page_info.get('MediaBox') and i < len(pymupdf_doc):
                        pymupdf_page = pymupdf_doc[i]
                        rect = pymupdf_page.rect
                        width_pt = rect.width
                        height_pt = rect.height
                        width_mm = width_pt * 0.352778
                        height_mm = height_pt * 0.352778
                        
                        page_info['MediaBox'] = {
                            'width': width_mm,
                            'height': height_mm,
                            'x': 0,
                            'y': 0,
                            'raw': (0, 0, width_pt, height_pt),
                            'source': 'PyMuPDF' # Marcar que veio do PyMuPDF
                        }
                        
                        self.add_log_message(f"Usando PyMuPDF para obter MediaBox na página {i+1}", "INFO")
                    
                    self.page_data.append(page_info)
                    
                    # Determinar o formato (retrato, paisagem, etc.)
                    mediabox = page_info.get('MediaBox')
                    if mediabox:
                        width, height = mediabox['width'], mediabox['height']
                        if width > height:
                            orientation = "Paisagem"
                        else:
                            orientation = "Retrato"
                        
                        # Formatos padrão de papel
                        formato = self.determine_paper_format(width, height)
                        format_info = f"{formato} ({orientation})"
                        page_formats.append(format_info)
                        
                        # Adicionar à lista
                        source = mediabox.get('source', 'PyPDF2')
                        self.page_list.addItem(f"Página {i+1}: {format_info} [{source}]")
                    else:
                        self.page_list.addItem(f"Página {i+1}: Formato desconhecido")
                        page_formats.append("Desconhecido")
                        self.add_log_message(f"Não foi possível determinar o formato da página {i+1}", "WARNING")
                
                pymupdf_doc.close()
            
            # Verificar se há formatos diferentes
            if len(set(page_formats)) > 1:
                self.format_alert.setText("ALERTA: O documento contém páginas com formatos diferentes!")
                self.add_log_message("O documento contém páginas com formatos diferentes", "WARNING")
            else:
                self.format_alert.setText("")
            
            # Gerar previews das páginas usando PyMuPDF
            self.generate_preview_with_pymupdf(pdf_path, num_pages)
            
            # Selecionar a primeira página automaticamente
            if num_pages > 0:
                self.page_list.setCurrentRow(0)
            
            # Se houver mensagens de log, mostrar a aba de logs
            if len(self.log_messages) > 1:  # Mais que só a mensagem inicial
                self.tabs.setCurrentIndex(1)  # Índice da aba de logs
        
        except Exception as e:
            import traceback
            error_msg = f"Erro ao analisar o PDF: {str(e)}"
            self.info_label.setText(error_msg)
            self.add_log_message(error_msg, "ERROR")
            self.add_log_message(traceback.format_exc(), "ERROR")
            QMessageBox.critical(self, "Erro", error_msg)
    
    def analyze_page_boxes(self, page, page_index):
        page_info = {}
        
        # Lista de possíveis boxes em um PDF
        box_types = ['MediaBox', 'CropBox', 'BleedBox', 'TrimBox', 'ArtBox']
        
        for box_type in box_types:
            try:
                # Verificar se este tipo de box existe na página
                if hasattr(page, box_type.lower()):
                    box = getattr(page, box_type.lower())
                    if box:
                        # Converter todos os valores para float para evitar problemas com Decimal
                        try:
                            x1 = float(box[0])
                            y1 = float(box[1])
                            x2 = float(box[2])
                            y2 = float(box[3])
                        except (TypeError, ValueError) as e:
                            # Se falhar a conversão direta, tentar ver se é Decimal
                            if isinstance(box[0], decimal.Decimal):
                                x1 = float(box[0])
                                y1 = float(box[1])
                                x2 = float(box[2])
                                y2 = float(box[3])
                            else:
                                raise e
                        
                        width = abs(x2 - x1)
                        height = abs(y2 - y1)
                        
                        # Converter de pontos para milímetros
                        width_mm = width * 0.352778
                        height_mm = height * 0.352778
                        x1_mm = x1 * 0.352778
                        y1_mm = y1 * 0.352778
                        
                        page_info[box_type] = {
                            'width': width_mm,
                            'height': height_mm,
                            'x': x1_mm,
                            'y': y1_mm,
                            'raw': (x1, y1, x2, y2)
                        }
            except Exception as e:
                error_msg = f"Erro ao analisar {box_type} na página {page_index+1}: {str(e)}"
                self.add_log_message(error_msg, "WARNING")
        
        return page_info
    
    def on_page_selected(self, current_row):
        if current_row >= 0 and current_row < len(self.page_data):
            # Atualizar a tabela de boxes para a página selecionada
            page_info = self.page_data[current_row]
            self.update_box_table(page_info)
            
            # Atualizar o preview para mostrar apenas a página selecionada
            if self.current_pdf_path:
                self.clear_preview()
                self.generate_single_page_preview(self.current_pdf_path, current_row)
    
    def update_box_table(self, page_info):
        self.box_table.setRowCount(0)
        
        for box_type, box_data in page_info.items():
            row_position = self.box_table.rowCount()
            self.box_table.insertRow(row_position)
            
            # Adicionar informações na tabela
            source = box_data.get('source', 'PyPDF2')
            display_type = f"{box_type} [{source}]" if 'source' in box_data else box_type
            
            self.box_table.setItem(row_position, 0, QTableWidgetItem(display_type))
            self.box_table.setItem(row_position, 1, QTableWidgetItem(f"{box_data['width']:.2f}"))
            self.box_table.setItem(row_position, 2, QTableWidgetItem(f"{box_data['height']:.2f}"))
            self.box_table.setItem(row_position, 3, QTableWidgetItem(f"{box_data['x']:.2f}"))
            self.box_table.setItem(row_position, 4, QTableWidgetItem(f"{box_data['y']:.2f}"))
    
    def determine_paper_format(self, width_mm, height_mm):
        # Ordenar para que width seja sempre o menor valor
        width_mm, height_mm = min(width_mm, height_mm), max(width_mm, height_mm)
        
        # Tolerância para comparações (em mm)
        tolerance = 5
        
        # Verificar formatos comuns
        if abs(width_mm - 210) < tolerance and abs(height_mm - 297) < tolerance:
            return "A4"
        elif abs(width_mm - 216) < tolerance and abs(height_mm - 279) < tolerance:
            return "Carta"
        elif abs(width_mm - 216) < tolerance and abs(height_mm - 356) < tolerance:
            return "Ofício"
        elif abs(width_mm - 297) < tolerance and abs(height_mm - 420) < tolerance:
            return "A3"
        elif abs(width_mm - 148) < tolerance and abs(height_mm - 210) < tolerance:
            return "A5"
        else:
            return f"Personalizado ({width_mm:.1f}mm x {height_mm:.1f}mm)"
    
    def generate_preview_with_pymupdf(self, pdf_path, num_pages):
        try:
            # Usar PyMuPDF (fitz) para gerar previews de todas as páginas
            pdf_document = fitz.open(pdf_path)
            max_pages = min(num_pages, 10)  # Limitar a 10 páginas para melhor desempenho
            
            for i in range(max_pages):
                # Obter a página
                page = pdf_document.load_page(i)
                
                # Renderizar em um pixmap
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))  # Escala de 0.5 para reduzir tamanho
                
                # Converter para QImage e QPixmap
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                # Adicionar imagem ao layout de visualização
                preview_label = QLabel()
                preview_label.setPixmap(pixmap)
                preview_label.setAlignment(Qt.AlignCenter)
                
                # Adicionar título da página
                page_title = QLabel(f"Página {i+1}")
                page_title.setAlignment(Qt.AlignCenter)
                page_title.setStyleSheet("font-weight: bold; margin-top: 15px;")
                
                self.preview_layout.addWidget(page_title)
                self.preview_layout.addWidget(preview_label)
            
            pdf_document.close()
        
        except Exception as e:
            error_msg = f"Erro ao gerar previews: {str(e)}"
            self.add_log_message(error_msg, "ERROR")
            error_label = QLabel(error_msg)
            error_label.setStyleSheet("color: red;")
            self.preview_layout.addWidget(error_label)
    
    def generate_single_page_preview(self, pdf_path, page_index):
        try:
            # Gerar preview apenas para a página selecionada
            pdf_document = fitz.open(pdf_path)
            
            # Obter a página
            page = pdf_document.load_page(page_index)
            
            # Renderizar em um pixmap com escala maior por ser apenas uma página
            pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
            
            # Converter para QImage e QPixmap
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            
            # Redimensionar para caber na área de visualização
            pixmap = pixmap.scaled(self.scroll_area.width() - 30, 
                                  int(self.scroll_area.width() * pixmap.height() / pixmap.width()),
                                  Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            # Adicionar imagem ao layout de visualização
            preview_label = QLabel()
            preview_label.setPixmap(pixmap)
            preview_label.setAlignment(Qt.AlignCenter)
            
            # Adicionar título da página
            page_title = QLabel(f"Página {page_index+1}")
            page_title.setAlignment(Qt.AlignCenter)
            page_title.setStyleSheet("font-weight: bold; margin-top: 15px;")
            
            self.preview_layout.addWidget(page_title)
            self.preview_layout.addWidget(preview_label)
            
            # Adicionar visualização de Boxes
            self.add_box_visualization(page_index)
            
            pdf_document.close()
        
        except Exception as e:
            error_msg = f"Erro ao gerar preview: {str(e)}"
            self.add_log_message(error_msg, "ERROR")
            error_label = QLabel(error_msg)
            error_label.setStyleSheet("color: red;")
            self.preview_layout.addWidget(error_label)
    
    def add_box_visualization(self, page_index):
        """
        Adiciona uma explicação visual dos diferentes boxes
        """
        if page_index < len(self.page_data):
            page_info = self.page_data[page_index]
            
            if page_info:
                box_info_label = QLabel()
                box_info_label.setWordWrap(True)
                box_info_label.setStyleSheet("margin-top: 20px;")
                
                # Criar texto explicativo
                box_text = "<b>Boxes na página:</b><br>"
                
                # Adicionar legenda de cores para cada tipo de box
                colors = {
                    'MediaBox': 'blue',
                    'CropBox': 'green',
                    'BleedBox': 'red',
                    'TrimBox': 'orange',
                    'ArtBox': 'purple'
                }
                
                for box_type in page_info.keys():
                    width = page_info[box_type]['width']
                    height = page_info[box_type]['height']
                    color = colors.get(box_type, 'black')
                    source = page_info[box_type].get('source', '')
                    source_info = f" [{source}]" if source else ""
                    box_text += f"<span style='color:{color};'>■</span> <b>{box_type}{source_info}:</b> {width:.2f}mm x {height:.2f}mm<br>"
                
                # Explicação dos boxes
                box_text += "<br><b>Significado dos boxes:</b><br>"
                box_text += "• <b>MediaBox:</b> Define o tamanho total da página.<br>"
                box_text += "• <b>CropBox:</b> Define a região visível da página quando exibida ou impressa.<br>"
                box_text += "• <b>BleedBox:</b> Define a região de sangria para impressão.<br>"
                box_text += "• <b>TrimBox:</b> Define o tamanho final da página após o corte.<br>"
                box_text += "• <b>ArtBox:</b> Define a área de conteúdo significativo da página.<br>"
                
                box_info_label.setText(box_text)
                self.preview_layout.addWidget(box_info_label)
    
    def clear_preview(self):
        # Limpar layout de preview
        while self.preview_layout.count():
            item = self.preview_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFAnalyzerApp()
    ex.show()
    sys.exit(app.exec_())
