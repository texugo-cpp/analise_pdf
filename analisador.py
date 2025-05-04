import sys
import os
import tempfile
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
        self.initUI()
        self.current_pdf_path = None
        self.page_images = []
        self.poppler_path = None

    def initUI(self):
        self.setWindowTitle('Analisador de PDF')
        self.setGeometry(100, 100, 1200, 700)

        # Layout principal
        main_layout = QHBoxLayout()
        
        # Painel esquerdo para controles e informações básicas
        left_panel = QVBoxLayout()
        
        # Botão para configurar Poppler (opcional)
        self.poppler_btn = QPushButton('Configurar Poppler (Opcional)', self)
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
        
    def set_poppler_path(self):
        path, ok = QInputDialog.getText(
            self, 'Configurar Poppler', 
            'Digite o caminho completo para a pasta bin do Poppler:\n'
            'Ex: C:\\Poppler\\bin ou /usr/local/bin',
            QLineEdit.Normal
        )
        
        if ok and path:
            self.poppler_path = path
            QMessageBox.information(self, "Configuração Salva", 
                                  f"Caminho do Poppler configurado para: {path}")

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
            # Limpar visualizações anteriores
            self.page_list.clear()
            self.clear_preview()
            self.box_table.setRowCount(0)
            self.page_data = []
            
            # Abrir o PDF e obter informações básicas
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                num_pages = len(pdf_reader.pages)
                
                # Atualizar informações básicas
                file_name = os.path.basename(pdf_path)
                self.info_label.setText(f'Arquivo: {file_name}\nTotal de páginas: {num_pages}')
                
                # Analisar cada página
                page_formats = []
                
                for i in range(num_pages):
                    page = pdf_reader.pages[i]
                    page_info = self.analyze_page_boxes(page, i)
                    self.page_data.append(page_info)
                    
                    # Determinar o formato (retrato, paisagem, etc.)
                    mediabox = page_info['MediaBox'] if 'MediaBox' in page_info else None
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
                        self.page_list.addItem(f"Página {i+1}: {format_info}")
                    else:
                        self.page_list.addItem(f"Página {i+1}: Formato desconhecido")
                        page_formats.append("Desconhecido")
            
            # Verificar se há formatos diferentes
            if len(set(page_formats)) > 1:
                self.format_alert.setText("ALERTA: O documento contém páginas com formatos diferentes!")
            else:
                self.format_alert.setText("")
            
            # Gerar previews das páginas usando PyMuPDF
            self.generate_preview_with_pymupdf(pdf_path, num_pages)
            
            # Selecionar a primeira página automaticamente
            if num_pages > 0:
                self.page_list.setCurrentRow(0)
        
        except Exception as e:
            import traceback
            self.info_label.setText(f"Erro ao analisar o PDF: {str(e)}")
            QMessageBox.critical(self, "Erro", f"Erro ao analisar o PDF: {str(e)}\n\n{traceback.format_exc()}")
    
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
                        x1, y1, x2, y2 = box
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
                print(f"Erro ao analisar {box_type} na página {page_index+1}: {str(e)}")
        
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
            self.box_table.setItem(row_position, 0, QTableWidgetItem(box_type))
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
            error_label = QLabel(f"Erro ao gerar previews: {str(e)}")
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
            error_label = QLabel(f"Erro ao gerar preview: {str(e)}")
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
                    box_text += f"<span style='color:{color};'>■</span> <b>{box_type}:</b> {width:.2f}mm x {height:.2f}mm<br>"
                
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
