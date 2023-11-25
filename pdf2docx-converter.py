from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QHBoxLayout, QFileDialog, QMessageBox
from pdf2docx import Converter
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtGui import QFont

class PDFtoDOCXConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PDF to DOCX Converter")
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Ekran çözünürlüğü
        screen = QGuiApplication.screens()[0].size()
        width = int(screen.width()/2.0)-350
        height = int(screen.height()/2.0)-75
        self.setGeometry(width,height,700,150)

        main_layout = QVBoxLayout(central_widget)

        # İlk satır
        row1_layout = QHBoxLayout()
        label1 = QLabel("PDF Dosyası:")
        row1_layout.addWidget(label1)
        self.pdf_line_edit = QLineEdit()
        row1_layout.addWidget(self.pdf_line_edit)
        browse_pdf_button = QPushButton("PDF Seç")
        browse_pdf_button.clicked.connect(self.get_pdf_file)
        row1_layout.addWidget(browse_pdf_button)
        main_layout.addLayout(row1_layout)

        # İkinci satır
        row2_layout = QHBoxLayout()
        label2 = QLabel("DOCX Dosyası:")
        row2_layout.addWidget(label2)
        self.docx_line_edit = QLineEdit()
        row2_layout.addWidget(self.docx_line_edit)
        browse_docx_button = QPushButton("DOCX Kaydet")
        browse_docx_button.clicked.connect(self.save_docx_file)
        row2_layout.addWidget(browse_docx_button)
        main_layout.addLayout(row2_layout)

        # Altta uzatılmış buton
        convert_button = QPushButton("PDF'den DOCX'e Dönüştür")
        convert_button.clicked.connect(self.pdf_to_docx)
        main_layout.addWidget(convert_button)

    def set_styles(self,app):
        font = QFont()
        font.setPointSize(12)
        app.setFont(font)

    def get_pdf_file(self):
        file_dialog = QFileDialog()
        filename, _ = file_dialog.getOpenFileName(self, "PDF Dosyasını Seç", "", "PDF Dosyaları (*.pdf)")
        self.pdf_line_edit.setText(filename)

    def save_docx_file(self):
        file_dialog = QFileDialog()
        filename, _ = file_dialog.getSaveFileName(self, "DOCX Dosyasını Kaydet", "", "DOCX Dosyaları (*.docx)")
        self.docx_line_edit.setText(filename)

    def pdf_to_docx(self):
        pdf_filename = self.pdf_line_edit.text()
        docx_filename = self.docx_line_edit.text()

        if not pdf_filename or not docx_filename:
            self.show_error_message()
            return

        # PDF dosyasını DOCX'e dönüştür
        cv = Converter(pdf_filename)
        cv.convert(docx_filename, start=0, end=None)
        cv.close()
        self.show_success_message()

    def show_error_message(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Hata")
        msg_box.setText("PDF ve DOCX dosyalarını seçmelisiniz!")
        msg_box.exec()

    def show_success_message(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Dönüşüm Tamamlandı")
        msg_box.setText("PDF'den DOCX'e dönüşüm tamamlandı.")
        msg_box.exec()

def run_app():
    app = QApplication([])
    window = PDFtoDOCXConverter()
    window.set_styles(app)
    window.show()
    app.exec()

run_app()
