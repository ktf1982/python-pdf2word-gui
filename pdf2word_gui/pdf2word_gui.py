import sys
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PySide6.QtCore import Qt
from pdf2docx import Converter
import os

class PDFtoWordConverter(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('PDF to Word Converter')
        self.setGeometry(100, 100, 300, 200)

        layout = QVBoxLayout()

        self.label = QLabel('Select a PDF file to convert:')
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.select_button = QPushButton('Select PDF File')
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        self.convert_button = QPushButton('Convert to Word')
        self.convert_button.clicked.connect(self.convert_file)
        layout.addWidget(self.convert_button)

        self.setLayout(layout)

    def select_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select PDF File', '', 'PDF Files (*.pdf)', options=options)
        if file_name:
            self.label.setText(f'Selected: {file_name}')
            self.pdf_file = file_name

    def convert_file(self):
        if hasattr(self, 'pdf_file'):
            word_file = os.path.splitext(self.pdf_file)[0] + '.docx'
            try:
                cv = Converter(self.pdf_file)
                cv.convert(word_file)
                cv.close()
                QMessageBox.information(self, 'Success', f'PDF has been converted to Word: {word_file}')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to convert PDF: {str(e)}')
        else:
            QMessageBox.warning(self, 'Warning', 'Please select a PDF file first.')

def main():
    app = QApplication(sys.argv)
    window = PDFtoWordConverter()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
