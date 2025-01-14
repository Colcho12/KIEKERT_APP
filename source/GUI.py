## IMPORT STATEMENTS
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QFileDialog, QLabel, QFormLayout, QComboBox
)
from PySide6.QtCore import Qt
import fase1
import fase2
import fase3
import pandas as pd

## MAIN WINDOW
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Proyecto Kiekert')
        self.setGeometry(300, 100, 600, 400)

        ## MAIN WIDGET
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)


        ## LOADING THE EXCEL FILE
        self.label_file = QLabel('Step 1: Select Excel File')
        self.label_file.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.label_file)

        self.load_button = QPushButton('Select Excel File')
        self.load_button.clicked.connect(self.load_excel_file)
        self.layout.addWidget(self.load_button)

        ## SHEET NAMES AS DROPDOWNS LOADED WITH THE FILE
        self.label_sheets = QLabel('Step 2: Select Sheets for each stage:')
        self.label_sheets.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.label_sheets)

        self.form_layout = QFormLayout()
        self.sheet_dropdown_fase1 = QComboBox()
        self.sheet_dropdown_fase2 = QComboBox()
        self.sheet_dropdown_fase3 = QComboBox()
        self.sheet_dropdown_write = QComboBox()

        self.form_layout.addRow('Sheet for Fase1:', self.sheet_dropdown_fase1)
        self.form_layout.addRow('Sheet for Fase2:', self.sheet_dropdown_fase2)
        self.form_layout.addRow('Sheet for Fase3:', self.sheet_dropdown_fase3)
        self.form_layout.addRow('Sheet for Writing:', self.sheet_dropdown_write)
        self.layout.addLayout(self.form_layout)

        self.submit_sheets_button = QPushButton('Submit Sheet Names')
        self.submit_sheets_button.clicked.connect(self.submit_sheets)
        self.submit_sheets_button.setEnabled(False)
        self.layout.addWidget(self.submit_sheets_button)

        ## data attributes
        self.file_path = None
        self.sheet_names = []

    def load_excel_file(self):
        self.file_path, _ = QFileDialog.getOpenFileName(self, 'Select an Excel file', "", "Excel Files (*.xlsx)")
        if self.file_path:
            try:
                excel_file = pd.ExcelFile(self.file_path)
                availablesheets = excel_file.sheet_names

                self.sheet_dropdown_fase1.addItems(availablesheets)
                self.sheet_dropdown_fase2.addItems(availablesheets)
                self.sheet_dropdown_fase3.addItems(availablesheets)
                self.sheet_dropdown_write.addItems(availablesheets)

                self.label_file.setText(f'Selected File: {self.file_path}')
                self.submit_sheets_button.setEnabled(True)
            except Exception as e:
                self.label_file.setText(f"Error reading Excel file: {str(e)}")
    
    def submit_sheets(self):
        self.sheet_names = [
            self.sheet_dropdown_fase1.currentText(),
            self.sheet_dropdown_fase2.currentText(),
            self.sheet_dropdown_fase3.currentText(),
            self.sheet_dropdown_write.currentText()
            ]
        if all(self.sheet_names):
            self.label_sheets.setText("Sheet names submitted successfully.")
            # self.process_button.setEnabled(True)            
        else:
            self.label_sheets.setText("Please select all sheet names.")

if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()