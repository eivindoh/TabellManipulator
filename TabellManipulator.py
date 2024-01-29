import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QPushButton, QWidget

class CsvExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV/Excel Processor')
        self.setGeometry(100, 100, 600, 400)

        centralWidget = QWidget(self)
        self.setCentralWidget(centralWidget)
        layout = QVBoxLayout(centralWidget)

        openFileButton = QPushButton('Åpne fil', self)
        openFileButton.clicked.connect(self.openFileNameDialog)
        layout.addWidget(openFileButton)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Åpne CSV eller Excel-fil", "", "Alle filer (*);;CSV-filer (*.csv);;Excel-filer (*.xlsx)", options=options)
        if fileName:
            print("Valgt fil:", fileName)  # Her vil vi legge til funksjonalitet for å lese og behandle filen.

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CsvExcelProcessor()
    ex.show()
    sys.exit(app.exec_())