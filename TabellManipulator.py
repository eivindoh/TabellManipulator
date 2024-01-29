import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QPushButton, QWidget, QComboBox, QHBoxLayout, QVBoxLayout, QRadioButton, QButtonGroup, QLineEdit, QCheckBox, QListWidget
import chardet

class CsvExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
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

        # Opprette hoved horisontalt layout
        mainLayout = QHBoxLayout()
        

        # Opprette vertikale layouts for hver kolonne
        col1Layout = QVBoxLayout()
        col2Layout = QVBoxLayout()
        col3Layout = QVBoxLayout()
        col4Layout = QVBoxLayout()
        col5Layout = QVBoxLayout()

        # Nedtrekksmenyer for kolonnevalg
        self.columnComboBox1 = QComboBox(self)
        self.columnComboBox2 = QComboBox(self)
        self.columnComboBox3 = QComboBox(self)
        col1Layout.addWidget(self.columnComboBox1)
        col3Layout.addWidget(self.columnComboBox2)
        col5Layout.addWidget(self.columnComboBox3)

        

        # Radioknapper og deres gruppering
        self.radioGroup1 = QButtonGroup(self)
        for text in ["ER LIK NOEN LINJE", "ER LIK SAMME LINJE", "ER IKKE LIK NOEN LINJE", "ER IKKE LIK SAMME LINJE"]:
            radioBtn = QRadioButton(text, self)
            self.radioGroup1.addButton(radioBtn)
            col2Layout.addWidget(radioBtn)

        self.radioGroup2 = QButtonGroup(self)
        for text in ["SKRIV TEKST FRA TEKSTFELT", "SKRIV TEKST FRA DATAFELT"]:
            radioBtn = QRadioButton(text, self)
            self.radioGroup2.addButton(radioBtn)
            col4Layout.addWidget(radioBtn)

        # Legg til ekstra nedtrekksmeny og tekstfelt
        self.extraColumnComboBox = QComboBox(self)
        self.extraColumnComboBox.hide()  # Skjul som standard
        col4Layout.addWidget(self.extraColumnComboBox)

        self.lineEdit = QLineEdit(self)
        self.lineEdit.hide()  # Skjul som standard
        col4Layout.addWidget(self.lineEdit)

        # Koble radioknapp-signaler til en metode for å vise/skjule tekstfelt og nedtrekksmeny
        self.radioGroup2.buttonClicked.connect(self.handleRadioSelection)

        # Opprettelse av avkrysningsbokser for eksportformat
        self.exportCsvCheckBox = QCheckBox("Eksporter som CSV", self)
        self.exportExcelCheckBox = QCheckBox("Eksporter som Excel", self)
        layout.addWidget(self.exportCsvCheckBox)
        layout.addWidget(self.exportExcelCheckBox)

        # Checkbox for å velge sletting av kolonner
        self.deleteColumnsCheckBox = QCheckBox("Slett kolonner", self)
        layout.addWidget(self.deleteColumnsCheckBox)
        self.deleteColumnsCheckBox.toggled.connect(self.toggleColumnDeletionList)

        # Liste for å velge kolonner som skal slettes
        self.columnDeletionList = QListWidget(self)
        self.columnDeletionList.setSelectionMode(QListWidget.ExtendedSelection)
        self.columnDeletionList.setHidden(True)
        layout.addWidget(self.columnDeletionList)

        bottomLayout = QHBoxLayout()

        # Opprettelse av eksportknapp
        exportButton = QPushButton('Eksporter', self)
        exportButton.clicked.connect(self.exportData)
        bottomLayout.addWidget(exportButton)

        # Legge til kolonnelayouts til hovedlayout
        mainLayout.addLayout(col1Layout)
        mainLayout.addLayout(col2Layout)
        mainLayout.addLayout(col3Layout)
        mainLayout.addLayout(col4Layout)
        mainLayout.addLayout(col5Layout)
        centralWidget.layout().addLayout(mainLayout)
        centralWidget.layout().addLayout(bottomLayout)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Åpne CSV eller Excel-fil", "", "Alle filer (*);;CSV-filer (*.csv);;Excel-filer (*.xlsx)", options=options)
        if fileName:
            self.loadFile(fileName)

    def toggleColumnDeletionList(self, checked):
        if self.df.empty:
            print("Ingen data å jobbe med")
            return
        
        self.columnDeletionList.setHidden(not checked)
        if checked:
            self.columnDeletionList.addItems(self.df.columns)

    def applyColumnDeletions(self):
        if self.deleteColumnsCheckBox.isChecked():
            selected_items = [item.text() for item in self.columnDeletionList.selectedItems()]
            self.df.drop(columns=selected_items, inplace=True)
            

    def loadFile(self, filePath):
        # Bruk chardet for å gjenkjenne tegnsettet
        with open(filePath, 'rb') as file:
            encoding = chardet.detect(file.read(100000))['encoding']

        # Hvis chardet ikke finner et tegnsett, bruker vi UTF-8 som fallback
        if not encoding:
            encoding = 'utf-8'

        # Les inn filen med gjenkjent tegnsett
        if filePath.endswith('.csv'):
            self.df = pd.read_csv(filePath, sep=';', encoding=encoding)
        elif filePath.endswith('.xlsx'):
            self.df = pd.read_excel(filePath)
        else:
            print("Ugyldig filformat")
            return
            
        self.updateColumnComboBoxes()

    def applyRules(self):
        if self.df.empty:
            print("Ingen data å jobbe med")
            return

        source_column = self.columnComboBox1.currentText()
        target_column = self.columnComboBox3.currentText()
        condition = self.radioGroup1.checkedButton().text()
        action = self.radioGroup2.checkedButton().text()
        comparison_column = self.columnComboBox2.currentText()

        if condition == "ER LIK NOEN LINJE":
            matching_rows = self.df[source_column].isin(self.df[comparison_column])
        elif condition == "ER LIK SAMME LINJE":
            matching_rows = self.df[source_column] == self.df[comparison_column]
        else:
            print("Ugyldig betingelse valgt")
            return

        # Utfør handlingen basert på valgt betingelse
        if action == "SKRIV TEKST FRA TEKSTFELT":
            text_to_write = self.lineEdit.text()
            self.df.loc[matching_rows, target_column] = text_to_write
        elif action == "SKRIV TEKST FRA DATAFELT":
            data_field_to_copy = self.extraColumnComboBox.currentText()
            self.df.loc[matching_rows, target_column] = self.df[data_field_to_copy]

    def convert_floats_to_ints(self, df):
        for col in df.select_dtypes(include=['float']).columns:
            if all(df[col].dropna().apply(float.is_integer)):
                df[col] = df[col].astype('Int64')
        return df

    def exportData(self):
        self.applyRules()
        self.applyColumnDeletions()
        self.df = self.convert_floats_to_ints(self.df)

        if self.df.empty:
            print("Ingen data å eksportere")
            return

        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getSaveFileName(self, "Lagre fil som", "", "Alle filer (*);;CSV-filer (*.csv);;Excel-filer (*.xlsx)", options=options)
        
        if not filePath:
            return

        if self.exportCsvCheckBox.isChecked():
            self.df.to_csv(filePath, sep=';', encoding='utf-8', index=False)
        elif self.exportExcelCheckBox.isChecked():
            self.df.to_excel(filePath, index=False)
        else:
            print("Vennligst velg et eksportformat")

    def handleRadioSelection(self, button):
        if button.text() == "SKRIV TEKST FRA TEKSTFELT":
            self.lineEdit.show()
            self.extraColumnComboBox.hide()
        elif button.text() == "SKRIV TEKST FRA DATAFELT":
            self.lineEdit.hide()
            self.extraColumnComboBox.show()

    def updateColumnComboBoxes(self):
        columns = self.df.columns
        for comboBox in [self.columnComboBox1, self.columnComboBox2, self.columnComboBox3, self.extraColumnComboBox]:
            comboBox.clear()
            comboBox.addItems(columns)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CsvExcelProcessor()
    ex.show()
    sys.exit(app.exec_())