import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QPushButton, QWidget, QComboBox, QHBoxLayout, QVBoxLayout, QRadioButton, QButtonGroup, QLineEdit, QLabel, QCheckBox, QListWidget, QTabWidget
import chardet

class CsvExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV/Excel Processor')
        self.setGeometry(150, 150, 800, 600)

        # Create tab widget
        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        # Create tabs
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()

        # Add tabs
        self.tabs.addTab(self.tab1, "Redigere CSV")
        self.tabs.addTab(self.tab2, "Forberede org-import")
        self.tabs.addTab(self.tab3, "Forberede brukerimport")

        self.setupTab1()
        self.setupTab2()
        self.setupTab3()

    def setupTab1(self):
        layout = QVBoxLayout(self.tab1)
        openFileButton = QPushButton('Åpne fil', self.tab1)
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
        self.columnComboBox1 = QComboBox(self.tab1)
        self.columnComboBox2 = QComboBox(self.tab1)
        self.columnComboBox3 = QComboBox(self.tab1)
        col1Layout.addWidget(self.columnComboBox1)
        col3Layout.addWidget(self.columnComboBox2)
        col5Layout.addWidget(self.columnComboBox3)

        

        # Radioknapper og deres gruppering
        self.radioGroup1 = QButtonGroup(self.tab1)
        for text in ["ER LIK NOEN LINJE", "ER LIK SAMME LINJE", "ER IKKE LIK NOEN LINJE", "ER IKKE LIK SAMME LINJE"]:
            radioBtn = QRadioButton(text, self.tab1)
            self.radioGroup1.addButton(radioBtn)
            col2Layout.addWidget(radioBtn)

        self.radioGroup2 = QButtonGroup(self.tab1)
        for text in ["SKRIV TEKST FRA TEKSTFELT", "SKRIV TEKST FRA DATAFELT"]:
            radioBtn = QRadioButton(text, self.tab1)
            self.radioGroup2.addButton(radioBtn)
            col4Layout.addWidget(radioBtn)

        # Legg til ekstra nedtrekksmeny og tekstfelt
        self.extraColumnComboBox = QComboBox(self.tab1)
        self.extraColumnComboBox.hide()  # Skjul som standard
        col4Layout.addWidget(self.extraColumnComboBox)

        self.lineEdit = QLineEdit(self.tab1)
        self.lineEdit.hide()  # Skjul som standard
        col4Layout.addWidget(self.lineEdit)

        # Koble radioknapp-signaler til en metode for å vise/skjule tekstfelt og nedtrekksmeny
        self.radioGroup2.buttonClicked.connect(self.handleRadioSelection)

        # Opprettelse av avkrysningsbokser for eksportformat
        self.exportCsvCheckBox = QCheckBox("Eksporter som CSV", self.tab1)
        self.exportExcelCheckBox = QCheckBox("Eksporter som Excel", self.tab1)
        layout.addWidget(self.exportCsvCheckBox)
        layout.addWidget(self.exportExcelCheckBox)

        # Checkbox for å velge sletting av kolonner
        self.deleteColumnsCheckBox = QCheckBox("Slett kolonner", self.tab1)
        layout.addWidget(self.deleteColumnsCheckBox)
        self.deleteColumnsCheckBox.toggled.connect(self.toggleColumnDeletionList)

        # Liste for å velge kolonner som skal slettes
        self.columnDeletionList = QListWidget(self.tab1)
        self.columnDeletionList.setSelectionMode(QListWidget.ExtendedSelection)
        self.columnDeletionList.setHidden(True)
        layout.addWidget(self.columnDeletionList)

        bottomLayout = QHBoxLayout()

        # Opprettelse av eksportknapp
        exportButton = QPushButton('Eksporter', self.tab1)
        exportButton.clicked.connect(self.exportData)
        bottomLayout.addWidget(exportButton)

        # Legge til kolonnelayouts til hovedlayout
        mainLayout.addLayout(col1Layout)
        mainLayout.addLayout(col2Layout)
        mainLayout.addLayout(col3Layout)
        mainLayout.addLayout(col4Layout)
        mainLayout.addLayout(col5Layout)
        layout.addLayout(mainLayout)
        layout.addLayout(bottomLayout)

    def setupTab2(self):
        self.tab2_layout = QVBoxLayout(self.tab2)
        

        openFileButton = QPushButton('Åpne fil', self.tab2)
        openFileButton.clicked.connect(self.openFileNameDialog_tab2)
        self.tab2_layout.addWidget(openFileButton)

        # Tekstfelt for Navn nivå 0
        self.tab2_layout.addWidget(QLabel("Navn nivå 0:"))
        self.level_0_name_input = QLineEdit(self.tab2)
        self.tab2_layout.addWidget(self.level_0_name_input)

        # Tekstfelt for Prefiks ID number
        self.tab2_layout.addWidget(QLabel("Prefiks ID number:"))
        self.id_number_prefix_input = QLineEdit(self.tab2)
        self.tab2_layout.addWidget(self.id_number_prefix_input)

        # Knapp for å åpne fil og starte konverteringsprosessen
        self.convert_button = QPushButton('Konverter og Eksporter CSV', self.tab2)
        self.convert_button.clicked.connect(self.convert_and_export_org_csv)
        self.tab2_layout.addWidget(self.convert_button)

        

        # Opprette hoved horisontalt layout
      #  mainLayout = QHBoxLayout()
        

        # Opprette vertikale layouts for hver kolonne
       # col1Layout = QVBoxLayout()
       # col2Layout = QVBoxLayout()
       # col3Layout = QVBoxLayout()
        

    def setupTab3(self):
        layout = QVBoxLayout(self.tab3)

        openFileButton = QPushButton('Åpne fil', self.tab3)
        openFileButton.clicked.connect(self.openFileNameDialog)
        layout.addWidget(openFileButton)

        # Opprette hoved horisontalt layout
        mainLayout = QHBoxLayout()
        

        # Opprette vertikale layouts for hver kolonne
        col1Layout = QVBoxLayout()
        col2Layout = QVBoxLayout()
        col3Layout = QVBoxLayout()


    
        

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Åpne CSV eller Excel-fil", "", "Alle filer (*);;CSV-filer (*.csv);;Excel-filer (*.xlsx)", options=options)
        if fileName:
            self.loadFile(fileName)

    def openFileNameDialog_tab2(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Åpne CSV-fil", "", "CSV-filer (*.csv)", options=options)
        if file_name:
            self.org_csv_file_path = file_name

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

    def convert_and_export_org_csv(self):
        if self.org_csv_file_path:
            # Les inn filen
            with open(self.org_csv_file_path, 'rb') as file:
                encoding = chardet.detect(file.read(100000))['encoding']
            df = pd.read_csv(self.org_csv_file_path, sep=';', encoding=encoding, header=None)

            # Identifiser header-row
            header_row = df[df.apply(lambda x: x.str.contains(r'Enhetstype \(nivå 1\)', na=False, regex=True)).any(axis=1)].index[0]
            df = pd.read_csv(self.org_csv_file_path, sep=';', encoding=encoding, header=header_row)

            # Generer kolonnene
            top_level_name = self.level_0_name_input.text()
            id_prefix = self.id_number_prefix_input.text()
            df_final = self.generate_columns(df, top_level_name, id_prefix)

            # Eksport til ny CSV-fil
            export_file_path, _ = QFileDialog.getSaveFileName(self, "Lagre fil som", "", "CSV-filer (*.csv)")
            if export_file_path:
                df_final.to_csv(export_file_path, sep=';', encoding='utf-8', index=False)

    def find_parent(self, df, level, row):
        # Topplevel har ingen parent.
        if level == 0:
            return "Top"
        if level == 1:
            return self.level_0_name_input.text()
        # Finn parent ved å gå bakover i nivåene til vi finner en ikke-null enhetskode.
        for parent_level in range(level - 1, 0, -1):
            parent_code_col = f'Enhetskode (nivå {parent_level})'
            parent_name_col = f'Enhetsnavn (nivå {parent_level})'
            if pd.notna(row[parent_code_col]):
                return f"{row[parent_code_col]} - {row[parent_name_col]}"
        return None

    def generate_columns(self, df, top_level_name, id_prefix):
        # Initialiserer output DataFrame
        output_df = pd.DataFrame(columns=["Name", "ID number", "Description", "Parent", "ID Sort Key"])
        id_counter = 1  # Starter telleren fra 1 for toppnivået

        # Legger til toppnivået
        output_df.loc[id_counter] = {
            "Name": top_level_name, 
            "ID number": f"{id_prefix}-{id_counter}", 
            "Description": "", 
            "Parent": "Top",
            "ID Sort Key": id_counter
        }
        id_counter += 1

        # Prosesser hver enhetstype nivå for nivå
        for level in range(1, df.filter(like='Enhetstype (nivå').shape[1] + 1):
            # Prosesser hver rad i DataFrame
            for _, row in df.iterrows():
                # Sjekk for enhetsnavn på gjeldende nivå
                name_col = f'Enhetsnavn (nivå {level})'
                code_col = f'Enhetskode (nivå {level})'

                if pd.notna(row[name_col]):
                    name = f"{row[code_col]} - {row[name_col]}"
                    # Finn parent basert på nivået over
                    parent = self.find_parent(df, level, row)

                    # Legg til raden i output_df
                    output_df.loc[id_counter] = {
                        "Name": name, 
                        "ID number": f"{id_prefix}-{id_counter}", 
                        "Description": "", 
                        "Parent": parent,
                        "ID Sort Key": id_counter
                    }
                    id_counter += 1

        # Fjern duplikate navn og sorter basert på ID Sort Key
        output_df = output_df.drop_duplicates(subset='Name')
        output_df = output_df.sort_values(by='ID Sort Key').reset_index(drop=True)
        for index, _ in enumerate(output_df.index):
            output_df.at[index + 1, 'ID number'] = f"{id_prefix}-{index + 1}"
        output_df.drop('ID Sort Key', axis=1, inplace=True)  # Fjerner den midlertidige sorteringkolonnen
        output_df = output_df[output_df['Name'].notna() & (output_df['Name'].str.strip() != '')]

        return output_df
    
        



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