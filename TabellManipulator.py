import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QPushButton, QWidget, QComboBox, QHBoxLayout, QVBoxLayout, QRadioButton, QButtonGroup, QLineEdit, QLabel, QMessageBox, QCheckBox, QListWidget, QTabWidget
import chardet
import logging
import traceback

class CsvExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.initUI()
        # Oppsett for logging
        logging.basicConfig(filename='Tabellmanipulator.log', 
                            filemode='w', 
                            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                            level=logging.WARNING) # Dette vil fange opp WARNING, ERROR og CRITICAL meldinger

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

        self.createColumnCheckBox = QCheckBox("Opprett ny kolonne for tekst", self.tab1)
        col4Layout.addWidget(self.createColumnCheckBox)

        self.newColumnNameLineEdit = QLineEdit(self.tab1)
        self.newColumnNameLineEdit.setPlaceholderText("Kolonnenavn")
        self.newColumnNameLineEdit.hide()  # Skjul som standard
        col4Layout.addWidget(self.newColumnNameLineEdit)

# Oppdater visningen basert på checkbox-tilstanden
        self.createColumnCheckBox.stateChanged.connect(self.updateColumnInputMethod)

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


    def updateColumnInputMethod(self):
        if self.createColumnCheckBox.isChecked():
            self.newColumnNameLineEdit.show()
            self.columnComboBox3.hide()  # Skjul nedtrekksmenyen når checkbox er avkrysset
        else:
            self.newColumnNameLineEdit.hide()
            self.columnComboBox3.show()
        

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
            self.updateColumnDeletionList()
            self.columnDeletionList.addItems(self.df.columns)

    def updateColumnDeletionList(self):
        self.columnDeletionList.clear()  # Fjern eksisterende elementer
        if self.df is not None:
            columns = self.df.columns
            for column in columns:
                self.columnDeletionList.addItem(column)

    def applyColumnDeletions(self):
        if self.deleteColumnsCheckBox.isChecked():
            selected_items = [item.text() for item in self.columnDeletionList.selectedItems()]
            self.df.drop(columns=selected_items, inplace=True)
            

    def loadFile(self, filePath):
    # Bruk chardet for å gjenkjenne tegnsettet
        with open(filePath, 'rb') as file:
            encoding_result = chardet.detect(file.read(100000))
            encoding = encoding_result['encoding']

        # Hvis chardet ikke finner et tegnsett, bruker vi UTF-8 som fallback
        if not encoding:
            encoding = 'utf-8'
        
        # Les de første linjene for å gjette separator
        with open(filePath, 'r', encoding=encoding) as file:
            sample = file.read(2048)  # Les en liten del av filen for å gjette separator
            separator = self.guess_separator(sample)

        # Les inn filen med gjenkjent tegnsett og separator
        if filePath.endswith('.csv'):
            self.df = pd.read_csv(filePath, sep=separator, encoding=encoding)
        elif filePath.endswith('.xlsx'):
            self.df = pd.read_excel(filePath)
        else:
            print("Ugyldig filformat")
            return
        
        self.updateColumnDeletionList()
        self.updateColumnComboBoxes()

    def guess_separator(self, sample):
        separators = [',', ';', '\t', '|']
        separator_counts = {sep: sample.count(sep) for sep in separators}
        guessed_separator = max(separator_counts, key=separator_counts.get)
        print(f"Separator counts: {separator_counts}, guessed: {guessed_separator}")  # Feilsøking
        return guessed_separator
        
        
            
    

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
        if self.createColumnCheckBox.isChecked():
            new_col_name = self.newColumnNameLineEdit.text().strip()
            if not new_col_name:  # Sjekk at kolonnenavn ikke er tomt
                print("Kolonnenavn er ikke spesifisert.")
                return
            self.df[new_col_name] = ''  # Opprett en ny kolonne hvis den ikke eksisterer
            target_column = new_col_name
        else:
            target_column = self.extraColumnComboBox.currentText()

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
            QMessageBox.information(self, "Ingen data", "Ingen data å eksportere")
            return

        options = QFileDialog.Options()
        baseFilePath, _ = QFileDialog.getSaveFileName(self, "Lagre fil som", "", "Alle filer (*);;CSV-filer (*.csv);;Excel-filer (*.xlsx)", options=options)
        
        if not baseFilePath:
            return

        exportToCsv = self.exportCsvCheckBox.isChecked()
        exportToExcel = self.exportExcelCheckBox.isChecked()

        if not exportToCsv and not exportToExcel:
            QMessageBox.warning(self, "Eksportformat ikke valgt", "Vennligst velg minst ett eksportformat")
            return

        exportedFiles = []  # En liste for å holde stiene til eksporterte filer

        try:
            if exportToCsv:
                csvFilePath = baseFilePath if baseFilePath.endswith('.csv') else baseFilePath + '.csv'
                self.df.to_csv(csvFilePath, sep=';', encoding='utf-8', index=False)
                exportedFiles.append(csvFilePath)  # Legg til CSV-filbanen til listen

            if exportToExcel:
                if baseFilePath.endswith('.csv'):
                    baseFilePath = filePath[:-4]
                excelFilePath = baseFilePath if baseFilePath.endswith('.xlsx') else baseFilePath + '.xlsx'
                self.df.to_excel(excelFilePath, index=False)
                exportedFiles.append(excelFilePath)  # Legg til Excel-filbanen til listen

            # Sjekk antall eksporterte filer og vis tilsvarende meldingsboks
            if len(exportedFiles) == 2:  # Begge formatene ble eksportert
                QMessageBox.information(self, "Eksport fullført", f"Data eksportert i begge formatene:\nCSV: {exportedFiles[0]}\nExcel: {exportedFiles[1]}")
            elif len(exportedFiles) == 1:  # Kun ett format ble eksportert
                QMessageBox.information(self, "Eksport fullført", f"Data eksportert til: {exportedFiles[0]}")
            else:  # Ingen filer ble eksportert, noe som normalt ikke skulle skje gitt tidligere sjekker
                QMessageBox.warning(self, "Ingen eksport utført", "Ingen data ble eksportert. Vennligst velg et eksportformat og prøv igjen.")

        except Exception as e:
            QMessageBox.critical(self, "Eksportfeil", f"En feil oppstod under eksportering: {str(e)}")

    def convert_and_export_org_csv(self):
        try:
            if self.org_csv_file_path:
                # Les inn filen
                with open(self.org_csv_file_path, 'rb') as file:
                    encoding = chardet.detect(file.read(100000))['encoding']
                df = pd.read_csv(self.org_csv_file_path, sep=';', encoding=encoding, header=None, dtype=str)

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
                    QMessageBox.information(self, "Eksport fullført", f"'{export_file_path}' ble eksportert uten problemer og lagret i '{export_file_path}'")
            else:
                QMessageBox.information(self, "Kildedata mangler", f"Glemte du å importere csv-filen?")
        except Exception as e:
            logging.error(f"En uventet feil oppsto: {e}")
            QMessageBox.warning(self, "Oops!", f"En utventet feil oppsto. Kontakt support med følgende feilkode: {e}")
            logging.error(traceback.format_exc())  # Logger stack trace


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
    
    def clean_enhetskoder(self, df):
        # Finn alle kolonner som starter med "Enhetskode (niv"
        enhetskode_columns = [col for col in df.columns if col.startswith('Enhetskode (niv')]
        
        # Gå gjennom alle funnede kolonner og fjern '.0'
        for col in enhetskode_columns:
            df[col] = df[col].astype(str).str.replace('.0', '', regex=False)
        
        return df

    def generate_columns(self, df, top_level_name, id_prefix):
        # Initialiserer output DataFrame
        output_df = pd.DataFrame(columns=["Name", "ID number", "Description", "Parent", "ID Sort Key"])
        id_counter = 0  # Starter telleren fra 0 for nivå 0 (0 er topp)
        df = self.clean_enhetskoder(df)

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
                    code = str(row[code_col]).replace('.0', '')
                    name = f"{code} - {row[name_col]}"
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