from PyQt5 import QtWidgets, QtCore
import pandas as pd
import os

class DropArea(QtWidgets.QLabel):
    """Drag & Drop Bereich für CSV- und XLS-Dateien."""
    def __init__(self, file_dropped_callback, parent=None):
        super().__init__(parent)
        self.file_dropped_callback = file_dropped_callback
        self.setAcceptDrops(True)
        self.setText("Ziehe deine CSV- oder XLS-Datei hierher")
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                padding: 40px;
                font-size: 14pt;
                color: #555;
            }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            file_ext = urls[0].toLocalFile().lower()
            if file_ext.endswith('.csv') or file_ext.endswith('.xls') or file_ext.endswith('.xlsx'):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith('.csv') or file_path.lower().endswith('.xls') or file_path.lower().endswith('.xlsx'):
                if self.file_dropped_callback:
                    self.file_dropped_callback(file_path)
            else:
                QtWidgets.QMessageBox.warning(self, "Falscher Dateityp", "Bitte eine CSV- oder XLS-Datei ablegen.")

class ManageDataModal(QtWidgets.QDialog):
    """Dialog zur Verwaltung der CSV- und XLS-Daten."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Daten verwalten - CSV & Excel")
        self.setGeometry(300, 200, 700, 500)

        layout = QtWidgets.QVBoxLayout(self)

        self.drop_area = DropArea(self.handle_file, self)
        layout.addWidget(self.drop_area)

        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["ID", "Datum", "Stunden"])
        layout.addWidget(self.table)

        self.status_label = QtWidgets.QLabel("")
        layout.addWidget(self.status_label)

        self.load_existing_data()

    def load_existing_data(self):
        """Lädt vorhandene Daten aus kapa_data.csv und zeigt sie an."""
        kapa_file = "data/kapa_data.csv"
        if os.path.exists(kapa_file):
            df = pd.read_csv(kapa_file, sep=';', header=None, names=["ID", "Datum", "Stunden"], dtype=str)
            df.sort_values(by="Datum", ascending=True, inplace=True)
            self.display_data_in_table(df)

    def handle_file(self, file_path):
        """Verarbeitet eine Datei basierend auf dem Dateityp."""
        if file_path.endswith(".csv"):
            self.process_csv(file_path)
        elif file_path.endswith(".xls") or file_path.endswith(".xlsx"):
            self.process_xls(file_path)

    def process_csv(self, file_path):
        """Verarbeitet eine CSV-Datei mit ID-Eingabe."""
        id_input, ok = QtWidgets.QInputDialog.getText(self, "Benutzer-ID eingeben", "Bitte die Benutzer-ID eingeben:")
        if not ok or not id_input.strip():
            return

        try:
            df = pd.read_csv(file_path, sep=';', header=None, dtype=str)
            if df.shape[1] < 13:
                QtWidgets.QMessageBox.warning(self, "Fehler", "Die CSV-Datei hat zu wenige Spalten.")
                return

            df["ID"] = id_input
            extracted_data = df[["ID", 7, 12]].copy()
            extracted_data.columns = ["ID", "Datum", "Stunden"]
            extracted_data["Stunden"] = pd.to_numeric(extracted_data["Stunden"], errors='coerce').fillna(0)

            self.append_to_kapa_data(extracted_data)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Fehler", f"Fehler beim Verarbeiten der CSV-Datei:\n{e}")

    def process_xls(self, file_path):
        """Verarbeitet eine XLS-Datei."""
        try:
            df = pd.read_excel(file_path, dtype=str)
            required_columns = ["ID", "Datum", "Stunden"]
            if not all(col in df.columns for col in required_columns):
                QtWidgets.QMessageBox.warning(self, "Fehler", "Die Excel-Datei enthält nicht alle erforderlichen Spalten.")
                return

            df["Stunden"] = pd.to_numeric(df["Stunden"], errors='coerce').fillna(0)
            self.append_to_kapa_data(df[required_columns])
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Fehler", f"Fehler beim Verarbeiten der Excel-Datei:\n{e}")

    def append_to_kapa_data(self, new_data):
        """Fügt neue Daten zu kapa_data.csv hinzu und aktualisiert die Tabelle."""
        kapa_file = "data/kapa_data.csv"
        if os.path.exists(kapa_file):
            existing_df = pd.read_csv(kapa_file, sep=';', header=None, names=["ID", "Datum", "Stunden"], dtype=str)
            combined_df = pd.concat([existing_df, new_data], ignore_index=True)
        else:
            combined_df = new_data

        combined_df.sort_values(by="Datum", ascending=True, inplace=True)
        combined_df.to_csv(kapa_file, sep=';', index=False, header=False)
        self.display_data_in_table(combined_df)
        self.status_label.setText("Daten aktualisiert.")

    def display_data_in_table(self, df):
        """Zeigt DataFrame-Daten in der Tabelle an."""
        self.df = df
        self.table.setRowCount(len(df))
        for row_idx, row in df.iterrows():
            self.table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row["ID"])))
            self.table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row["Datum"])))
            self.table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row["Stunden"])))
