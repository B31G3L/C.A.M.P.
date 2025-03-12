from PyQt5 import QtWidgets, QtCore
import pandas as pd
import os
import re
from datetime import datetime
from io import StringIO

class DropAreaHTML(QtWidgets.QLabel):
    """Drag & Drop Bereich für HTML-Dateien."""
    def __init__(self, file_dropped_callback, parent=None):
        super().__init__(parent)
        self.file_dropped_callback = file_dropped_callback
        self.setAcceptDrops(True)
        self.setText("Ziehe deine HTML-Datei hierher")
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
            if file_ext.endswith('.html') or file_ext.endswith('.xls'):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith('.html') or file_path.lower().endswith('.xls'):
                if self.file_dropped_callback:
                    self.file_dropped_callback(file_path)
            else:
                QtWidgets.QMessageBox.warning(self, "Falscher Dateityp", "Bitte eine HTML- oder XLS-Datei ablegen.")

class ManageDataModal(QtWidgets.QDialog):
    """Dialog zur Verwaltung der Daten aus HTML-Dateien."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Daten verwalten - HTML Import")
        self.setGeometry(300, 200, 700, 500)

        layout = QtWidgets.QVBoxLayout(self)

        self.drop_area_html = DropAreaHTML(self.handle_html_file, self)
        layout.addWidget(self.drop_area_html)

        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Name", "Datum", "Stunden"])
        layout.addWidget(self.table)

        # Hinzufügen eines Layouts für Status und Buttons
        bottom_layout = QtWidgets.QHBoxLayout()

        # Label für die Anzahl der Einträge
        self.entry_count_label = QtWidgets.QLabel("")
        bottom_layout.addWidget(self.entry_count_label)

        self.status_label = QtWidgets.QLabel("")
        bottom_layout.addWidget(self.status_label)

        # Button zum Löschen der Daten
        self.clear_button = QtWidgets.QPushButton("Daten löschen")
        self.clear_button.clicked.connect(self.clear_kapa_data)
        bottom_layout.addWidget(self.clear_button)

        layout.addLayout(bottom_layout)

        self.ensure_data_directory()
        self.load_existing_data()

    def ensure_data_directory(self):
        """Stellt sicher, dass das 'data'-Verzeichnis existiert."""
        data_dir = "data"
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

    def load_existing_data(self):
        """Lädt vorhandene Daten aus kapa_data.csv und zeigt sie an."""
        kapa_file = "data/kapa_data.csv"
        if os.path.exists(kapa_file):
            try:
                df = pd.read_csv(kapa_file, sep=';', header=None, names=["Name", "Datum", "Stunden"], dtype=str)
                df.sort_values(by="Datum", ascending=True, inplace=True)
                self.display_data_in_table(df)
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Fehler", f"Fehler beim Laden von kapa_data.csv:\n{e}")

    def handle_html_file(self, file_path):
        """Verarbeitet eine HTML-ähnliche Datei."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                mime_data = f.read()

            extracted_data = self.extract_data_from_mime(mime_data)
            if extracted_data:
                df = pd.DataFrame(extracted_data, columns=["Name", "Datum", "Stunden"])
                self.append_to_kapa_data(df)
            else:
                QtWidgets.QMessageBox.warning(self, "Fehler", "Keine Daten zum Importieren gefunden.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Fehler", f"Fehler beim Verarbeiten der HTML-Datei:\n{e}")

    def extract_data_from_mime(self, mime_data):
        """Extrahiert Name, Datum und Gesamtstunden aus einer MIME-ähnlichen HTML-Datei."""
        rows = re.findall(r'<tr.*?>(.*?)</tr>', mime_data, re.DOTALL)
        extracted_data = []

        for row in rows:
            cells = re.findall(r'<td.*?>(.*?)</td>', row, re.DOTALL)
            if len(cells) >= 14:  # Stellen Sie sicher, dass genügend Zellen vorhanden sind
                try:
                    # Name extrahieren
                    sixth_value = re.sub(r'<.*?>', '', cells[5]).replace("&#32;", " ")
                    if sixth_value == "Fremdleistung OPS":
                        name = re.sub(r'<.*?>', '', cells[6]).replace("&#32;", " ")
                    else:
                        name = sixth_value

                    # Datum extrahieren
                    date_cell = cells[7]
                    date_value = re.sub(r'<.*?>', '', date_cell).strip()
                    try:
                        excel_date = int(float(date_value.replace(",", ".")))
                        datetime_date = datetime(1899, 12, 30) + datetime.timedelta(days=excel_date)
                        formatted_date = datetime_date.strftime('%d.%m.%Y')
                    except ValueError:
                        formatted_date = date_value

                    # Gesamtstunden extrahieren
                    hours_cell = cells[12]
                    hours_value = re.sub(r'<.*?>', '', hours_cell).strip()
                    try:
                        hours = float(hours_value.replace(",", "."))
                    except ValueError:
                        hours = hours_value

                    extracted_data.append((name, formatted_date, hours))
                except Exception as e:
                    print(f"Fehler beim Verarbeiten einer Zeile: {e}")

        return extracted_data

    def append_to_kapa_data(self, new_data):
        """Fügt neue Daten zu kapa_data.csv hinzu und aktualisiert die Tabelle."""
        kapa_file = "data/kapa_data.csv"
        if os.path.exists(kapa_file):
            
            existing_df = pd.read_csv(kapa_file, sep=';', header=None, names=["Name", "Datum", "Stunden"], dtype=str)
            combined_df = pd.concat([existing_df, new_data], ignore_index=True)
        else:
            combined_df = new_data

        combined_df.sort_values(by="Datum", ascending=True, inplace=True)
        combined_df.to_csv(kapa_file, sep=';', index=False, header=False)
        self.display_data_in_table(combined_df)
        self.status_label.setText("Daten aktualisiert.")

    def display_data_in_table(self, df):
        self.df = df
        self.table.setRowCount(len(df))
        for row_idx, row in df.iterrows():
            self.table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row["Name"])))
            self.table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row["Datum"])))
            self.table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row["Stunden"])))
        self.update_entry_count() # Hier wird die Anzahl der Einträge aktualisiert

    def update_entry_count(self):
        """Aktualisiert die Anzeige der Anzahl der Einträge."""
        if hasattr(self, 'df'):
            count = len(self.df)
            self.entry_count_label.setText(f"Einträge: {count}")
        else:
            self.entry_count_label.setText("Einträge: 0")

    def clear_kapa_data(self):
        kapa_file = "data/kapa_data.csv"
        if os.path.exists(kapa_file):
            with open(kapa_file, 'w'):  # Datei im Schreibmodus öffnen und sofort schließen
                pass  # Inhalt wird gelöscht
            self.table.setRowCount(0)
            self.status_label.setText("Daten gelöscht.")
            self.update_entry_count() # Hier wird die Anzahl der Einträge aktualisiert
        else:
            self.status_label.setText("Keine Daten zum Löschen.")