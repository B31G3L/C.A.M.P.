"""
Modal-Fenster zur Anzeige von CSV-Daten mit Such- und Löschfunktion
"""
import os
import csv
import customtkinter as ctk
from typing import List, Dict, Optional, Any
import tkinter as tk
from tkinter import messagebox, filedialog
import shutil

class DataModal(ctk.CTkToplevel):
    """
    Modal-Fenster zur Anzeige von Daten aus einer CSV-Datei mit Such- und Löschfunktion
    """
    def __init__(self, master, csv_path: str = "data/kapa_data.csv"):
        """
        Initialisiert das Modal-Fenster
        
        Args:
            master: Das übergeordnete Widget
            csv_path: Pfad zur CSV-Datei (Standard: "data/kapa_data.csv")
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("Daten")
        self.geometry("800x600")
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.csv_path = csv_path
        self.data = []
        self.headers = []
        self.filtered_data = []  # Für die Suchfunktion
        self.search_active = False
        self.selected_rows = set()  # Für die Auswahl von Zeilen zum Löschen
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Daten laden
        self._load_data()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
    
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        
        
        # Such-Frame
        self.search_frame = ctk.CTkFrame(self)
        
        # Suchfeld
        self.search_entry = ctk.CTkEntry(
            self.search_frame,
            placeholder_text="Suchen...",
            width=200
        )
        # Suchfeld reagiert auf Enter-Taste
        self.search_entry.bind("<Return>", lambda event: self._search())
        
        # Suchspalten-Auswahl
        self.search_column_var = ctk.StringVar(value="Alle")
        self.search_column_dropdown = ctk.CTkOptionMenu(
            self.search_frame,
            values=["Alle", "Name", "Datum", "Stunden"],
            variable=self.search_column_var
        )
        
        # Such-Button
        self.search_button = ctk.CTkButton(
            self.search_frame,
            text="Suchen",
            command=self._search,
            width=80
        )
        
        # Reset-Button für die Suche
        self.reset_button = ctk.CTkButton(
            self.search_frame,
            text="Zurücksetzen",
            command=self._reset_search,
            width=100,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40")
        )
        
        # Info-Frame
        self.info_frame = ctk.CTkFrame(self)
        self.file_label = ctk.CTkLabel(
            self.info_frame, 
            text=f"Datei: {self.csv_path}",
            anchor="w"
        )
        self.count_label = ctk.CTkLabel(
            self.info_frame, 
            text="Anzahl Einträge: 0",
            anchor="w"
        )
        
        # Tabellen-Frame
        self.table_frame = ctk.CTkScrollableFrame(self)
        
        # Lösch-Info-Frame (neu, unterhalb der Tabelle)
        self.delete_info_frame = ctk.CTkFrame(self)
        
        # Lösch-Info-Label (jetzt im neuen Frame)
        self.delete_info_label = ctk.CTkLabel(
            self.delete_info_frame,
            text="Keine Einträge ausgewählt",
            anchor="w",
            text_color=("gray60", "gray70")
        )
        
        # Button-Frame
        self.button_frame = ctk.CTkFrame(self)
        self.close_button = ctk.CTkButton(
            self.button_frame, 
            text="Schließen",
            command=self.destroy
        )
        self.refresh_button = ctk.CTkButton(
            self.button_frame, 
            text="Aktualisieren",
            command=self._load_data
        )
        
        # Lösch-Button
        self.delete_button = ctk.CTkButton(
            self.button_frame,
            text="Löschen",
            command=self._delete_selected_rows,
            fg_color=("red3", "red4"),
            hover_color=("red4", "red3"),
            state="disabled"  # Anfangs deaktiviert
        )

        # Neuer Button: Alles löschen
        self.delete_all_button = ctk.CTkButton(
            self.button_frame,
            text="Alles löschen",
            command=self._delete_all_data,
            fg_color=("red3", "red4"),
            hover_color=("red4", "red3"),
            width=120
        )

        # Import/Export Frame
        self.io_frame = ctk.CTkFrame(self)

        # Import-Button
        self.import_button = ctk.CTkButton(
            self.io_frame,
            text="Import CSV",
            command=self._import_data,
            width=120
        )

        # Export-Button
        self.export_button = ctk.CTkButton(
            self.io_frame,
            text="Export CSV",
            command=self._export_data,
            width=120
        )

    # 2. Jetzt passen wir das Layout an, um das delete_info_frame unter der Tabelle zu platzieren
    # Ändere diese Methode:

    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        
        # Info-Frame
        self.file_label.pack(anchor="w", padx=10, pady=2)
        self.count_label.pack(anchor="w", padx=10, pady=2)
        
        # Such-Frame
        self.search_frame.pack(fill="x", padx=15, pady=5)
        self.search_entry.pack(side="left", padx=(10, 5), pady=5)
        self.search_column_dropdown.pack(side="left", padx=5, pady=5)
        self.search_button.pack(side="left", padx=5, pady=5)
        self.reset_button.pack(side="left", padx=5, pady=5)

        # Import/Export-Frame
        self.io_frame.pack(fill="x", padx=15, pady=5)
        self.import_button.pack(side="left", padx=(10, 5), pady=5)
        self.export_button.pack(side="left", padx=5, pady=5)
        
        # Tabellen-Frame (nimmt den meisten Platz ein)
        self.table_frame.pack(fill="both", expand=True, padx=15, pady=10)
        self.info_frame.pack(fill="x", padx=15, pady=5)

        # Lösch-Info-Frame (jetzt unter der Tabelle)
        self.delete_info_frame.pack(fill="x", padx=15, pady=(0, 5))
        self.delete_info_label.pack(fill="x", padx=10, pady=2)
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=15, pady=(5, 15))
        self.close_button.pack(side="right", padx=10)
        self.refresh_button.pack(side="right", padx=10)
        self.delete_button.pack(side="left", padx=10)
        self.delete_all_button.pack(side="left", padx=10)
    
    def _load_data(self):
        """Lädt die Daten aus der CSV-Datei"""
        self.data = []
        self.headers = []
        self.filtered_data = []
        self.search_active = False
        self.selected_rows = set()
        
        # Lösch-Button deaktivieren, da keine Zeilen ausgewählt sind
        self.delete_button.configure(state="disabled")
        self.delete_info_label.configure(text="Keine Einträge ausgewählt")
        
        # Prüfe, ob die Datei existiert
        if not os.path.exists(self.csv_path):
            self._show_error(f"Die Datei {self.csv_path} existiert nicht.")
            return
        
        try:
            # Datei öffnen und lesen
            with open(self.csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                # Versuche den Delimiter zu erkennen
                try:
                    dialect = csv.Sniffer().sniff(csvfile.read(1024))
                    csvfile.seek(0)
                    reader = csv.reader(csvfile, dialect)
                except:
                    # Fallback, wenn Sniffer fehlschlägt
                    csvfile.seek(0)
                    reader = csv.reader(csvfile, delimiter=';')
                
                # Alle Zeilen lesen
                all_rows = list(reader)
                if not all_rows:
                    self._show_error("Die CSV-Datei ist leer.")
                    return
                
                # Prüfe, ob die erste Zeile ein Header ist
                first_row = all_rows[0]
                if first_row and len(first_row) >= 3 and first_row[0] == "ID" and first_row[1] == "DATUM":
                    # Es ist ein Header
                    self.headers = first_row
                    self.data = all_rows[1:]  # Alle Zeilen außer dem Header
                else:
                    # Keine Header, verwende Standardnamen
                    self.headers = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]
                    self.data = all_rows
                
                # Aktualisiere die Suchspalten-Dropdown mit den Header-Namen
                self.search_column_dropdown.configure(
                    values=["Alle"] + self.headers
                )
                
                # GUI aktualisieren
                self.count_label.configure(text=f"Anzahl Einträge: {len(self.data)}")
                self._display_data(self.data)
        
        except Exception as e:
            self._show_error(f"Fehler beim Lesen der CSV-Datei: {e}")
    
    def _search(self):
        """Führt eine Suche basierend auf dem Suchbegriff und der ausgewählten Spalte durch"""
        search_term = self.search_entry.get().strip().lower()
        
        if not search_term:
            # Wenn der Suchbegriff leer ist, zeige alle Daten an
            if self.search_active:
                self._reset_search()
            return
        
        # Spalte für die Suche auswählen
        search_column = self.search_column_var.get()
        
        # Filtern der Daten
        self.filtered_data = []
        for row in self.data:
            if search_column == "Alle":
                # Suche in allen Spalten
                if any(search_term in str(cell).lower() for cell in row):
                    self.filtered_data.append(row)
            else:
                # Suche in einer bestimmten Spalte
                col_idx = self.headers.index(search_column)
                if col_idx < len(row) and search_term in str(row[col_idx]).lower():
                    self.filtered_data.append(row)
        
        # Suche ist jetzt aktiv
        self.search_active = True
        
        # Reset-Button ist bereits sichtbar
        
        # Tabelle aktualisieren
        if self.filtered_data:
            self.count_label.configure(
                text=f"Gefundene Einträge: {len(self.filtered_data)} von {len(self.data)}"
            )
            self._display_data(self.filtered_data)
        else:
            # Keine Ergebnisse gefunden
            self._show_message("Keine Einträge gefunden für: " + search_term)
    
    def _reset_search(self):
        """Setzt die Suche zurück und zeigt alle Daten an"""
        self.search_active = False
        self.filtered_data = []
        self.search_entry.delete(0, 'end')  # Suchfeld leeren
        
        # Reset-Button bleibt sichtbar
        # (Keine Änderung am Button-Status erforderlich)
        
        # Tabelle aktualisieren
        self.count_label.configure(text=f"Anzahl Einträge: {len(self.data)}")
        self._display_data(self.data)
    
    def _toggle_row_selection(self, row_idx, row_data, row_frame):
        """
        Wählt eine Zeile aus oder hebt die Auswahl auf
        
        Args:
            row_idx: Index der Zeile
            row_data: Daten der Zeile
            row_frame: Das Frame-Widget der Zeile
        """
        # Bestimme, ob wir gefilterte oder alle Daten anzeigen
        display_data = self.filtered_data if self.search_active else self.data
        
        # Eindeutigen Schlüssel für die Zeile erstellen
        # Wir verwenden einen Tupel aus den Zeilenwerten als Schlüssel
        row_key = tuple(row_data)
        
        # Auswahl umschalten
        if row_key in self.selected_rows:
            self.selected_rows.remove(row_key)
            # Hintergrundfarbe zurücksetzen (abhängig vom Zeilenindex)
            bg_color = "transparent" if row_idx % 2 == 0 else ("gray90", "gray20")
            row_frame.configure(fg_color=bg_color)
        else:
            self.selected_rows.add(row_key)
            # Hintergrundfarbe für ausgewählte Zeile setzen
            row_frame.configure(fg_color=("light blue", "dark blue"))
        
        # Lösch-Button-Status aktualisieren
        if self.selected_rows:
            self.delete_button.configure(state="normal")
            self.delete_info_label.configure(
                text=f"{len(self.selected_rows)} Einträge zum Löschen ausgewählt"
            )
        else:
            self.delete_button.configure(state="disabled")
            self.delete_info_label.configure(text="Keine Einträge ausgewählt")
    
    def _delete_selected_rows(self):
        """Löscht die ausgewählten Zeilen aus der CSV-Datei"""
        if not self.selected_rows:
            return
        
        # Bestätigungsdialog anzeigen
        confirm = messagebox.askyesno(
            "Einträge löschen",
            f"Möchten Sie wirklich {len(self.selected_rows)} Einträge löschen?",
            parent=self
        )
        
        if not confirm:
            return
        
        try:
            # Neue Liste ohne die ausgewählten Zeilen erstellen
            new_data = [row for row in self.data if tuple(row) not in self.selected_rows]
            
            # Datei sichern (Backup)
            backup_path = self.csv_path + ".bak"
            try:
                import shutil
                shutil.copy2(self.csv_path, backup_path)
            except Exception as e:
                print(f"Warnung: Backup konnte nicht erstellt werden: {e}")
            
            # Datei mit neuen Daten schreiben
            with open(self.csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, delimiter=';')
                for row in new_data:
                    writer.writerow(row)
            
            # Erfolgsmeldung
            messagebox.showinfo(
                "Einträge gelöscht",
                f"{len(self.selected_rows)} Einträge wurden erfolgreich gelöscht.",
                parent=self
            )
            
            # Daten neu laden
            self._load_data()
            
        except Exception as e:
            messagebox.showerror(
                "Fehler beim Löschen",
                f"Fehler beim Löschen der Einträge: {e}",
                parent=self
            )
            
            # Versuche, das Backup wiederherzustellen
            if os.path.exists(backup_path):
                try:
                    import shutil
                    shutil.copy2(backup_path, self.csv_path)
                    messagebox.showinfo(
                        "Wiederherstellung",
                        "Die Datei wurde aus dem Backup wiederhergestellt.",
                        parent=self
                    )
                except Exception as restore_error:
                    messagebox.showerror(
                        "Fehler bei der Wiederherstellung",
                        f"Fehler bei der Wiederherstellung aus dem Backup: {restore_error}",
                        parent=self
                    )
    
    def _display_data(self, data_to_display):
        """
        Zeigt die übergebenen Daten in einer Tabelle an
        
        Args:
            data_to_display: Die anzuzeigenden Daten
        """
        # Vorhandene Tabelle löschen
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        # Keine Daten zu zeigen
        if not data_to_display:
            no_data_label = ctk.CTkLabel(
                self.table_frame, 
                text="Keine Daten verfügbar",
                font=ctk.CTkFont(size=14)
            )
            no_data_label.pack(expand=True, pady=50)
            return
        
        # Erstelle Tabellen-Container
        table_container = ctk.CTkFrame(self.table_frame)
        table_container.pack(fill="both", expand=True)
        
        # Spaltenbreiten festlegen
        col_width = 200  # Standard-Spaltenbreite
        
        # Kopfzeile erstellen
        header_frame = ctk.CTkFrame(table_container)
        header_frame.pack(fill="x", pady=(0, 2))
        
        # Ermittle Anzahl der Spalten (entweder aus Headers oder aus erster Datenzeile)
        num_cols = len(self.headers) if self.headers else len(data_to_display[0])
        
        # Header-Labels erstellen, wenn vorhanden
        if self.headers:
            for i, header in enumerate(self.headers):
                header_label = ctk.CTkLabel(
                    header_frame, 
                    text=header,
                    font=ctk.CTkFont(weight="bold"),
                    width=col_width
                )
                header_label.grid(row=0, column=i, padx=2, pady=2, sticky="w")
        
        # Datenzeilen erstellen
        for row_idx, row in enumerate(data_to_display):
            # Einige Zeilen überspringen, wenn zu viele (z.B. nur die ersten 100 anzeigen)
            if row_idx >= 100:
                more_label = ctk.CTkLabel(
                    table_container, 
                    text=f"... und {len(data_to_display) - 100} weitere Einträge",
                    font=ctk.CTkFont(size=12, slant="italic")
                )
                more_label.pack(anchor="w", padx=10, pady=5)
                break
            
            # Zeilen-Frame erstellen
            row_frame = ctk.CTkFrame(table_container)
            row_frame.pack(fill="x", pady=1)
            
            # Eintragsfarbe alternieren für bessere Lesbarkeit
            bg_color = "transparent" if row_idx % 2 == 0 else ("gray90", "gray20")
            row_frame.configure(fg_color=bg_color)
            
            # Zeile auswählbar machen
            row_frame.bind("<Button-1>", lambda e, idx=row_idx, r=row, f=row_frame: 
                          self._toggle_row_selection(idx, r, f))
            
            # Prüfen, ob die Zeile bereits ausgewählt ist
            if tuple(row) in self.selected_rows:
                row_frame.configure(fg_color=("light blue", "dark blue"))
            
            # Suchbegriff für Hervorhebung
            search_term = self.search_entry.get().strip().lower() if self.search_active else ""
            
            # Zellen erstellen
            for col_idx, cell_value in enumerate(row[:num_cols]):  # Beschränke auf die Anzahl der Spalten
                cell_text = str(cell_value)
                
                # Hervorhebung des Suchbegriffs, falls aktiv
                if self.search_active and search_term and search_term in cell_text.lower():
                    # Spezielle Formatierung für Zellen, die den Suchbegriff enthalten
                    cell_label = ctk.CTkLabel(
                        row_frame, 
                        text=cell_text,
                        anchor="w",
                        width=col_width,
                        fg_color=("light green", "dark green"),
                        corner_radius=4,
                        text_color=("black", "white")
                    )
                else:
                    # Normale Zelle
                    cell_label = ctk.CTkLabel(
                        row_frame, 
                        text=cell_text,
                        anchor="w",
                        width=col_width
                    )
                
                # Zelle ebenfalls klickbar machen
                cell_label.bind("<Button-1>", lambda e, idx=row_idx, r=row, f=row_frame: 
                               self._toggle_row_selection(idx, r, f))
                
                cell_label.grid(row=0, column=col_idx, padx=2, pady=1, sticky="w")
    
    def _show_error(self, message: str):
        """
        Zeigt eine Fehlermeldung im Modal an
        
        Args:
            message: Die anzuzeigende Fehlermeldung
        """
        # Vorhandene Tabelle löschen
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        # Fehlermeldung anzeigen
        error_frame = ctk.CTkFrame(self.table_frame)
        error_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        error_label = ctk.CTkLabel(
            error_frame, 
            text=message,
            font=ctk.CTkFont(size=14),
            text_color=("red", "#ff6b6b")  # Rot in Light/Dark Mode
        )
        error_label.pack(expand=True, pady=30)
        
        # Info-Label aktualisieren
        self.count_label.configure(text="Anzahl Einträge: 0")
    
    def _show_message(self, message: str):
        """
        Zeigt eine Nachricht im Modal an
        
        Args:
            message: Die anzuzeigende Nachricht
        """
        # Vorhandene Tabelle löschen
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        # Nachricht anzeigen
        message_frame = ctk.CTkFrame(self.table_frame)
        message_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        message_label = ctk.CTkLabel(
            message_frame, 
            text=message,
            font=ctk.CTkFont(size=14)
        )
        message_label.pack(expand=True, pady=30)
    
    def center_window(self):
        """Zentriert das Fenster auf dem Bildschirm"""
        self.update_idletasks()
        
        # Fenstergröße
        width = self.winfo_width()
        height = self.winfo_height()
        
        # Bildschirmgröße
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # Position berechnen
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Fenster positionieren
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _delete_all_data(self):
        """Löscht alle Daten aus der CSV-Datei mit Sicherheitsabfrage, behält aber den Header"""
        # Prüfen, ob die Datei existiert
        if not os.path.exists(self.csv_path):
            messagebox.showerror(
                "Datei nicht gefunden",
                f"Die Datei {self.csv_path} existiert nicht.",
                parent=self
            )
            return

        # Bestätigungsdialog anzeigen
        confirm = messagebox.askyesno(
            "Alle Daten löschen",
            "Möchten Sie wirklich ALLE Einträge aus der Datei löschen?\n\nDiese Aktion kann nicht rückgängig gemacht werden!",
            parent=self
        )

        if not confirm:
            return

        try:
            # Datei sichern (Backup)
            backup_path = self.csv_path + ".bak"
            try:
                shutil.copy2(self.csv_path, backup_path)
            except Exception as e:
                print(f"Warnung: Backup konnte nicht erstellt werden: {e}")

            # Header definieren
            header = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]

            # Datei mit nur dem Header erstellen
            with open(self.csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, delimiter=';')
                writer.writerow(header)

            # Erfolgsmeldung
            messagebox.showinfo(
                "Daten gelöscht",
                "Alle Einträge wurden erfolgreich gelöscht.",
                parent=self
            )

            # Daten neu laden
            self._load_data()

        except Exception as e:
            messagebox.showerror(
                "Fehler beim Löschen",
                f"Fehler beim Löschen aller Einträge: {e}",
                parent=self
            )

    def _export_data(self):
        """Exportiert die CSV-Datei an einen vom Benutzer gewählten Ort"""
        # Prüfen, ob die Datei existiert
        if not os.path.exists(self.csv_path):
            messagebox.showerror(
                "Datei nicht gefunden",
                f"Die Datei {self.csv_path} existiert nicht.",
                parent=self
            )
            return

        # Datei-Dialog öffnen
        export_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")],
            title="CSV-Datei exportieren",
            initialfile=os.path.basename(self.csv_path)
        )

        if not export_path:
            return  # Benutzer hat abgebrochen

        try:
            # Datei kopieren
            shutil.copy2(self.csv_path, export_path)

            # Erfolgsmeldung
            messagebox.showinfo(
                "Export erfolgreich",
                f"Die Daten wurden erfolgreich nach {export_path} exportiert.",
                parent=self
            )

        except Exception as e:
            messagebox.showerror(
                "Export fehlgeschlagen",
                f"Fehler beim Exportieren der Datei: {e}",
                parent=self
            )

    def _import_data(self):
        """Importiert eine CSV-Datei und fragt, ob bestehende Daten überschrieben werden sollen"""
        # Datei-Dialog öffnen
        import_path = filedialog.askopenfilename(
            filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")],
            title="CSV-Datei importieren"
        )

        if not import_path:
            return  # Benutzer hat abgebrochen

        # Prüfen, ob die Datei existiert
        if not os.path.exists(import_path):
            messagebox.showerror(
                "Datei nicht gefunden",
                f"Die Datei {import_path} existiert nicht.",
                parent=self
            )
            return

        # Frage, ob bestehende Daten überschrieben werden sollen
        if os.path.exists(self.csv_path):
            overwrite = messagebox.askyesno(
                "Daten überschreiben",
                "Sollen die bestehenden Daten überschrieben werden?\n\n" +
                "Ja: Alle bestehenden Daten werden ersetzt.\n" +
                "Nein: Die importierten Daten werden an die bestehenden angehängt.",
                parent=self
            )
        else:
            overwrite = True  # Wenn keine Datei existiert, immer überschreiben

        try:
            # Bestehende Datei sichern (Backup)
            if os.path.exists(self.csv_path):
                backup_path = self.csv_path + ".bak"
                try:
                    shutil.copy2(self.csv_path, backup_path)
                except Exception as e:
                    print(f"Warnung: Backup konnte nicht erstellt werden: {e}")

            # Importierte Daten laden
            with open(import_path, 'r', newline='', encoding='utf-8') as importfile:
                try:
                    dialect = csv.Sniffer().sniff(importfile.read(1024))
                    importfile.seek(0)
                    reader = csv.reader(importfile, dialect)
                except:
                    importfile.seek(0)
                    reader = csv.reader(importfile, delimiter=';')
                
                imported_data = list(reader)

            if overwrite:
                # Bestehende Daten überschreiben
                with open(self.csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile, delimiter=';')
                    for row in imported_data:
                        writer.writerow(row)
                
                message = f"{len(imported_data)} Einträge wurden importiert und bestehende Daten überschrieben."
            else:
                # Bestehende Daten laden
                existing_data = []
                if os.path.exists(self.csv_path):
                    with open(self.csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                        try:
                            dialect = csv.Sniffer().sniff(csvfile.read(1024))
                            csvfile.seek(0)
                            reader = csv.reader(csvfile, dialect)
                        except:
                            csvfile.seek(0)
                            reader = csv.reader(csvfile, delimiter=';')
                        
                        existing_data = list(reader)

                # Anhängen der neuen Daten
                with open(self.csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile, delimiter=';')
                    for row in existing_data:
                        writer.writerow(row)
                    for row in imported_data:
                        writer.writerow(row)
                
                message = f"{len(imported_data)} Einträge wurden importiert und an bestehende Daten angehängt."

            # Erfolgsmeldung
            messagebox.showinfo(
                "Import erfolgreich",
                message,
                parent=self
            )

            # Daten neu laden
            self._load_data()

        except Exception as e:
            messagebox.showerror(
                "Import fehlgeschlagen",
                f"Fehler beim Importieren der Datei: {e}",
                parent=self
            )

            # Versuche, das Backup wiederherzustellen
            if os.path.exists(backup_path):
                try:
                    shutil.copy2(backup_path, self.csv_path)
                    messagebox.showinfo(
                        "Wiederherstellung",
                        "Die Datei wurde aus dem Backup wiederhergestellt.",
                        parent=self
                    )
                except Exception as restore_error:
                    messagebox.showerror(
                        "Fehler bei der Wiederherstellung",
                        f"Fehler bei der Wiederherstellung aus dem Backup: {restore_error}",
                        parent=self
                    )