"""
Modal-Fenster zur Anzeige von CSV-Daten mit Suchfunktion
"""
import os
import csv
import customtkinter as ctk
from typing import List, Dict, Optional, Any

class DataModal(ctk.CTkToplevel):
    """
    Modal-Fenster zur Anzeige von Daten aus einer CSV-Datei mit Suchfunktion
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
        self.title("CSV-Daten Anzeige")
        self.geometry("800x600")
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.csv_path = csv_path
        self.data = []
        self.headers = []
        self.filtered_data = []  # Für die Suchfunktion
        self.search_active = False
        
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
        # Titel-Frame
        self.title_frame = ctk.CTkFrame(self)
        self.title_label = ctk.CTkLabel(
            self.title_frame, 
            text="CSV-Daten",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        
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
        
        # Reset-Button (nur sichtbar wenn Suche aktiv)
        self.reset_button = ctk.CTkButton(
            self.search_frame,
            text="Zurücksetzen",
            command=self._reset_search,
            width=100,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40")
        )
        
        # Info-Label
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
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Frame
        self.title_frame.pack(fill="x", padx=15, pady=(15, 5))
        self.title_label.pack(anchor="w", padx=10, pady=5)
        
        # Such-Frame
        self.search_frame.pack(fill="x", padx=15, pady=5)
        self.search_entry.pack(side="left", padx=(10, 5), pady=5)
        self.search_column_dropdown.pack(side="left", padx=5, pady=5)
        self.search_button.pack(side="left", padx=5, pady=5)
        self.reset_button.pack(side="left", padx=5, pady=5)
        
        # Das Reset-Button ist zunächst ausgeblendet
        if not self.search_active:
            self.reset_button.pack_forget()
        
        # Info-Frame
        self.info_frame.pack(fill="x", padx=15, pady=5)
        self.file_label.pack(anchor="w", padx=10, pady=2)
        self.count_label.pack(anchor="w", padx=10, pady=2)
        
        # Tabellen-Frame (nimmt den meisten Platz ein)
        self.table_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=15, pady=(5, 15))
        self.close_button.pack(side="right", padx=10)
        self.refresh_button.pack(side="right", padx=10)
    
    def _load_data(self):
        """Lädt die Daten aus der CSV-Datei"""
        self.data = []
        self.headers = []
        self.filtered_data = []
        self.search_active = False
        
        # Entferne den Reset-Button, da wir die Suche zurücksetzen
        self.reset_button.pack_forget()
        
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
                
                # Erste Zeile könnte Header sein, wir prüfen das später
                first_row = next(reader, None)
                if not first_row:
                    self._show_error("Die CSV-Datei ist leer.")
                    return
                
                # Restliche Zeilen lesen
                remaining_rows = list(reader)
                
                # Entscheide, ob die erste Zeile Header oder Daten sind
                # Für diesen einfachen Fall nehmen wir an, dass keine Header existieren
                # In einer realen Anwendung könnte man das intelligenter machen
                self.headers = ["Name", "Datum", "Stunden"]  # Annahme für kapa_data.csv
                self.data = [first_row] + remaining_rows
                
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
        
        # Reset-Button anzeigen
        self.reset_button.pack(side="left", padx=5, pady=5)
        
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
        
        # Reset-Button ausblenden
        self.reset_button.pack_forget()
        
        # Tabelle aktualisieren
        self.count_label.configure(text=f"Anzahl Einträge: {len(self.data)}")
        self._display_data(self.data)
    
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