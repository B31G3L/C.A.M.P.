"""
Modal-Fenster zum Importieren von Daten mit Tabs für BW Tool Daten und Projektdaten
"""
import os
import customtkinter as ctk
from typing import Dict, Callable, Optional, List
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import json
import shutil
import datetime
import pandas as pd
import re

class DragDropFrame(ctk.CTkFrame):
    """
    Ein Frame, das Drag & Drop von Dateien ermöglicht
    """
    def __init__(self, master, callback, accepted_files=None, **kwargs):
        """
        Initialisiert den Drag & Drop Frame
        
        Args:
            master: Das übergeordnete Widget
            callback: Funktion, die aufgerufen wird, wenn eine Datei fallen gelassen wird
            accepted_files: Liste der akzeptierten Dateierweiterungen (z.B. ['.xls', '.xlsx'])
        """
        super().__init__(master, **kwargs)
        
        self.callback = callback
        self.accepted_files = accepted_files or ['.xls', '.xlsx']
        
        # Aussehen des Frames
        self.configure(
            fg_color=("gray90", "gray20"),
            corner_radius=10,
            border_width=2,
            border_color=("gray70", "gray40")
        )
        
        # Text und Icon für den Drop-Bereich
        self.text_label = ctk.CTkLabel(
            self,
            text="Datei hier hineinziehen oder klicken zum Auswählen",
            font=ctk.CTkFont(size=14)
        )
        self.text_label.pack(expand=True, pady=(20, 10))
        
        self.info_label = ctk.CTkLabel(
            self,
            text=f"Akzeptierte Dateitypen: {', '.join(self.accepted_files)}",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray70")
        )
        self.info_label.pack(expand=True, pady=(0, 20))
        
        # Erhöhe den Frame so, dass er gut als Drop-Ziel funktioniert
        self.configure(height=150)
        
        # Ereignisbehandlung für Mausklicks
        self.bind("<Button-1>", self._on_click)
        self.text_label.bind("<Button-1>", self._on_click)
        self.info_label.bind("<Button-1>", self._on_click)
        
        # Richte Drag & Drop ein für Windows (TkDND)
        self._setup_tkdnd()
    
    def _setup_tkdnd(self):
        """Richtet TkDND für Drag & Drop-Funktionalität ein, falls verfügbar"""
        try:
            self.winfo_toplevel().tk.call('package', 'require', 'tkdnd')
            self.winfo_toplevel().tk.call('tkdnd::drop_target', 'register', self._get_widget_path(), self.accepted_files)
            self.bind('<<Drop>>', self._on_drop)
            
            # Visuelle Rückmeldung beim Überfahren
            self.bind('<<DropEnter>>', lambda e: self.configure(border_color=("blue", "light blue")))
            self.bind('<<DropLeave>>', lambda e: self.configure(border_color=("gray70", "gray40")))
        except tk.TclError:
            # TkDND ist nicht verfügbar, füge eine Info hinzu
            self.info_label.configure(text="Drag & Drop nicht verfügbar, bitte klicken zum Auswählen")
    
    def _get_widget_path(self):
        """Gibt den TK-Widget-Pfad für TkDND zurück"""
        return self.winfo_toplevel().register(self)
    
    def _on_click(self, event):
        """Wird aufgerufen, wenn auf den Frame geklickt wird"""
        filetypes = [("Excel-Dateien", "*.xls *.xlsx"), ("Alle Dateien", "*.*")]
        filename = filedialog.askopenfilename(
            title="BW Tool Datei auswählen",
            filetypes=filetypes
        )
        
        if filename:
            self.callback(filename)
    
    def _on_drop(self, event):
        """Wird aufgerufen, wenn eine Datei auf den Frame gezogen wird"""
        self.configure(border_color=("gray70", "gray40"))  # Zurücksetzen der Rahmenfarbe
        
        # Prüfe, ob die Datei einen der akzeptierten Dateitypen hat
        filename = event.data
        _, file_extension = os.path.splitext(filename)
        
        if file_extension.lower() in self.accepted_files:
            self.callback(filename)
        else:
            messagebox.showerror(
                "Ungültiger Dateityp",
                f"Die Datei {os.path.basename(filename)} hat einen ungültigen Dateityp. "
                f"Akzeptierte Typen: {', '.join(self.accepted_files)}",
                parent=self.winfo_toplevel()
            )

class ImportModal(ctk.CTkToplevel):
    """
    Modal-Fenster zum Importieren von Daten mit Tabs für BW Tool Daten und Projektdaten
    """
    def __init__(self, master, callbacks: Optional[Dict[str, Callable]] = None):
        """
        Initialisiert das Modal-Fenster
        
        Args:
            master: Das übergeordnete Widget
            callbacks: Dictionary mit Callback-Funktionen
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("Daten importieren")
        self.geometry("700x550")
        self.minsize(600, 450)  # Minimale Fenstergröße festlegen
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.callbacks = callbacks or {}
        self.selected_file = None
        self.import_status = None
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
        
        # Stelle sicher, dass das Fenster tatsächlich angezeigt wird
        self.attributes('-topmost', True)
        self.after(200, lambda: self.attributes('-topmost', False))
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        # Titel-Frame
        self.title_frame = ctk.CTkFrame(self)
        self.title_label = ctk.CTkLabel(
            self.title_frame, 
            text="Daten importieren",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        
        # Status-Label
        self.status_frame = ctk.CTkFrame(self)
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray70")
        )
        
        # Tab-View für die verschiedenen Import-Optionen
        self.tab_view = ctk.CTkTabview(self)
        
        # Tabs erstellen
        self.bw_tool_tab = self.tab_view.add("BW Tool Daten")
        self.projekt_tab = self.tab_view.add("Projektdaten")
        
        # Inhalte für BW Tool Daten-Tab
        self._create_bw_tool_tab()
        
        # Inhalte für Projektdaten-Tab
        self._create_projekt_tab()
        
        # Button-Frame (unten)
        self.button_frame = ctk.CTkFrame(self)
        self.close_button = ctk.CTkButton(
            self.button_frame, 
            text="Schließen",
            command=self.destroy,
            width=120
        )
    
    def _create_bw_tool_tab(self):
        """Erstellt die Inhalte für den BW Tool Daten-Tab"""
        # Container-Frame für den Tab
        container = ctk.CTkFrame(self.bw_tool_tab)
        
        # Info-Text
        info_text = (
            "Importieren Sie BW Tool Daten aus einer Excel-Datei (.xls, .xlsx).\n\n"
            "Die Daten werden aus der Excel-Datei extrahiert und in folgendem Format gespeichert:\n"
            "MITARBEITER_ID;DATUM;STUNDEN\n\n"
            "Es werden nur Einträge berücksichtigt, die in Spalte 11 'FCAST', 'Forecast' oder 'FORECAST' enthalten."
        )
        
        info_label = ctk.CTkLabel(
            container,
            text=info_text,
            justify="left",
            anchor="w",
            wraplength=600
        )
        
        # Drag & Drop Frame
        self.drop_frame = DragDropFrame(
            container,
            callback=self._on_file_selected,
            accepted_files=['.xls', '.xlsx']
        )
        
        # Dateiauswahl-Frame (für Anzeige des ausgewählten Dateinamens)
        file_frame = ctk.CTkFrame(container)
        
        # Label für ausgewählte Datei
        file_label = ctk.CTkLabel(
            file_frame,
            text="Ausgewählte Datei:",
            font=ctk.CTkFont(weight="bold"),
            width=120
        )
        
        # Anzeige des Dateinamens
        self.file_name_label = ctk.CTkLabel(
            file_frame,
            text="Keine Datei ausgewählt",
            anchor="w"
        )
        
        # Optionen-Frame
        options_frame = ctk.CTkFrame(container)
        
        # Optionen für den Import
        self.bw_tool_overwrite_var = ctk.BooleanVar(value=True)
        self.bw_tool_overwrite_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Bestehende Einträge überschreiben (gleiche ID und Datum)",
            variable=self.bw_tool_overwrite_var
        )
        
        self.bw_tool_backup_var = ctk.BooleanVar(value=True)
        self.bw_tool_backup_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Backup erstellen",
            variable=self.bw_tool_backup_var
        )
        
        # Import-Button
        self.bw_tool_import_button = ctk.CTkButton(
            container,
            text="BW Tool Daten importieren",
            command=self._import_bw_tool_data,
            font=ctk.CTkFont(weight="bold"),
            height=40,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3"),
            state="disabled"  # Initial deaktiviert
        )
        
        # Layout für den BW Tool Daten-Tab
        container.pack(fill="both", expand=True, padx=20, pady=20)
        info_label.pack(anchor="w", pady=(0, 15))
        
        # Drag & Drop Frame
        self.drop_frame.pack(fill="x", pady=(0, 15))
        
        # Dateiauswahl-Frame
        file_frame.pack(fill="x", pady=(0, 15))
        file_label.pack(side="left")
        self.file_name_label.pack(side="left", padx=10, fill="x", expand=True)
        
        # Optionen-Frame
        options_frame.pack(fill="x", pady=(0, 20))
        self.bw_tool_overwrite_checkbox.pack(anchor="w", pady=5)
        self.bw_tool_backup_checkbox.pack(anchor="w", pady=5)
        
        # Import-Button
        self.bw_tool_import_button.pack(fill="x", pady=(10, 0))
    
    def _create_projekt_tab(self):
        """Erstellt die Inhalte für den Projektdaten-Tab"""
        # Container-Frame für den Tab
        container = ctk.CTkFrame(self.projekt_tab)
        
        # Info-Text
        info_text = (
            "Importieren Sie Projektdaten aus einer JSON-Datei.\n\n"
            "Die JSON-Datei sollte eine Struktur mit Projekten, Teilnehmern und Sprints enthalten.\n\n"
            "Beispiel:\n"
            '{\n  "projekte": [\n    {\n      "name": "Projektname",\n      "teilnehmer": [...],\n      "sprints": [...]\n    }\n  ]\n}'
        )
        
        info_label = ctk.CTkLabel(
            container,
            text=info_text,
            justify="left",
            anchor="w",
            wraplength=600
        )
        
        # Dateiauswahl-Frame
        file_frame = ctk.CTkFrame(container)
        
        # Label für Dateiauswahl
        file_label = ctk.CTkLabel(
            file_frame,
            text="JSON-Datei:",
            font=ctk.CTkFont(weight="bold"),
            width=80
        )
        
        # Eingabefeld für Dateipfad
        self.projekt_file_entry = ctk.CTkEntry(file_frame, width=400)
        
        # Button für Datei-Browser
        self.projekt_browse_button = ctk.CTkButton(
            file_frame,
            text="Durchsuchen",
            command=lambda: self._browse_file("json", self.projekt_file_entry),
            width=100
        )
        
        # Optionen-Frame
        options_frame = ctk.CTkFrame(container)
        
        # Optionen für den Import
        self.projekt_merge_var = ctk.BooleanVar(value=False)
        self.projekt_merge_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Zusammenführen (bestehende Projekte beibehalten)",
            variable=self.projekt_merge_var
        )
        
        self.projekt_backup_var = ctk.BooleanVar(value=True)
        self.projekt_backup_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Backup erstellen",
            variable=self.projekt_backup_var
        )
        
        # Import-Button
        self.projekt_import_button = ctk.CTkButton(
            container,
            text="Projektdaten importieren",
            command=self._import_projekt_data,
            font=ctk.CTkFont(weight="bold"),
            height=40,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3")
        )
        
        # Layout für den Projektdaten-Tab
        container.pack(fill="both", expand=True, padx=20, pady=20)
        info_label.pack(anchor="w", pady=(0, 15))
        
        # Dateiauswahl-Frame
        file_frame.pack(fill="x", pady=(0, 15))
        file_label.pack(side="left")
        self.projekt_file_entry.pack(side="left", padx=10)
        self.projekt_browse_button.pack(side="left")
        
        # Optionen-Frame
        options_frame.pack(fill="x", pady=(0, 20))
        self.projekt_merge_checkbox.pack(anchor="w", pady=5)
        self.projekt_backup_checkbox.pack(anchor="w", pady=5)
        
        # Import-Button
        self.projekt_import_button.pack(fill="x", pady=(10, 0))
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Frame
        self.title_frame.pack(fill="x", padx=20, pady=(20, 10))
        self.title_label.pack(pady=10)
        
        # Status-Frame
        self.status_frame.pack(fill="x", padx=20, pady=(0, 10))
        self.status_label.pack(pady=5, fill="x")
        
        # Tab-View (nimmt den meisten Platz ein)
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Button-Frame (unten)
        self.button_frame.pack(fill="x", padx=20, pady=(5, 20))
        self.close_button.pack(side="right", padx=10)
    
    def _on_file_selected(self, filename):
        """
        Wird aufgerufen, wenn eine Datei ausgewählt oder fallen gelassen wurde
        
        Args:
            filename: Pfad zur ausgewählten Datei
        """
        self.selected_file = filename
        self.file_name_label.configure(text=os.path.basename(filename))
        self.bw_tool_import_button.configure(state="normal")
        
        # Status zurücksetzen
        self._show_status("", False)
    
    def _browse_file(self, file_type, entry_widget):
        """
        Öffnet einen Datei-Browser-Dialog
        
        Args:
            file_type: Der Dateityp (csv oder json)
            entry_widget: Das Eingabefeld, in das der Dateipfad eingetragen werden soll
        """
        if file_type.lower() == "csv":
            filetypes = [("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")]
            title = "CSV-Datei auswählen"
        elif file_type.lower() == "json":
            filetypes = [("JSON-Dateien", "*.json"), ("Alle Dateien", "*.*")]
            title = "JSON-Datei auswählen"
        else:
            filetypes = [("Alle Dateien", "*.*")]
            title = "Datei auswählen"
        
        # Datei-Dialog öffnen
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes
        )
        
        # Wenn eine Datei ausgewählt wurde, den Pfad in das Eingabefeld eintragen
        if filename:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, filename)
            
            # Status zurücksetzen
            self._show_status("", False)
    
    def _import_bw_tool_data(self):
        """Importiert BW Tool Daten aus einer Excel-Datei"""
        # Prüfen, ob eine Datei ausgewählt wurde
        if not self.selected_file:
            self._show_status("Bitte wählen Sie eine Excel-Datei aus.", error=True)
            return
        
        # Prüfen, ob die Datei existiert
        if not os.path.exists(self.selected_file):
            self._show_status(f"Die Datei '{self.selected_file}' existiert nicht.", error=True)
            return
        
        try:
            # Excel-Datei einlesen
            self._show_status("Excel-Datei wird eingelesen...", error=False)
            
            # Ziel-Dateipfad
            target_path = os.path.join("data", "kapa_data.csv")
            
            # Backup erstellen, falls gewünscht
            if self.bw_tool_backup_var.get() and os.path.exists(target_path):
                backup_path = f"{target_path}.bak"
                shutil.copy2(target_path, backup_path)
            
            # Verarbeite die Excel-Datei
            processed_data = self._process_bw_tool_file(self.selected_file)
            
            # Wenn keine Daten verarbeitet wurden
            if not processed_data:
                self._show_status("Keine gültigen Daten in der Datei gefunden.", error=True)
                return
            
            # Einträge überschreiben, falls vorhanden
            if self.bw_tool_overwrite_var.get() and os.path.exists(target_path):
                final_data = self._merge_with_existing_data(processed_data, target_path)
            else:
                # Nur neue Daten hinzufügen
                final_data = processed_data
                if os.path.exists(target_path):
                    with open(target_path, 'r', newline='', encoding='utf-8') as existing_file:
                        reader = csv.reader(existing_file, delimiter=';')
                        existing_data = list(reader)
                        # Nur Zeilen hinzufügen, die noch nicht existieren
                        id_date_pairs = {(row[0], row[1]) for row in existing_data}
                        final_data = [row for row in processed_data if (row[0], row[1]) not in id_date_pairs]
                        final_data = existing_data + final_data
            
            # Daten speichern
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, 'w', newline='', encoding='utf-8') as output_file:
                writer = csv.writer(output_file, delimiter=';')
                writer.writerows(final_data)
            
            # Erfolgsmeldung anzeigen
            self._show_status(f"Import erfolgreich: {len(processed_data)} Einträge verarbeitet", error=False)
        
        except Exception as e:
            # Fehlermeldung anzeigen
            self._show_status(f"Fehler beim Import: {str(e)}", error=True)
            import traceback
            traceback.print_exc()
    
    def _process_bw_tool_file(self, file_path):
        """
        Verarbeitet die BW Tool Datei unter Berücksichtigung von MIME-Formaten
        mit erweitertem Debug-Modus für Spaltenidentifikation
        
        Args:
            file_path: Pfad zur Datei
            
        Returns:
            List[List[str]]: Verarbeitete Daten im Format [[ID, DATUM, STUNDEN], ...]
        """
        try:
            # Aktiviere Debug-Modus
            debug_mode = True
            
            # Zuerst prüfen, um welche Art von Datei es sich handelt
            with open(file_path, 'rb') as f:
                header = f.read(4096)
            
            # MIME-Typ-Erkennung
            if header.startswith(b'MIME-Ver') or b'Content-Type:' in header or b'<html' in header.lower() or b'<!DOCTYPE' in header.lower():
                # Es handelt sich um eine E-Mail, HTML oder ähnliche Datei
                self._show_status("MIME/HTML-Datei erkannt, versuche Extraktion...", error=False)
                return self._process_mime_file(file_path, debug_mode=debug_mode)
            
            # Standard Excel-Verarbeitung versuchen
            try:
                # Excel-Datei mit pandas einlesen
                extension = os.path.splitext(file_path)[1].lower()
                
                if extension == '.xls':
                    df = pd.read_excel(file_path, engine='xlrd')
                else:
                    df = pd.read_excel(file_path, engine='openpyxl')
                
                if debug_mode:
                    self._show_status(f"Excel-Datei gelesen: {df.shape[0]} Zeilen, {df.shape[1]} Spalten", error=False)
                    
                    # Die ersten 10 Spaltenüberschriften ausgeben, falls vorhanden
                    print("Spaltenüberschriften:")
                    for i in range(min(df.shape[1], 20)):
                        print(f"Spalte {i+1} (Index {i}): {df.columns[i]}")
                    
                    # Die ersten 5 Zeilen ausgeben, um zu sehen, welche Daten darin enthalten sind
                    print("\nErste Zeilen der Datei:")
                    for row_idx in range(min(df.shape[0], 5)):
                        print(f"Zeile {row_idx+1}:")
                        for col_idx in range(min(df.shape[1], 20)):
                            print(f"  Spalte {col_idx+1} (Index {col_idx}): {df.iloc[row_idx, col_idx]}")
                    
                    # Suche nach Forecast-Einträgen in den Spalten
                    print("\nSuche nach Forecast-Einträgen:")
                    forecast_columns = []
                    for col_idx in range(df.shape[1]):
                        column_data = df.iloc[:min(df.shape[0], 20), col_idx].astype(str)
                        forecast_found = False
                        for value in column_data:
                            if "FCAST" in value.upper() or "FORECAST" in value.upper():
                                forecast_found = True
                                break
                        if forecast_found:
                            forecast_columns.append(col_idx)
                            print(f"Forecast-Einträge in Spalte {col_idx+1} (Index {col_idx}) gefunden")
                
                # Prüfe, ob die Datei genügend Spalten hat
                if df.shape[1] < 13:
                    self._show_status(f"Warnung: Die Datei hat nur {df.shape[1]} Spalten statt der erwarteten 13+", error=False)
                
                # Liste für die verarbeiteten Daten
                processed_data = []
                
                # Versuche, die richtigen Spaltenindizes zu finden
                id_col = None
                date_col = None
                hours_col = None
                forecast_col = None
                
                # Durchsuche Spaltenüberschriften, falls vorhanden
                column_headers = df.columns.astype(str)
                for i, header in enumerate(column_headers):
                    header_lower = header.lower()
                    if "id" in header_lower or "mitarbeiter" in header_lower or "name" in header_lower:
                        id_col = i
                    elif "datum" in header_lower or "date" in header_lower:
                        date_col = i
                    elif "stunde" in header_lower or "hour" in header_lower:
                        hours_col = i
                    elif "forecast" in header_lower or "fcast" in header_lower:
                        forecast_col = i
                
                # Spalten manuell zuweisen, wenn sie nicht gefunden wurden
                if id_col is None:
                    # Standard: Spalte 6 (Index 5)
                    id_col = 5
                    
                if date_col is None:
                    # Standard: Spalte 8 (Index 7)
                    date_col = 7
                    
                if hours_col is None:
                    # Standard: Spalte 13 (Index 12)
                    hours_col = 12
                    
                if forecast_col is None:
                    # Standard: Spalte 11 (Index 10)
                    forecast_col = 10
                
                if debug_mode:
                    print(f"\nSpaltenzuordnung:")
                    print(f"ID-Spalte: {id_col+1} (Index {id_col})")
                    print(f"Datums-Spalte: {date_col+1} (Index {date_col})")
                    print(f"Stunden-Spalte: {hours_col+1} (Index {hours_col})")
                    print(f"Forecast-Spalte: {forecast_col+1} (Index {forecast_col})")
                
                # Zeilen durchgehen und Daten extrahieren
                found_entries = 0
                for row_idx, row in df.iterrows():
                    try:
                        # Prüfe, ob die Zeile ein Forecast-Eintrag ist
                        forecast_value = str(row.iloc[forecast_col]).upper() if forecast_col < len(row) and pd.notna(row.iloc[forecast_col]) else ""
                        if not any(keyword in forecast_value for keyword in ["FCAST", "FORECAST"]):
                            continue
                        
                        # ID extrahieren
                        mitarbeiter_id = str(row.iloc[id_col]).strip() if id_col < len(row) and pd.notna(row.iloc[id_col]) else ""
                        if not mitarbeiter_id:
                            continue
                        
                        # Datum extrahieren
                        datum_value = row.iloc[date_col] if date_col < len(row) else None
                        if pd.isna(datum_value):
                            continue
                        
                        # Datum in das richtige Format konvertieren
                        if isinstance(datum_value, str):
                            # Versuche, Datumsstring zu parsen
                            try:
                                datum_obj = pd.to_datetime(datum_value, dayfirst=True)
                            except:
                                continue
                        else:
                            # Bereits ein Datum/Timestamp
                            datum_obj = pd.to_datetime(datum_value)
                        
                        datum = datum_obj.strftime("%d.%m.%Y")
                        
                        # Stunden extrahieren
                        stunden_value = row.iloc[hours_col] if hours_col < len(row) else None
                        if pd.isna(stunden_value):
                            stunden = "0.0"
                        else:
                            try:
                                stunden = str(float(stunden_value))
                            except:
                                stunden = "0.0"
                        
                        # Daten zur Liste hinzufügen
                        processed_data.append([mitarbeiter_id, datum, stunden])
                        found_entries += 1
                        
                        # Debug-Ausgabe für die ersten paar gefundenen Einträge
                        if debug_mode and found_entries <= 5:
                            print(f"\nGefundener Eintrag {found_entries}:")
                            print(f"ID: {mitarbeiter_id}")
                            print(f"Datum: {datum}")
                            print(f"Stunden: {stunden}")
                            print(f"Forecast-Wert: {forecast_value}")
                    
                    except Exception as row_error:
                        print(f"Fehler beim Verarbeiten von Zeile {row_idx+1}: {row_error}")
                        continue
                
                if debug_mode:
                    print(f"\nInsgesamt {found_entries} Einträge verarbeitet")
                
                return processed_data
                
            except Exception as excel_error:
                print(f"Standard Excel-Verarbeitung fehlgeschlagen: {excel_error}")
                # Versuche es mit der MIME-Verarbeitung als Fallback
                self._show_status("Excel-Format nicht erkannt, versuche alternatives Format...", error=False)
                return self._process_mime_file(file_path)
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            raise Exception(f"Fehler beim Verarbeiten der Datei: {str(e)}")

    def _process_mime_file(self, file_path, debug_mode=False):
        """
        Verarbeitet eine MIME-formatierte Datei (z.B. E-Mail oder HTML)
        
        Args:
            file_path: Pfad zur MIME/HTML-Datei
            debug_mode: Aktiviert zusätzliche Debug-Ausgaben
            
        Returns:
            List[List[str]]: Verarbeitete Daten im Format [[ID, DATUM, STUNDEN], ...]
        """
        try:
            # Lese den Dateiinhalt mit verschiedenen Kodierungen
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                try:
                    with open(file_path, 'r', encoding='latin1') as f:
                        content = f.read()
                except:
                    with open(file_path, 'r', encoding='cp1252') as f:
                        content = f.read()
            
            if debug_mode:
                # Zeige die ersten 500 Zeichen der Datei
                print(f"\nDateiinhalt (Anfang):\n{content[:500]}...\n")
            
            # Liste für die verarbeiteten Daten
            processed_data = []
            
            # Import regex
            import re
            
            # Extrahiere alle Tabellen aus HTML
            table_pattern = r'<table[^>]*>(.*?)</table>'
            tables = re.findall(table_pattern, content, re.DOTALL | re.IGNORECASE)
            
            if debug_mode:
                print(f"{len(tables)} Tabellen in der Datei gefunden")
            
            if tables:
                # Es wurden Tabellen gefunden, versuche die Daten zu extrahieren
                for table_idx, table in enumerate(tables):
                    # Extrahiere alle Tabellenzeilen
                    row_pattern = r'<tr[^>]*>(.*?)</tr>'
                    rows = re.findall(row_pattern, table, re.DOTALL | re.IGNORECASE)
                    
                    if debug_mode:
                        print(f"Tabelle {table_idx+1}: {len(rows)} Zeilen gefunden")
                    
                    # Speichere die Spaltenindizes für ID, Datum und Stunden
                    id_col = None
                    date_col = None
                    hours_col = None
                    forecast_col = None
                    
                    # Zeilen durchgehen
                    for row_idx, row in enumerate(rows):
                        # Extrahiere alle Zellen in dieser Zeile
                        cell_pattern = r'<t[dh][^>]*>(.*?)</t[dh]>'
                        cells = re.findall(cell_pattern, row, re.DOTALL | re.IGNORECASE)
                        
                        # Bereinige HTML-Tags und Whitespace
                        cells = [re.sub(r'<[^>]+>', '', cell).strip() for cell in cells]
                        
                        if debug_mode and row_idx < 3:
                            print(f"Zeile {row_idx+1} in Tabelle {table_idx+1}: {cells}")
                        
                        # Wenn es die erste Zeile ist, versuche die Spaltenüberschriften zu identifizieren
                        if row_idx == 0 and not (id_col and date_col and hours_col):
                            for i, cell in enumerate(cells):
                                cell_text = cell.lower()
                                # Suche nach ID-Spalte
                                if 'id' in cell_text or 'mitarbeiter' in cell_text or 'name' in cell_text:
                                    id_col = i
                                # Suche nach Datumsspalte
                                elif 'datum' in cell_text or 'date' in cell_text:
                                    date_col = i
                                # Suche nach Stundenspalte
                                elif 'stunde' in cell_text or 'hour' in cell_text:
                                    hours_col = i
                                # Suche nach Forecast-Spalte
                                elif 'forecast' in cell_text or 'fcast' in cell_text:
                                    forecast_col = i
                        
                        # Überspringe die erste Zeile, da sie wahrscheinlich Überschriften enthält
                        if row_idx == 0:
                            continue
                        
                        # Wenn wir Forecast-Einträge identifizieren können
                        forecast_found = False
                        if forecast_col is not None and len(cells) > forecast_col:
                            # Prüfe, ob die Zeile ein Forecast-Eintrag ist
                            forecast_value = cells[forecast_col].upper()
                            if any(keyword in forecast_value for keyword in ["FCAST", "FORECAST"]):
                                forecast_found = True
                        else:
                            # Wenn wir keine Forecast-Spalte haben, prüfe alle Zellen
                            for cell in cells:
                                if "FCAST" in cell.upper() or "FORECAST" in cell.upper():
                                    forecast_found = True
                                    break
                        
                        if not forecast_found:
                            continue
                        
                        # Extrahiere die benötigten Informationen
                        # Wir verwenden die identifizierten Spalten oder suchen in allen Zellen
                        
                        # ID extrahieren
                        mitarbeiter_id = ""
                        if id_col is not None and len(cells) > id_col:
                            mitarbeiter_id = cells[id_col].strip()
                        else:
                            # Suche nach ID in allen Zellen
                            for cell in cells:
                                if re.match(r'^[A-Z0-9][A-Z0-9._-]{2,}$', cell.strip()):
                                    mitarbeiter_id = cell.strip()
                                    break
                        
                        if not mitarbeiter_id:
                            continue
                        
                        # Datum extrahieren
                        datum = ""
                        if date_col is not None and len(cells) > date_col:
                            date_text = cells[date_col]
                            try:
                                datum_obj = pd.to_datetime(date_text, dayfirst=True)
                                datum = datum_obj.strftime("%d.%m.%Y")
                            except:
                                pass
                        
                        if not datum:
                            # Suche nach Datum in allen Zellen
                            date_patterns = [
                                r'\b\d{2}\.\d{2}\.\d{4}\b',  # DD.MM.YYYY
                                r'\b\d{2}/\d{2}/\d{4}\b',    # DD/MM/YYYY
                                r'\b\d{4}-\d{2}-\d{2}\b'     # YYYY-MM-DD
                            ]
                            
                            for cell in cells:
                                for pattern in date_patterns:
                                    date_match = re.search(pattern, cell)
                                    if date_match:
                                        try:
                                            datum_obj = pd.to_datetime(date_match.group(0), dayfirst=True)
                                            datum = datum_obj.strftime("%d.%m.%Y")
                                            break
                                        except:
                                            pass
                                if datum:
                                    break
                        
                        if not datum:
                            continue
                        
                        # Stunden extrahieren
                        stunden = "0.0"
                        if hours_col is not None and len(cells) > hours_col:
                            hours_text = cells[hours_col]
                            try:
                                hours_value = float(hours_text.replace(',', '.'))
                                stunden = str(hours_value)
                            except:
                                # Versuche, eine Zahl aus dem Text zu extrahieren
                                hours_match = re.search(r'\b\d+[,.]?\d*\b', hours_text)
                                if hours_match:
                                    try:
                                        hours_value = float(hours_match.group(0).replace(',', '.'))
                                        stunden = str(hours_value)
                                    except:
                                        pass
                        
                        if stunden == "0.0":
                            # Suche nach Stunden in allen Zellen
                            for cell in cells:
                                hours_match = re.search(r'\b\d+[,.]?\d*\b', cell)
                                if hours_match:
                                    try:
                                        hours_value = float(hours_match.group(0).replace(',', '.'))
                                        if 0 <= hours_value <= 24:
                                            stunden = str(hours_value)
                                            break
                                    except:
                                        pass
                        
                        # Daten zur Liste hinzufügen
                        processed_data.append([mitarbeiter_id, datum, stunden])
                        
                        if debug_mode and len(processed_data) <= 5:
                            print(f"\nGefundener Eintrag {len(processed_data)}:")
                            print(f"ID: {mitarbeiter_id}")
                            print(f"Datum: {datum}")
                            print(f"Stunden: {stunden}")
            
            # Wenn keine Tabellen gefunden wurden oder keine Daten extrahiert werden konnten,
            # versuche es mit dem Gesamttext
            if not processed_data:
                if debug_mode:
                    print("\nKeine Daten aus Tabellen extrahiert, versuche Text-Analyse...")
                
                # Suche nach relevanten Zeilen im Gesamttext
                lines = content.split('\n')
                forecast_lines = [line for line in lines if 'FCAST' in line.upper() or 'FORECAST' in line.upper()]
                
                if debug_mode:
                    print(f"{len(forecast_lines)} Zeilen mit 'FORECAST' gefunden")
                
                for line_idx, line in enumerate(forecast_lines):
                    # Entferne HTML-Tags
                    line = re.sub(r'<[^>]+>', ' ', line)
                    
                    if debug_mode and line_idx < 5:
                        print(f"Zeile {line_idx+1} mit Forecast: {line[:100]}...")
                    
                    # Suche nach ID
                    id_match = re.search(r'\b[A-Z0-9][A-Z0-9._-]{2,}\b', line)
                    if not id_match:
                        continue
                    
                    mitarbeiter_id = id_match.group(0)
                    
                    # Suche nach Datum
                    date_patterns = [
                        r'\b\d{2}\.\d{2}\.\d{4}\b',  # DD.MM.YYYY
                        r'\b\d{2}/\d{2}/\d{4}\b',    # DD/MM/YYYY
                        r'\b\d{4}-\d{2}-\d{2}\b'     # YYYY-MM-DD
                    ]
                    
                    datum = ""
                    for pattern in date_patterns:
                        date_match = re.search(pattern, line)
                        if date_match:
                            try:
                                datum_obj = pd.to_datetime(date_match.group(0), dayfirst=True)
                                datum = datum_obj.strftime("%d.%m.%Y")
                                break
                            except:
                                pass
                    
                    if not datum:
                        continue
                    
                    # Suche nach Stunden
                    stunden = "0.0"
                    hours_match = re.search(r'\b\d+[,.]?\d*\b', line)
                    if hours_match:
                        try:
                            hours_value = float(hours_match.group(0).replace(',', '.'))
                            if 0 <= hours_value <= 24:
                                stunden = str(hours_value)
                        except:
                            pass
                    
                    # Daten zur Liste hinzufügen
                    processed_data.append([mitarbeiter_id, datum, stunden])
                    
                    if debug_mode and len(processed_data) <= 5:
                        print(f"\nGefundener Eintrag {len(processed_data)} aus Text:")
                        print(f"ID: {mitarbeiter_id}")
                        print(f"Datum: {datum}")
                        print(f"Stunden: {stunden}")
            
            if debug_mode:
                print(f"\nInsgesamt {len(processed_data)} Einträge gefunden und verarbeitet")
            
            return processed_data
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"MIME-Verarbeitung fehlgeschlagen: {e}")
            return []
    
    def _import_projekt_data(self):
        """Importiert Projektdaten aus einer JSON-Datei"""
        # Dateipfad auslesen
        file_path = self.projekt_file_entry.get().strip()
        
        # Prüfen, ob eine Datei ausgewählt wurde
        if not file_path:
            self._show_status("Bitte wählen Sie eine JSON-Datei aus.", error=True)
            return
        
        # Prüfen, ob die Datei existiert
        if not os.path.exists(file_path):
            self._show_status(f"Die Datei '{file_path}' existiert nicht.", error=True)
            return
        
        try:
            # Ziel-Dateipfad
            target_path = os.path.join("data", "project.json")
            
            # Backup erstellen, falls gewünscht
            if self.projekt_backup_var.get() and os.path.exists(target_path):
                backup_path = f"{target_path}.bak"
                shutil.copy2(target_path, backup_path)
            
            # JSON-Datei einlesen
            with open(file_path, 'r', encoding='utf-8') as json_file:
                data = json.load(json_file)
            
            # Wenn Zusammenführen gewählt wurde, bestehende Daten laden
            if self.projekt_merge_var.get() and os.path.exists(target_path):
                with open(target_path, 'r', encoding='utf-8') as existing_file:
                    existing_data = json.load(existing_file)
                
                # Daten zusammenführen
                if "projekte" in data and "projekte" in existing_data:
                    # Erstelle ein Set mit den Namen der bestehenden Projekte
                    existing_project_names = {p["name"] for p in existing_data["projekte"]}
                    
                    # Füge nur neue Projekte hinzu
                    for project in data["projekte"]:
                        if project["name"] not in existing_project_names:
                            existing_data["projekte"].append(project)
                    
                    # Aktualisierte Daten verwenden
                    data = existing_data
            
            # Daten in die Zieldatei schreiben
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, 'w', encoding='utf-8') as output_file:
                json.dump(data, output_file, indent=4, ensure_ascii=False)
            
            # Erfolgsmeldung anzeigen
            projekt_count = len(data.get("projekte", []))
            self._show_status(f"Import erfolgreich: {projekt_count} Projekte in '{target_path}'", error=False)
        
        except Exception as e:
            # Fehlermeldung anzeigen
            self._show_status(f"Fehler beim Import: {e}", error=True)
    
    def _show_status(self, message, error=False):
        """
        Zeigt eine Statusmeldung an
        
        Args:
            message: Die anzuzeigende Nachricht
            error: True für eine Fehlermeldung, False für eine Erfolgsmeldung
        """
        if not message:
            self.status_label.configure(text="")
            return
        
        # Nachricht anzeigen
        self.status_label.configure(
            text=message,
            text_color="red" if error else "green"
        )
    
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

    def _merge_with_existing_data(self, new_data, existing_file_path):
        """
        Führt neue Daten mit vorhandenen Daten zusammen, wobei vorhandene Einträge überschrieben werden
        
        Args:
            new_data: Neue Daten [[ID, DATUM, STUNDEN], ...]
            existing_file_path: Pfad zur bestehenden CSV-Datei
                
        Returns:
            List[List[str]]: Zusammengeführte Daten
        """
        if not os.path.exists(existing_file_path):
            return new_data
        
        try:
            # Vorhandene Daten einlesen
            with open(existing_file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file, delimiter=';')
                existing_data = list(reader)
            
            # Dictionary für schnellere Suche erstellen (ID+DATUM -> Zeilenindex)
            id_date_indices = {(row[0], row[1]): i for i, row in enumerate(existing_data)}
            
            # Neue Daten durchgehen
            for row in new_data:
                key = (row[0], row[1])
                if key in id_date_indices:
                    # Vorhandenen Eintrag überschreiben
                    existing_data[id_date_indices[key]] = row
                else:
                    # Neuen Eintrag hinzufügen
                    existing_data.append(row)
            
            return existing_data
        
        except Exception as e:
            raise Exception(f"Fehler beim Zusammenführen der Daten: {str(e)}")