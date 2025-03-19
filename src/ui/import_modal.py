"""
Modal-Fenster zum Importieren von Daten mit zwei Tabs
"""
import os
import customtkinter as ctk
from typing import Dict, Callable, Optional
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import json
import shutil
import datetime

class ImportModal(ctk.CTkToplevel):
    """
    Modal-Fenster zum Importieren von Daten mit zwei Tabs
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
        self.geometry("700x500")
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
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        # Titel-Frame
        self.title_frame = ctk.CTkFrame(self)
        self.title_label = ctk.CTkLabel(
            self.title_frame, 
            text="Daten importieren",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        
        # Status-Label (initial versteckt)
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
        self.kapa_tab = self.tab_view.add("Kapazitätsdaten")
        self.projekt_tab = self.tab_view.add("Projektdaten")
        
        # Inhalte für Kapazitätsdaten-Tab
        self._create_kapa_tab()
        
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
    
    def _create_kapa_tab(self):
        """Erstellt die Inhalte für den Kapazitätsdaten-Tab"""
        # Container-Frame für den Tab
        container = ctk.CTkFrame(self.kapa_tab)
        
        # Info-Text
        info_text = (
            "Importieren Sie Kapazitätsdaten aus einer CSV-Datei.\n\n"
            "Die CSV-Datei sollte folgendes Format haben:\n"
            "MITARBEITER_ID;DATUM;STUNDEN\n\n"
            "Beispiel:\n"
            "DESUCIU;01.04.2025;8.0\n"
            "DEPERCIULEAC;01.04.2025;8.0\n"
            "..."
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
            text="CSV-Datei:",
            font=ctk.CTkFont(weight="bold"),
            width=80
        )
        
        # Eingabefeld für Dateipfad
        self.kapa_file_entry = ctk.CTkEntry(file_frame, width=400)
        
        # Button für Datei-Browser
        self.kapa_browse_button = ctk.CTkButton(
            file_frame,
            text="Durchsuchen",
            command=lambda: self._browse_file("csv", self.kapa_file_entry),
            width=100
        )
        
        # Optionen-Frame
        options_frame = ctk.CTkFrame(container)
        
        # Optionen für den Import
        self.kapa_append_var = ctk.BooleanVar(value=True)
        self.kapa_append_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Anhängen (bestehende Daten beibehalten)",
            variable=self.kapa_append_var
        )
        
        self.kapa_backup_var = ctk.BooleanVar(value=True)
        self.kapa_backup_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Backup erstellen",
            variable=self.kapa_backup_var
        )
        
        # Import-Button
        self.kapa_import_button = ctk.CTkButton(
            container,
            text="Kapazitätsdaten importieren",
            command=self._import_kapa_data,
            font=ctk.CTkFont(weight="bold"),
            height=40,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3")
        )
        
        # Layout für den Kapazitätsdaten-Tab
        container.pack(fill="both", expand=True, padx=20, pady=20)
        info_label.pack(anchor="w", pady=(0, 15))
        
        # Dateiauswahl-Frame
        file_frame.pack(fill="x", pady=(0, 15))
        file_label.pack(side="left")
        self.kapa_file_entry.pack(side="left", padx=10)
        self.kapa_browse_button.pack(side="left")
        
        # Optionen-Frame
        options_frame.pack(fill="x", pady=(0, 20))
        self.kapa_append_checkbox.pack(anchor="w", pady=5)
        self.kapa_backup_checkbox.pack(anchor="w", pady=5)
        
        # Import-Button
        self.kapa_import_button.pack(fill="x", pady=(10, 0))
    
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
        
        # Status-Frame (initial unsichtbar)
        self.status_frame.pack(fill="x", padx=20, pady=(0, 10))
        self.status_label.pack(pady=5, fill="x")
        
        # Tab-View (nimmt den meisten Platz ein)
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Button-Frame (unten)
        self.button_frame.pack(fill="x", padx=20, pady=(5, 20))
        self.close_button.pack(side="right", padx=10)
    
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
    
    def _import_kapa_data(self):
        """Importiert Kapazitätsdaten aus einer CSV-Datei"""
        # Dateipfad auslesen
        file_path = self.kapa_file_entry.get().strip()
        
        # Prüfen, ob eine Datei ausgewählt wurde
        if not file_path:
            self._show_status("Bitte wählen Sie eine CSV-Datei aus.", error=True)
            return
        
        # Prüfen, ob die Datei existiert
        if not os.path.exists(file_path):
            self._show_status(f"Die Datei '{file_path}' existiert nicht.", error=True)
            return
        
        try:
            # Ziel-Dateipfad
            target_path = os.path.join("data", "kapa_data.csv")
            
            # Backup erstellen, falls gewünscht
            if self.kapa_backup_var.get() and os.path.exists(target_path):
                backup_path = f"{target_path}.bak"
                shutil.copy2(target_path, backup_path)
            
            # CSV-Datei einlesen
            with open(file_path, 'r', newline='', encoding='utf-8') as csv_file:
                try:
                    # Versuche den Delimiter zu erkennen
                    dialect = csv.Sniffer().sniff(csv_file.read(1024))
                    csv_file.seek(0)
                    reader = csv.reader(csv_file, dialect)
                except:
                    # Fallback, wenn Sniffer fehlschlägt
                    csv_file.seek(0)
                    reader = csv.reader(csv_file, delimiter=';')
                
                # Daten einlesen
                data = list(reader)
            
            # Wenn Anhängen gewählt wurde, bestehende Daten laden
            if self.kapa_append_var.get() and os.path.exists(target_path):
                with open(target_path, 'r', newline='', encoding='utf-8') as existing_file:
                    try:
                        # Versuche den Delimiter zu erkennen
                        dialect = csv.Sniffer().sniff(existing_file.read(1024))
                        existing_file.seek(0)
                        existing_reader = csv.reader(existing_file, dialect)
                    except:
                        # Fallback, wenn Sniffer fehlschlägt
                        existing_file.seek(0)
                        existing_reader = csv.reader(existing_file, delimiter=';')
                    
                    # Bestehende Daten laden
                    existing_data = list(existing_reader)
                
                # Neue Daten an bestehende Daten anhängen
                # Duplikate vermeiden
                existing_lines = set(tuple(line) for line in existing_data)
                new_lines = [line for line in data if tuple(line) not in existing_lines]
                data = existing_data + new_lines
            
            # Daten in die Zieldatei schreiben
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, 'w', newline='', encoding='utf-8') as output_file:
                writer = csv.writer(output_file, delimiter=';')
                writer.writerows(data)
            
            # Erfolgsmeldung anzeigen
            self._show_status(f"Import erfolgreich: {len(data)} Einträge in '{target_path}'", error=False)
        
        except Exception as e:
            # Fehlermeldung anzeigen
            self._show_status(f"Fehler beim Import: {e}", error=True)
    
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