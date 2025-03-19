"""
Implementierung der Mitarbeiter-Funktionalitäten (Hinzufügen, Bearbeiten, Löschen)
für den CAMP Manager
"""
import os
import json
import customtkinter as ctk
from typing import Dict, List, Callable, Optional, Any
import datetime
import tkinter as tk
from tkinter import messagebox

class MitarbeiterDialog(ctk.CTkToplevel):
    """
    Dialog zum Hinzufügen oder Bearbeiten eines Mitarbeiters
    """
    def __init__(self, master, title="Mitarbeiter", on_save=None, mitarbeiter_data=None):
        """
        Initialisiert den Dialog
        
        Args:
            master: Das übergeordnete Widget
            title: Titel des Dialogs
            on_save: Callback-Funktion, die aufgerufen wird, wenn der Benutzer speichert
            mitarbeiter_data: Vorhandene Mitarbeiterdaten zum Bearbeiten (oder None für einen neuen Mitarbeiter)
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title(title)
        self.geometry("450x320")
        self.resizable(False, False)
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.on_save = on_save
        self.mitarbeiter_data = mitarbeiter_data or {}
        self.result = None
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Eingabefelder mit vorhandenen Daten füllen, falls Daten übergeben wurden
        self._fill_fields()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für den Dialog"""
        # Titel-Label
        self.title_label = ctk.CTkLabel(
            self, 
            text="Mitarbeiter Informationen",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        # Formular-Frame
        self.form_frame = ctk.CTkFrame(self)
        
        # ID-Feld
        self.id_label = ctk.CTkLabel(
            self.form_frame, 
            text="ID:",
            anchor="w",
            width=100
        )
        self.id_entry = ctk.CTkEntry(self.form_frame, width=250)
        
        # Name-Feld
        self.name_label = ctk.CTkLabel(
            self.form_frame, 
            text="Name:",
            anchor="w",
            width=100
        )
        self.name_entry = ctk.CTkEntry(self.form_frame, width=250)
        
        # Rolle-Feld
        self.rolle_label = ctk.CTkLabel(
            self.form_frame, 
            text="Rolle:",
            anchor="w",
            width=100
        )
        self.rolle_entry = ctk.CTkEntry(self.form_frame, width=250)
        
        # FTE-Feld (Full Time Equivalent)
        self.fte_label = ctk.CTkLabel(
            self.form_frame, 
            text="FTE (0.0-1.0):",
            anchor="w",
            width=100
        )
        self.fte_entry = ctk.CTkEntry(self.form_frame, width=250)
        self.fte_entry.insert(0, "1.0")  # Standardwert
        
        # Zusätzliches Info-Label
        self.info_label = ctk.CTkLabel(
            self, 
            text="Alle Felder müssen ausgefüllt werden.",
            text_color=("gray60", "gray70")
        )
        
        # Button-Frame
        self.button_frame = ctk.CTkFrame(self)
        
        # Abbrechen-Button
        self.cancel_button = ctk.CTkButton(
            self.button_frame, 
            text="Abbrechen",
            command=self.destroy,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            width=100
        )
        
        # Speichern-Button
        self.save_button = ctk.CTkButton(
            self.button_frame, 
            text="Speichern",
            command=self._on_save,
            width=100
        )
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Label
        self.title_label.pack(padx=20, pady=(20, 15))
        
        # Formular-Frame
        self.form_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        # Formular-Felder
        self.id_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.name_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        self.rolle_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rolle_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        self.fte_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.fte_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        
        # Info-Label
        self.info_label.pack(padx=20, pady=(5, 15))
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=20, pady=(0, 20))
        self.cancel_button.pack(side="left", padx=5)
        self.save_button.pack(side="right", padx=5)
    
    def _fill_fields(self):
        """Füllt die Eingabefelder mit vorhandenen Daten, falls vorhanden"""
        if not self.mitarbeiter_data:
            return
        
        # Felder mit Daten füllen
        if "id" in self.mitarbeiter_data:
            self.id_entry.delete(0, "end")
            self.id_entry.insert(0, self.mitarbeiter_data["id"])
        
        if "name" in self.mitarbeiter_data:
            self.name_entry.delete(0, "end")
            self.name_entry.insert(0, self.mitarbeiter_data["name"])
        
        if "rolle" in self.mitarbeiter_data:
            self.rolle_entry.delete(0, "end")
            self.rolle_entry.insert(0, self.mitarbeiter_data["rolle"])
        
        if "fte" in self.mitarbeiter_data:
            self.fte_entry.delete(0, "end")
            self.fte_entry.insert(0, str(self.mitarbeiter_data["fte"]))
    
    def _on_save(self):
        """Wird aufgerufen, wenn der Speichern-Button geklickt wird"""
        # Daten aus Formular auslesen
        mitarbeiter_id = self.id_entry.get().strip()
        name = self.name_entry.get().strip()
        rolle = self.rolle_entry.get().strip()
        fte_str = self.fte_entry.get().strip()
        
        # Validieren
        if not mitarbeiter_id or not name or not rolle or not fte_str:
            messagebox.showerror(
                "Fehlende Daten",
                "Bitte füllen Sie alle Felder aus.",
                parent=self
            )
            return
        
        # FTE-Wert validieren
        try:
            fte = float(fte_str)
            if fte < 0 or fte > 1:
                messagebox.showerror(
                    "Ungültiger FTE-Wert",
                    "Der FTE-Wert muss zwischen 0.0 und 1.0 liegen.",
                    parent=self
                )
                return
        except ValueError:
            messagebox.showerror(
                "Ungültiger FTE-Wert",
                "Bitte geben Sie eine gültige Zahl für FTE ein.",
                parent=self
            )
            return
        
        # Ergebnis zusammenstellen
        self.result = {
            "id": mitarbeiter_id,
            "name": name,
            "rolle": rolle,
            "fte": fte
        }
        
        # Callback aufrufen
        if self.on_save:
            self.on_save(self.result)
        
        # Dialog schließen
        self.destroy()
    
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


class CAMPManagerModal(ctk.CTkToplevel):
    """
    Modal-Fenster für den CAMP Manager mit Tabs für Projekte, Sprints und Mitarbeiter
    Mitarbeiter-Funktionen sind vollständig implementiert (Hinzufügen, Bearbeiten, Löschen)
    """
    def __init__(self, master, callbacks: Optional[Dict[str, Callable]] = None):
        """
        Initialisiert das Modal-Fenster
        
        Args:
            master: Das übergeordnete Widget
            callbacks: Dictionary mit Callback-Funktionen für die Aktionen
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("CAMP Manager")
        self.geometry("800x600")
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.callbacks = callbacks or {}
        self.projects_data = self._load_project_data()
        self.current_project = None
        self.current_sprint = None
        self.selected_employee = None  # Für die Mitarbeiter-Auswahl
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
    
    def _load_project_data(self):
        """
        Lädt die Projektdaten aus der project.json-Datei
        
        Returns:
            dict: Die geladenen Projektdaten oder ein leeres Dict
        """
        try:
            json_path = os.path.join("data", "project.json")
            if os.path.exists(json_path):
                with open(json_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                print(f"Warnung: Datei {json_path} nicht gefunden.")
                return {"projekte": []}
        except Exception as e:
            print(f"Fehler beim Laden der Projektdaten: {e}")
            return {"projekte": []}
    
    def _save_project_data(self):
        """
        Speichert die Projektdaten in der project.json-Datei
        
        Returns:
            bool: True wenn erfolgreich, False wenn ein Fehler aufgetreten ist
        """
        try:
            # Stelle sicher, dass das Verzeichnis existiert
            os.makedirs("data", exist_ok=True)
            
            # Speichere die Daten
            json_path = os.path.join("data", "project.json")
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(self.projects_data, f, indent=4, ensure_ascii=False)
            
            return True
        except Exception as e:
            print(f"Fehler beim Speichern der Projektdaten: {e}")
            messagebox.showerror(
                "Speicherfehler",
                f"Die Projektdaten konnten nicht gespeichert werden: {e}",
                parent=self
            )
            return False
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        # Titel-Frame
        self.title_frame = ctk.CTkFrame(self)
        self.title_label = ctk.CTkLabel(
            self.title_frame, 
            text="CAMP Manager",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        
        # Tab-View für die verschiedenen Bereiche
        self.tab_view = ctk.CTkTabview(self)
        
        # Tabs erstellen
        self.project_tab = self.tab_view.add("Projekte")
        self.sprint_tab = self.tab_view.add("Sprints")
        self.employee_tab = self.tab_view.add("Mitarbeiter")
        
        # Inhalte für Projekte-Tab
        self._create_project_tab()
        
        # Inhalte für Sprints-Tab
        self._create_sprint_tab()
        
        # Inhalte für Mitarbeiter-Tab
        self._create_employee_tab()
        
        # Button-Frame (unten)
        self.button_frame = ctk.CTkFrame(self)
        self.close_button = ctk.CTkButton(
            self.button_frame, 
            text="Schließen",
            command=self.destroy,
            width=120
        )
    
    def _create_project_tab(self):
        """Erstellt die Inhalte für den Projekte-Tab"""
        # Container-Frame für den Projekte-Tab
        container = ctk.CTkFrame(self.project_tab)
        
        # Projekt-Auswahl-Frame
        selection_frame = ctk.CTkFrame(container)
        
        # Label für die Projektauswahl
        project_label = ctk.CTkLabel(
            selection_frame,
            text="Aktuelles Projekt:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Dropdown für Projektauswahl
        project_names = [projekt["name"] for projekt in self.projects_data.get("projekte", [])]
        self.project_var = ctk.StringVar(value=project_names[0] if project_names else "Kein Projekt")
        
        self.project_dropdown = ctk.CTkOptionMenu(
            selection_frame,
            values=project_names if project_names else ["Kein Projekt"],
            variable=self.project_var,
            command=self._on_project_selected,
            width=250
        )
        
        # Projekt-Details-Frame
        details_frame = ctk.CTkFrame(container)
        
        # Header für Projektdetails
        details_header = ctk.CTkLabel(
            details_frame,
            text="Projektdetails:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Textbox für Projektdetails
        self.project_details = ctk.CTkTextbox(details_frame, width=400, height=300)
        self.project_details.configure(state="disabled")  # Schreibgeschützt
        
        # Aktions-Frame für Projekt-Buttons
        action_frame = ctk.CTkFrame(container)
        
        # Buttons für Projektaktionen
        self.new_project_btn = ctk.CTkButton(
            action_frame,
            text="Neues Projekt",
            command=self._on_new_project,
            width=150
        )
        
        self.edit_project_btn = ctk.CTkButton(
            action_frame,
            text="Projekt bearbeiten",
            command=self._on_edit_project,
            width=150
        )
        
        self.delete_project_btn = ctk.CTkButton(
            action_frame,
            text="Projekt löschen",
            command=self._on_delete_project,
            fg_color=("red3", "red4"),
            hover_color=("red4", "red3"),
            width=150
        )
        
        # Layout für Projekt-Tab
        # Container nimmt den gesamten Tab ein
        container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Selection-Frame oben
        selection_frame.pack(fill="x", pady=(0, 15))
        project_label.pack(side="left", padx=(0, 10))
        self.project_dropdown.pack(side="left")
        
        # Details-Frame in der Mitte
        details_frame.pack(fill="both", expand=True, pady=(0, 15))
        details_header.pack(anchor="w", pady=(0, 5))
        self.project_details.pack(fill="both", expand=True)
        
        # Action-Frame unten
        action_frame.pack(fill="x")
        self.new_project_btn.pack(side="left", padx=(0, 10))
        self.edit_project_btn.pack(side="left", padx=10)
        self.delete_project_btn.pack(side="right")
        
        # Initial das erste Projekt auswählen (falls vorhanden)
        if project_names:
            self._on_project_selected(project_names[0])
    
    def _create_sprint_tab(self):
        """Erstellt die Inhalte für den Sprints-Tab"""
        # Container-Frame für den Sprints-Tab
        container = ctk.CTkFrame(self.sprint_tab)
        
        # Projekt-Auswahl-Frame (auch hier, um zu wissen, für welches Projekt die Sprints sind)
        project_frame = ctk.CTkFrame(container)
        
        # Label für die Projektauswahl
        project_label = ctk.CTkLabel(
            project_frame,
            text="Projekt:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Dropdown für Projektauswahl
        project_names = [projekt["name"] for projekt in self.projects_data.get("projekte", [])]
        self.sprint_project_var = ctk.StringVar(value=project_names[0] if project_names else "Kein Projekt")
        
        self.sprint_project_dropdown = ctk.CTkOptionMenu(
            project_frame,
            values=project_names if project_names else ["Kein Projekt"],
            variable=self.sprint_project_var,
            command=self._on_sprint_project_selected,
            width=250
        )
        
        # Sprint-Auswahl-Frame
        sprint_frame = ctk.CTkFrame(container)
        
        # Label für die Sprintauswahl
        sprint_label = ctk.CTkLabel(
            sprint_frame,
            text="Sprint:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Dropdown für Sprintauswahl (wird dynamisch gefüllt)
        self.sprint_var = ctk.StringVar(value="Kein Sprint")
        self.sprint_dropdown = ctk.CTkOptionMenu(
            sprint_frame,
            values=["Kein Sprint"],
            variable=self.sprint_var,
            command=self._on_sprint_selected,
            width=250
        )
        
        # Sprint-Details-Frame
        details_frame = ctk.CTkFrame(container)
        
        # Header für Sprintdetails
        details_header = ctk.CTkLabel(
            details_frame,
            text="Sprint-Details:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Textbox für Sprintdetails
        self.sprint_details = ctk.CTkTextbox(details_frame, width=400, height=200)
        self.sprint_details.configure(state="disabled")  # Schreibgeschützt
        
        # Aktions-Frame für Sprint-Buttons
        action_frame = ctk.CTkFrame(container)
        
        # Buttons für Sprintaktionen
        self.new_sprint_btn = ctk.CTkButton(
            action_frame,
            text="Neuer Sprint",
            command=self._on_new_sprint,
            width=150
        )
        
        self.edit_sprint_btn = ctk.CTkButton(
            action_frame,
            text="Sprint bearbeiten",
            command=self._on_edit_sprint,
            width=150
        )
        
        self.delete_sprint_btn = ctk.CTkButton(
            action_frame,
            text="Sprint löschen",
            command=self._on_delete_sprint,
            fg_color=("red3", "red4"),
            hover_color=("red4", "red3"),
            width=150
        )
        
        # Layout für Sprint-Tab
        # Container nimmt den gesamten Tab ein
        container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Projekt-Frame oben
        project_frame.pack(fill="x", pady=(0, 10))
        project_label.pack(side="left", padx=(0, 10))
        self.sprint_project_dropdown.pack(side="left")
        
        # Sprint-Frame darunter
        sprint_frame.pack(fill="x", pady=(0, 15))
        sprint_label.pack(side="left", padx=(0, 10))
        self.sprint_dropdown.pack(side="left")
        
        # Details-Frame in der Mitte
        details_frame.pack(fill="both", expand=True, pady=(0, 15))
        details_header.pack(anchor="w", pady=(0, 5))
        self.sprint_details.pack(fill="both", expand=True)
        
        # Action-Frame unten
        action_frame.pack(fill="x")
        self.new_sprint_btn.pack(side="left", padx=(0, 10))
        self.edit_sprint_btn.pack(side="left", padx=10)
        self.delete_sprint_btn.pack(side="right")
        
        # Initial das erste Projekt auswählen (falls vorhanden)
        if project_names:
            self._on_sprint_project_selected(project_names[0])
    
    def _create_employee_tab(self):
        """Erstellt die Inhalte für den Mitarbeiter-Tab"""
        # Container-Frame für den Mitarbeiter-Tab
        container = ctk.CTkFrame(self.employee_tab)
        
        # Projekt-Auswahl-Frame (auch hier, um zu wissen, für welches Projekt die Mitarbeiter sind)
        project_frame = ctk.CTkFrame(container)
        
        # Label für die Projektauswahl
        project_label = ctk.CTkLabel(
            project_frame,
            text="Projekt:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Dropdown für Projektauswahl
        project_names = [projekt["name"] for projekt in self.projects_data.get("projekte", [])]
        self.employee_project_var = ctk.StringVar(value=project_names[0] if project_names else "Kein Projekt")
        
        self.employee_project_dropdown = ctk.CTkOptionMenu(
            project_frame,
            values=project_names if project_names else ["Kein Projekt"],
            variable=self.employee_project_var,
            command=self._on_employee_project_selected,
            width=250
        )
        
        # Mitarbeiter-Frame mit Header
        employee_header_frame = ctk.CTkFrame(container)
        
        # Label für die Mitarbeiterliste
        employee_label = ctk.CTkLabel(
            employee_header_frame,
            text="Mitarbeiter:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        
        # Info-Label für die Anzahl der Mitarbeiter
        self.employee_count_label = ctk.CTkLabel(
            employee_header_frame,
            text="(0 Mitarbeiter)",
            text_color=("gray60", "gray70")
        )
        
        # Scrollbare Liste für Mitarbeiter
        self.employee_listbox_frame = ctk.CTkScrollableFrame(
            container, 
            width=350, 
            height=300
        )
        
        # Aktions-Frame für Mitarbeiter-Buttons
        action_frame = ctk.CTkFrame(container)
        
        # Buttons für Mitarbeiteraktionen
        self.add_employee_btn = ctk.CTkButton(
            action_frame,
            text="Mitarbeiter hinzufügen",
            command=self._on_add_employee,
            width=170
        )
        
        self.edit_employee_btn = ctk.CTkButton(
            action_frame,
            text="Mitarbeiter bearbeiten",
            command=self._on_edit_employee,
            width=170,
            state="disabled"  # Initial deaktiviert, bis ein Mitarbeiter ausgewählt ist
        )
        
        self.remove_employee_btn = ctk.CTkButton(
            action_frame,
            text="Mitarbeiter entfernen",
            command=self._on_remove_employee,
            fg_color=("red3", "red4"),
            hover_color=("red4", "red3"),
            width=170,
            state="disabled"  # Initial deaktiviert, bis ein Mitarbeiter ausgewählt ist
        )
        
        # Layout für Mitarbeiter-Tab
        # Container nimmt den gesamten Tab ein
        container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Projekt-Frame oben
        project_frame.pack(fill="x", pady=(0, 15))
        project_label.pack(side="left", padx=(0, 10))
        self.employee_project_dropdown.pack(side="left")
        
        # Mitarbeiter-Header-Frame
        employee_header_frame.pack(fill="x", pady=(0, 5))
        employee_label.pack(side="left", padx=(0, 10))
        self.employee_count_label.pack(side="left")
        
        # Mitarbeiter-Listbox-Frame
        self.employee_listbox_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Action-Frame unten
        action_frame.pack(fill="x")
        self.add_employee_btn.pack(side="left", padx=(0, 10))
        self.edit_employee_btn.pack(side="left", padx=10)
        self.remove_employee_btn.pack(side="right")
        
        # Initial das erste Projekt auswählen (falls vorhanden)
        if project_names:
            self._on_employee_project_selected(project_names[0])
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Frame
        self.title_frame.pack(fill="x", padx=20, pady=(20, 10))
        self.title_label.pack(pady=10)
        
        # Tab-View (nimmt den meisten Platz ein)
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Button-Frame (unten)
        self.button_frame.pack(fill="x", padx=20, pady=(5, 20))
        self.close_button.pack(side="right", padx=10)
    
    # Event-Handler-Methoden
    
    def _on_project_selected(self, project_name):
        """
        Wird aufgerufen, wenn ein Projekt im Dropdown ausgewählt wird
        
        Args:
            project_name: Der Name des ausgewählten Projekts
        """
        # Finde das Projekt in den Daten
        selected_project = None
        for projekt in self.projects_data.get("projekte", []):
            if projekt["name"] == project_name:
                selected_project = projekt
                break
        
        if not selected_project:
            return
        
        # Aktuelles Projekt speichern
        self.current_project = selected_project
        
        # Projektdetails anzeigen
        self.project_details.configure(state="normal")
        self.project_details.delete("1.0", "end")
        
        # Formatierte Projektdetails erstellen
        details = f"Name: {selected_project['name']}\n\n"
        
        # Anzahl der Mitarbeiter
        mitarbeiter_count = len(selected_project.get("teilnehmer", []))
        details += f"Anzahl Mitarbeiter: {mitarbeiter_count}\n\n"
        
        # Anzahl der Sprints
        sprints = selected_project.get("sprints", [])
        sprint_count = len(sprints)
        details += f"Anzahl Sprints: {sprint_count}\n\n"
        
        # Aktuelle/nächste Sprints
        if sprints:
            now = datetime.datetime.now()
            current_sprint = None
            next_sprint = None
            
            for sprint in sprints:
                start = self._parse_date(sprint["start"])
                ende = self._parse_date(sprint["ende"])
                
                if start <= now <= ende:
                    current_sprint = sprint
                    break
                elif start > now and (next_sprint is None or self._parse_date(next_sprint["start"]) > start):
                    next_sprint = sprint
            
            if current_sprint:
                details += f"Aktueller Sprint: {current_sprint['sprint_name']}\n"
                details += f"  Zeitraum: {current_sprint['start']} - {current_sprint['ende']}\n\n"
            elif next_sprint:
                details += f"Nächster Sprint: {next_sprint['sprint_name']}\n"
                details += f"  Zeitraum: {next_sprint['start']} - {next_sprint['ende']}\n\n"
        
        self.project_details.insert("1.0", details)
        self.project_details.configure(state="disabled")
    
    def _on_sprint_project_selected(self, project_name):
        """
        Wird aufgerufen, wenn ein Projekt im Sprints-Tab ausgewählt wird
        
        Args:
            project_name: Der Name des ausgewählten Projekts
        """
        # Finde das Projekt in den Daten
        selected_project = None
        for projekt in self.projects_data.get("projekte", []):
            if projekt["name"] == project_name:
                selected_project = projekt
                break
        
        if not selected_project:
            return
        
        # Sprint-Dropdown aktualisieren
        sprints = selected_project.get("sprints", [])
        sprint_names = [sprint["sprint_name"] for sprint in sprints]
        
        if not sprint_names:
            sprint_names = ["Kein Sprint"]
        
        self.sprint_dropdown.configure(values=sprint_names)
        self.sprint_var.set(sprint_names[0])
        
        # Den ersten Sprint anzeigen
        if sprint_names and sprint_names[0] != "Kein Sprint":
            self._on_sprint_selected(sprint_names[0])
        else:
            # Keine Sprints vorhanden
            self.sprint_details.configure(state="normal")
            self.sprint_details.delete("1.0", "end")
            self.sprint_details.insert("1.0", "Keine Sprints für dieses Projekt verfügbar.")
            self.sprint_details.configure(state="disabled")
    
    def _on_sprint_selected(self, sprint_name):
        """
        Wird aufgerufen, wenn ein Sprint im Dropdown ausgewählt wird
        
        Args:
            sprint_name: Der Name des ausgewählten Sprints
        """
        if sprint_name == "Kein Sprint":
            return
        
        # Aktuelles Projekt ermitteln
        project_name = self.sprint_project_var.get()
        selected_project = None
        for projekt in self.projects_data.get("projekte", []):
            if projekt["name"] == project_name:
                selected_project = projekt
                break
        
        if not selected_project:
            return
        
        # Sprint finden
        selected_sprint = None
        for sprint in selected_project.get("sprints", []):
            if sprint["sprint_name"] == sprint_name:
                selected_sprint = sprint
                break
        
        if not selected_sprint:
            return
        
        # Aktuellen Sprint speichern
        self.current_sprint = selected_sprint
        
        # Sprintdetails anzeigen
        self.sprint_details.configure(state="normal")
        self.sprint_details.delete("1.0", "end")
        
        # Formatierte Sprintdetails erstellen
        details = f"Sprint: {selected_sprint['sprint_name']}\n"
        
        # Prüfe, ob 'nummer' im Sprint vorhanden ist - wenn nicht, überspringen
        if 'nummer' in selected_sprint:
            details += f"Nummer: {selected_sprint['nummer']}\n"
        
        details += f"Zeitraum: {selected_sprint['start']} - {selected_sprint['ende']}\n\n"
        
        # Story Points
        if 'confimed_story_points' in selected_sprint:
            details += f"Geplante Story Points: {selected_sprint['confimed_story_points']}\n"
        if 'delivered_story_points' in selected_sprint:
            details += f"Gelieferte Story Points: {selected_sprint['delivered_story_points']}\n\n"
        
        # Sprint-Status
        start = self._parse_date(selected_sprint["start"])
        ende = self._parse_date(selected_sprint["ende"])
        now = datetime.datetime.now()
        
        if now < start:
            status = "Geplant"
            days_to_start = (start - now).days
            details += f"Status: {status} (startet in {days_to_start} Tagen)\n"
        elif start <= now <= ende:
            status = "Aktiv"
            days_left = (ende - now).days
            details += f"Status: {status} (noch {days_left} Tage)\n"
        else:
            status = "Abgeschlossen"
            days_since = (now - ende).days
            details += f"Status: {status} (vor {days_since} Tagen)\n"
        
        # Sprint-Dauer in Tagen
        duration = (ende - start).days + 1
        details += f"Dauer: {duration} Tage\n"
        
        self.sprint_details.insert("1.0", details)
        self.sprint_details.configure(state="disabled")
    
    def _on_employee_project_selected(self, project_name):
        """
        Wird aufgerufen, wenn ein Projekt im Mitarbeiter-Tab ausgewählt wird
        
        Args:
            project_name: Der Name des ausgewählten Projekts
        """
        # Finde das Projekt in den Daten
        selected_project = None
        for projekt in self.projects_data.get("projekte", []):
            if projekt["name"] == project_name:
                selected_project = projekt
                break
        
        if not selected_project:
            return
        
        # Aktuelles Projekt speichern
        self.current_project = selected_project
        
        # Mitarbeiterliste leeren
        for widget in self.employee_listbox_frame.winfo_children():
            widget.destroy()
        
        # Bearbeitungs-Buttons zurücksetzen (deaktivieren)
        self.edit_employee_btn.configure(state="disabled")
        self.remove_employee_btn.configure(state="disabled")
        self.selected_employee = None
        
        # Mitarbeiter anzeigen
        mitarbeiter = selected_project.get("teilnehmer", [])
        
        # Anzahl-Label aktualisieren
        self.employee_count_label.configure(text=f"({len(mitarbeiter)} Mitarbeiter)")
        
        if not mitarbeiter:
            # Keine Mitarbeiter vorhanden
            no_employee_label = ctk.CTkLabel(
                self.employee_listbox_frame,
                text="Keine Mitarbeiter für dieses Projekt verfügbar.",
                anchor="w"
            )
            no_employee_label.pack(fill="x", padx=5, pady=5)
            return
        
        # Mitarbeiterliste erstellen
        for i, mitarbeiter_data in enumerate(mitarbeiter):
            mitarbeiter_frame = ctk.CTkFrame(self.employee_listbox_frame)
            mitarbeiter_frame.pack(fill="x", padx=5, pady=2)
            
            # Hintergrundfarbe abwechseln
            if i % 2 == 0:
                mitarbeiter_frame.configure(fg_color=("gray90", "gray20"))
            
            # ID für die Mitarbeiterzeile (für Bearbeiten/Löschen)
            mitarbeiter_frame.mitarbeiter_index = i
            mitarbeiter_frame.mitarbeiter_data = mitarbeiter_data
            
            # Mitarbeitername
            name_label = ctk.CTkLabel(
                mitarbeiter_frame,
                text=mitarbeiter_data["name"],
                anchor="w",
                width=200
            )
            name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
            
            # Mitarbeiterrolle
            role_label = ctk.CTkLabel(
                mitarbeiter_frame,
                text=mitarbeiter_data["rolle"],
                anchor="w",
                width=100
            )
            role_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
            
            # FTE (falls vorhanden)
            if "fte" in mitarbeiter_data:
                fte_label = ctk.CTkLabel(
                    mitarbeiter_frame,
                    text=f"FTE: {mitarbeiter_data['fte']}",
                    anchor="w",
                    width=80
                )
                fte_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
            
            # Ereignisbehandlung für Klick auf Mitarbeiterzeile
            mitarbeiter_frame.bind("<Button-1>", lambda e, f=mitarbeiter_frame: self._on_employee_row_clicked(f))
            name_label.bind("<Button-1>", lambda e, f=mitarbeiter_frame: self._on_employee_row_clicked(f))
            role_label.bind("<Button-1>", lambda e, f=mitarbeiter_frame: self._on_employee_row_clicked(f))
            if "fte" in mitarbeiter_data:
                fte_label.bind("<Button-1>", lambda e, f=mitarbeiter_frame: self._on_employee_row_clicked(f))
    
    def _on_employee_row_clicked(self, row_frame):
        """
        Wird aufgerufen, wenn eine Mitarbeiterzeile angeklickt wird
        
        Args:
            row_frame: Das Frame-Widget der angeklickten Zeile
        """
        # Vorherige Auswahl zurücksetzen
        for widget in self.employee_listbox_frame.winfo_children():
            widget.configure(fg_color=("gray90", "gray20") if widget.winfo_children() and widget.winfo_children()[0].master.mitarbeiter_index % 2 == 0 else "transparent")
        
        # Diese Zeile als ausgewählt markieren
        row_frame.configure(fg_color=("light blue", "dark blue"))
        
        # Mitarbeiterdaten speichern
        self.selected_employee = {
            "index": row_frame.mitarbeiter_index,
            "data": row_frame.mitarbeiter_data
        }
        
        # Bearbeitungs-Buttons aktivieren
        self.edit_employee_btn.configure(state="normal")
        self.remove_employee_btn.configure(state="normal")
    
    def _on_add_employee(self):
        """Wird aufgerufen, wenn der 'Mitarbeiter hinzufügen' Button geklickt wird"""
        if self.current_project is None:
            messagebox.showwarning(
                "Kein Projekt ausgewählt",
                "Bitte wählen Sie zuerst ein Projekt aus.",
                parent=self
            )
            return
        
        # Dialog zum Hinzufügen eines Mitarbeiters anzeigen
        dialog = MitarbeiterDialog(
            self,
            title="Mitarbeiter hinzufügen",
            on_save=self._add_employee_callback
        )
        
        # Warten, bis der Dialog geschlossen wird
        self.wait_window(dialog)
    
    def _add_employee_callback(self, mitarbeiter_data):
        """
        Callback-Funktion für den Mitarbeiter-Dialog (Hinzufügen)
        
        Args:
            mitarbeiter_data: Die Daten des neuen Mitarbeiters
        """
        if not self.current_project:
            return
        
        # Stelle sicher, dass die Liste existiert
        if "teilnehmer" not in self.current_project:
            self.current_project["teilnehmer"] = []
        
        # Füge den neuen Mitarbeiter hinzu
        self.current_project["teilnehmer"].append(mitarbeiter_data)
        
        # Speichere die Projektdaten
        if self._save_project_data():
            # Aktualisiere die Anzeige
            self._on_employee_project_selected(self.current_project["name"])
            
            # Zeige Erfolgsmeldung
            messagebox.showinfo(
                "Mitarbeiter hinzugefügt",
                f"Der Mitarbeiter '{mitarbeiter_data['name']}' wurde erfolgreich hinzugefügt.",
                parent=self
            )
    
    def _on_edit_employee(self):
        """Wird aufgerufen, wenn der 'Mitarbeiter bearbeiten' Button geklickt wird"""
        if not self.selected_employee:
            return
        
        # Dialog zum Bearbeiten eines Mitarbeiters anzeigen
        dialog = MitarbeiterDialog(
            self,
            title="Mitarbeiter bearbeiten",
            on_save=self._edit_employee_callback,
            mitarbeiter_data=self.selected_employee["data"]
        )
        
        # Warten, bis der Dialog geschlossen wird
        self.wait_window(dialog)
    
    def _edit_employee_callback(self, mitarbeiter_data):
        """
        Callback-Funktion für den Mitarbeiter-Dialog (Bearbeiten)
        
        Args:
            mitarbeiter_data: Die aktualisierten Daten des Mitarbeiters
        """
        if not self.current_project or not self.selected_employee:
            return
        
        # Aktualisiere den Mitarbeiter
        index = self.selected_employee["index"]
        self.current_project["teilnehmer"][index] = mitarbeiter_data
        
        # Speichere die Projektdaten
        if self._save_project_data():
            # Aktualisiere die Anzeige
            self._on_employee_project_selected(self.current_project["name"])
            
            # Zeige Erfolgsmeldung
            messagebox.showinfo(
                "Mitarbeiter aktualisiert",
                f"Der Mitarbeiter '{mitarbeiter_data['name']}' wurde erfolgreich aktualisiert.",
                parent=self
            )
    
    def _on_remove_employee(self):
        """Wird aufgerufen, wenn der 'Mitarbeiter entfernen' Button geklickt wird"""
        if not self.selected_employee:
            return
        
        # Bestätigungsdialog anzeigen
        if not messagebox.askyesno(
            "Mitarbeiter entfernen",
            f"Möchten Sie den Mitarbeiter '{self.selected_employee['data']['name']}' wirklich entfernen?",
            parent=self
        ):
            return
        
        # Mitarbeiter entfernen
        index = self.selected_employee["index"]
        mitarbeiter_name = self.current_project["teilnehmer"][index]["name"]
        del self.current_project["teilnehmer"][index]
        
        # Speichere die Projektdaten
        if self._save_project_data():
            # Aktualisiere die Anzeige
            self._on_employee_project_selected(self.current_project["name"])
            
            # Zeige Erfolgsmeldung
            messagebox.showinfo(
                "Mitarbeiter entfernt",
                f"Der Mitarbeiter '{mitarbeiter_name}' wurde erfolgreich entfernt.",
                parent=self
            )
    
    def _on_new_project(self):
        """Wird aufgerufen, wenn der 'Neues Projekt' Button geklickt wird"""
        if "new_project" in self.callbacks:
            self.callbacks["new_project"]()
    
    def _on_edit_project(self):
        """Wird aufgerufen, wenn der 'Projekt bearbeiten' Button geklickt wird"""
        if "edit_project" in self.callbacks:
            self.callbacks["edit_project"]()
    
    def _on_delete_project(self):
        """Wird aufgerufen, wenn der 'Projekt löschen' Button geklickt wird"""
        if "delete_project" in self.callbacks:
            self.callbacks["delete_project"]()
    
    def _on_new_sprint(self):
        """Wird aufgerufen, wenn der 'Neuer Sprint' Button geklickt wird"""
        if "new_sprint" in self.callbacks:
            self.callbacks["new_sprint"]()
    
    def _on_edit_sprint(self):
        """Wird aufgerufen, wenn der 'Sprint bearbeiten' Button geklickt wird"""
        if "edit_sprint" in self.callbacks:
            self.callbacks["edit_sprint"]()
    
    def _on_delete_sprint(self):
        """Wird aufgerufen, wenn der 'Sprint löschen' Button geklickt wird"""
        # Implementiere die Sprint-Löschfunktion
        pass
    
    def _parse_date(self, date_str):
        """
        Wandelt einen Datums-String in ein datetime-Objekt um
        
        Args:
            date_str: Datums-String im Format "DD.MM.YYYY"
            
        Returns:
            datetime: Das entsprechende datetime-Objekt
        """
        try:
            day, month, year = map(int, date_str.split('.'))
            return datetime.datetime(year, month, day)
        except (ValueError, AttributeError):
            # Fallback: Aktuelles Datum zurückgeben
            return datetime.datetime.now()
    
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