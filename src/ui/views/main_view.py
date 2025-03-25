"""
Hauptansicht für CAMP
Zeigt Projektsprints, Sprintdaten und Mitarbeiterauslastung an
"""
import os
import json
import csv
import customtkinter as ctk
from datetime import datetime, timedelta
import calendar
import pandas as pd

class MainView(ctk.CTkFrame):
    """
    Hauptansicht der Anwendung, zeigt Projektsprints, Sprintdaten und Mitarbeiterauslastung an
    """
    def __init__(self, master):
        """
        Initialisiert die Hauptansicht
        
        Args:
            master: Das übergeordnete Widget
        """
        super().__init__(master)
        
        # Attribute
        self.projects_data = self._load_project_data()
        self.selected_project = None
        self.selected_sprint = None
        self.kapa_data = self._load_kapa_data()
        
        # Erstelle die UI-Elemente
        self._create_widgets()
        
        # Richte das Layout ein
        self._setup_layout()
        
        # Lade initial das erste Projekt, falls vorhanden
        self._load_initial_project()
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für die Hauptansicht"""
        # Oberer Frame für Projekt-Auswahl
        self.top_frame = ctk.CTkFrame(self)
        
        # Projekt-Label
        self.project_label = ctk.CTkLabel(
            self.top_frame,
            text="Projekt:",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=80
        )
        
        # Projekt-Auswahl
        project_names = [p.get("name", "") for p in self.projects_data.get("projekte", [])]
        self.project_var = ctk.StringVar(value=project_names[0] if project_names else "Kein Projekt")
        
        self.project_dropdown = ctk.CTkOptionMenu(
            self.top_frame,
            values=project_names if project_names else ["Kein Projekt"],
            variable=self.project_var,
            command=self._on_project_selected,
            width=300
        )
        
        # Haupt-Container (3 Spalten: Sprints, Sprint-Details, Team)
        self.main_container = ctk.CTkFrame(self)
        
        # Links: Sprint-Liste
        self.sprint_frame = ctk.CTkFrame(self.main_container)
        
        self.sprint_label = ctk.CTkLabel(
            self.sprint_frame,
            text="Sprints",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        self.sprint_list_frame = ctk.CTkScrollableFrame(
            self.sprint_frame
        )
        
        # Mitte: Sprint-Details
        self.details_frame = ctk.CTkFrame(self.main_container)
        
        self.details_label = ctk.CTkLabel(
            self.details_frame,
            text="Sprint Details",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        self.sprint_details = ctk.CTkTextbox(
            self.details_frame
        )
        self.sprint_details.configure(state="disabled")
        
        # Rechts: Team-Mitglieder und Auslastung
        self.team_frame = ctk.CTkFrame(self.main_container)
        
        self.team_label = ctk.CTkLabel(
            self.team_frame,
            text="Team Auslastung",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        self.team_list_frame = ctk.CTkScrollableFrame(
            self.team_frame
        )

        # Kalender-Container (unten)
        self.calendar_container = ctk.CTkFrame(self)

        self.calendar_label = ctk.CTkLabel(
            self.calendar_container,
            text="Sprint-Tageskalender",
            font=ctk.CTkFont(size=16, weight="bold")
        )

        # Calendar-Frame direkt
        self.calendar_frame = ctk.CTkFrame(self.calendar_container)
        
        # Statusleiste am unteren Rand
        self.status_bar = ctk.CTkFrame(self, height=25)
        self.status_label = ctk.CTkLabel(self.status_bar, text="Bereit", anchor="w")

    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Haupt-Layout: Vertikales Grid mit 2 Zeilen gleicher Höhe
        self.grid_rowconfigure(0, weight=0)  # Für top_frame
        self.grid_rowconfigure(1, weight=1)  # Für main_container
        self.grid_rowconfigure(2, weight=1)  # Für calendar_container
        self.grid_rowconfigure(3, weight=0)  # Für status_bar
        self.grid_columnconfigure(0, weight=1)  # Eine Spalte für alles
        
        # Oberer Frame (Projekt-Auswahl)
        self.top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.project_label.pack(side="left", padx=(10, 5))
        self.project_dropdown.pack(side="left", padx=5)
        
        # Haupt-Container (obere Hälfte - Sprint-Liste, Sprint-Details, Team)
        self.main_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        # Spalten im Haupt-Container konfigurieren
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(1, weight=2)
        self.main_container.grid_columnconfigure(2, weight=1)
        self.main_container.grid_rowconfigure(0, weight=1)  # Zeile soll sich ausdehnen
        
        # Links: Sprint-Liste
        self.sprint_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.sprint_label.pack(anchor="n", pady=10)
        self.sprint_list_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Mitte: Sprint-Details
        self.details_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        self.details_label.pack(anchor="n", pady=10)
        self.sprint_details.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Rechts: Team-Mitglieder
        self.team_frame.grid(row=0, column=2, sticky="nsew", padx=5, pady=5)
        self.team_label.pack(anchor="n", pady=10)
        self.team_list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Kalender-Container (untere Hälfte)
        self.calendar_container.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self.calendar_container.grid_rowconfigure(0, weight=0)  # Für den Titel
        self.calendar_container.grid_rowconfigure(1, weight=1)  # Für den Kalender
        self.calendar_container.grid_columnconfigure(0, weight=1)
        
        self.calendar_label.grid(row=0, column=0, sticky="nw", padx=5, pady=5)
        self.calendar_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        # Statusleiste am unteren Rand
        self.status_bar.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        self.status_label.pack(side="left", padx=10, pady=2, fill="x", expand=True)
    
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
                return {"projekte": []}
        except Exception as e:
            print(f"Fehler beim Laden der Projektdaten: {e}")
            return {"projekte": []}
    
    def _load_kapa_data(self):
        """
        Lädt die Kapazitätsdaten aus der kapa_data.csv-Datei
        
        Returns:
            list: Liste von Tupeln (id, datum, stunden, kapazität)
        """
        data = []
        try:
            csv_path = os.path.join("data", "kapa_data.csv")
            if os.path.exists(csv_path):
                with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=';')
                    for row in reader:
                        if len(row) >= 3:
                            # Stunden als Float
                            stunden = float(row[2].replace(',', '.'))
                            
                            # Kapazität aus der Datei lesen oder berechnen, falls nicht vorhanden
                            if len(row) >= 4:
                                kapazitaet = float(row[3].replace(',', '.'))
                            else:
                                # Kapazität berechnen: Stunden >= 8 entsprechen Kapazität 1
                                kapazitaet = 1.0 if stunden >= 8 else round(stunden / 8, 2)
                            
                            data.append((row[0], row[1], stunden, kapazitaet))
        except Exception as e:
            print(f"Fehler beim Laden der Kapazitätsdaten: {e}")
        
        return data
    
    def _load_initial_project(self):
        """Lädt das erste Projekt, falls vorhanden"""
        if self.projects_data and "projekte" in self.projects_data and self.projects_data["projekte"]:
            first_project = self.projects_data["projekte"][0]
            self._on_project_selected(first_project["name"])
    
    def _on_project_selected(self, project_name):
        """
        Wird aufgerufen, wenn ein Projekt im Dropdown ausgewählt wird
        
        Args:
            project_name: Der Name des ausgewählten Projekts
        """
        # Projekt finden
        self.selected_project = None
        for project in self.projects_data.get("projekte", []):
            if project.get("name") == project_name:
                self.selected_project = project
                break
        
        if not self.selected_project:
            self.show_message(f"Projekt '{project_name}' nicht gefunden")
            return
        
        # Sprintliste aktualisieren
        self._update_sprint_list()
        
        # Status aktualisieren
        self.show_message(f"Projekt '{project_name}' geladen")
    
    def _update_sprint_list(self):
        """Aktualisiert die Liste der Sprints für das aktuelle Projekt"""
        # Lösche alle vorhandenen Widgets
        for widget in self.sprint_list_frame.winfo_children():
            widget.destroy()
        
        # Keine Sprints vorhanden
        if not self.selected_project or "sprints" not in self.selected_project or not self.selected_project["sprints"]:
            no_sprints_label = ctk.CTkLabel(
                self.sprint_list_frame,
                text="Keine Sprints vorhanden",
                font=ctk.CTkFont(size=12, slant="italic")
            )
            no_sprints_label.pack(pady=20)
            
            # Sprint-Details und Team leeren
            self._clear_sprint_details()
            self._clear_team_list()
            self._clear_calendar()
            return
        
        # Aktuelles Datum für Vergleich
        now = datetime.now()
        
        # Sortiere Sprints nach Startdatum
        sprints = sorted(
            self.selected_project["sprints"],
            key=lambda s: self._parse_date(s.get("start", "01.01.2000"))
        )
        
        # Aktiven Sprint finden
        active_sprint = None
        for sprint in sprints:
            start_date = self._parse_date(sprint.get("start", "01.01.2000"))
            end_date = self._parse_date(sprint.get("ende", "01.01.2000"))
            
            if start_date <= now <= end_date:
                active_sprint = sprint
                break
        
        # Sprints darstellen
        for i, sprint in enumerate(sprints):
            sprint_name = sprint.get("sprint_name", f"Sprint {i+1}")
            start_date = self._parse_date(sprint.get("start", "01.01.2000"))
            end_date = self._parse_date(sprint.get("ende", "01.01.2000"))
            
            # Rahmen für Sprint
            sprint_frame = ctk.CTkFrame(self.sprint_list_frame)
            sprint_frame.pack(fill="x", padx=5, pady=3)
            
            # Sprint-Status bestimmen (vergangen, aktiv, zukünftig)
            if start_date <= now <= end_date:
                # Aktiver Sprint
                status = "Aktiv"
                status_color = "green"
                bg_color = ("light green", "dark green")
            elif end_date < now:
                # Vergangener Sprint
                status = "Vergangen"
                status_color = "gray"
                bg_color = ("gray90", "gray20")
            else:
                # Zukünftiger Sprint
                status = "Geplant"
                status_color = "blue"
                bg_color = ("gray80", "gray30")
            
            # Hintergrundfarbe für Sprint-Rahmen
            sprint_frame.configure(fg_color=bg_color)
            
            # Sprint-Label
            sprint_label = ctk.CTkLabel(
                sprint_frame,
                text=sprint_name,
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w"
            )
            sprint_label.pack(fill="x", padx=10, pady=(5, 2))
            
            # Datum-Label
            date_label = ctk.CTkLabel(
                sprint_frame,
                text=f"{sprint.get('start', '')} - {sprint.get('ende', '')}",
                font=ctk.CTkFont(size=10),
                anchor="w"
            )
            date_label.pack(fill="x", padx=10, pady=(0, 2))
            
            # Status-Label
            status_label = ctk.CTkLabel(
                sprint_frame,
                text=f"Status: {status}",
                font=ctk.CTkFont(size=10),
                text_color=status_color,
                anchor="w"
            )
            status_label.pack(fill="x", padx=10, pady=(0, 5))
            
            # Event-Binding für Klicks
            sprint_frame.bind("<Button-1>", lambda e, s=sprint: self._on_sprint_selected(s))
            sprint_label.bind("<Button-1>", lambda e, s=sprint: self._on_sprint_selected(s))
            date_label.bind("<Button-1>", lambda e, s=sprint: self._on_sprint_selected(s))
            status_label.bind("<Button-1>", lambda e, s=sprint: self._on_sprint_selected(s))
        
        # Wähle den aktiven Sprint aus, wenn vorhanden, sonst den neuesten
        if active_sprint:
            self._on_sprint_selected(active_sprint)
        elif sprints:
            self._on_sprint_selected(sprints[-1])
    
    def _on_sprint_selected(self, sprint):
        """
        Wird aufgerufen, wenn ein Sprint ausgewählt wird
        
        Args:
            sprint: Die Daten des ausgewählten Sprints
        """
        self.selected_sprint = sprint
        
        # Sprint-Details anzeigen
        self._update_sprint_details()
        
        # Team-Auslastung aktualisieren
        self._update_team_list()
        
        # Kalender aktualisieren
        self._update_calendar()
        
        # Status aktualisieren
        self.show_message(f"Sprint '{sprint.get('sprint_name', 'Unbenannt')}' ausgewählt")
    
    def _update_sprint_details(self):
        """Aktualisiert die Anzeige der Sprint-Details im Dashboard-Stil"""
        # Bestehenden Inhalt im Sprint-Details-Frame löschen
        for widget in self.details_frame.winfo_children():
            if widget != self.details_label:  # Den Titel behalten
                widget.destroy()
        
        # Container für das Dashboard erstellen
        self.sprint_details = ctk.CTkScrollableFrame(
            self.details_frame
        )
        self.sprint_details.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Prüfen, ob ein Sprint ausgewählt ist
        if not self.selected_sprint:
            no_sprint_label = ctk.CTkLabel(
                self.sprint_details,
                text="Kein Sprint ausgewählt",
                font=ctk.CTkFont(size=16, slant="italic")
            )
            no_sprint_label.pack(pady=50)
            return
        
        # Sprint-Daten extrahieren
        sprint_name = self.selected_sprint.get("sprint_name", "Unbenannt")
        start_date = self.selected_sprint.get("start", "")
        end_date = self.selected_sprint.get("ende", "")
        
        # Kapazitäten berechnen - nur für Teammitglieder
        sprint_data = self._get_sprint_capacity_data()
        total_capacity = sum(capacity for _, _, capacity in sprint_data)
        
        # Fester Umrechnungsfaktor
        umrechnungsfaktor = 1.4
        
        # Berechnete Story Points
        berechnete_sp = total_capacity / umrechnungsfaktor if total_capacity > 0 else 0
        
        # Header mit Sprint-Namen
        header_frame = ctk.CTkFrame(self.sprint_details, fg_color=("blue", "dark blue"), corner_radius=8)
        header_frame.pack(fill="x", pady=(0, 15))
        
        sprint_title = ctk.CTkLabel(
            header_frame,
            text=sprint_name,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        sprint_title.pack(pady=15)
        
        # Zeitraum-Banner
        timeline_frame = ctk.CTkFrame(self.sprint_details, corner_radius=8)
        timeline_frame.pack(fill="x", pady=(0, 15))
        
        timeline_label = ctk.CTkLabel(
            timeline_frame,
            text=f"Zeitraum: {start_date} bis {end_date}",
            font=ctk.CTkFont(size=14),
            anchor="center"
        )
        timeline_label.pack(pady=10)
        
        # Container für drei Karten nebeneinander
        cards_frame = ctk.CTkFrame(self.sprint_details)
        cards_frame.pack(fill="x", pady=(0, 10))
        
        # Grid-Konfiguration für gleichmäßige Spaltenbreite
        cards_frame.columnconfigure(0, weight=1)
        cards_frame.columnconfigure(1, weight=1)
        cards_frame.columnconfigure(2, weight=1)
        
        # 1. Karte: Kapazität
        capacity_card = self._create_card(
            cards_frame,
            "Kapazität",
            f"{total_capacity:.2f}",
            "Teamkapazität für den Sprint"
        )
        capacity_card.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # 2. Karte: Umrechnungsfaktor
        factor_card = self._create_card(
            cards_frame,
            "Faktor",
            f"{umrechnungsfaktor}",
            "Umrechnungsfaktor"
        )
        factor_card.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        # 3. Karte: Story Points
        sp_card = self._create_card(
            cards_frame,
            "Story Points",
            f"{berechnete_sp:.1f}",
            "Berechnete SP (Kapazität/Faktor)"
        )
        sp_card.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")

        
    def _clear_sprint_details(self):
        """Leert die Sprint-Details"""
        # Bestehenden Inhalt im Sprint-Details-Frame löschen (außer dem Titel)
        for widget in self.details_frame.winfo_children():
            if widget != self.details_label:  # Den Titel behalten
                widget.destroy()
        
        # Container für das Dashboard erstellen
        self.sprint_details = ctk.CTkScrollableFrame(
            self.details_frame
        )
        self.sprint_details.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Hinweis anzeigen
        no_sprint_label = ctk.CTkLabel(
            self.sprint_details,
            text="Kein Sprint ausgewählt",
            font=ctk.CTkFont(size=16, slant="italic")
        )
        no_sprint_label.pack(pady=50)
    
    def _update_team_list(self):
        """Aktualisiert die Liste der Team-Mitglieder und deren Auslastung"""
        # Lösche alle vorhandenen Widgets
        for widget in self.team_list_frame.winfo_children():
            widget.destroy()
        
        if not self.selected_project or not self.selected_sprint:
            no_team_label = ctk.CTkLabel(
                self.team_list_frame,
                text="Keine Daten verfügbar",
                font=ctk.CTkFont(size=12, slant="italic")
            )
            no_team_label.pack(pady=20)
            return
        
        # Team-Mitglieder
        team_members = self.selected_project.get("teilnehmer", [])
        
        if not team_members:
            no_team_label = ctk.CTkLabel(
                self.team_list_frame,
                text="Keine Team-Mitglieder vorhanden",
                font=ctk.CTkFont(size=12, slant="italic")
            )
            no_team_label.pack(pady=20)
            return
        
        # Kapazitätsdaten für den Sprint laden
        capacity_data = {}
        kapa_data = {}
        for member_id, hours, capacity in self._get_sprint_capacity_data():
            capacity_data[member_id] = hours
            kapa_data[member_id] = capacity
        
        # Container für die Tabelle
        table_container = ctk.CTkFrame(self.team_list_frame)
        table_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Konfiguriere das Grid-Layout mit dynamischen Spaltenbreiten
        # Mindestbreiten und Gewichtungen für jede Spalte
        col_configs = [
            {"minsize": 150, "weight": 3},  # Name (breiteste Spalte)
            {"minsize": 90, "weight": 2},   # Rolle (mittlere Breite)
            {"minsize": 60, "weight": 1},   # Stunden (schmalste Spalte)
            {"minsize": 70, "weight": 1}    # Kapazität (schmalste Spalte)
        ]
        
        # Spalten-Konfiguration auf den Hauptcontainer anwenden
        for i, config in enumerate(col_configs):
            table_container.grid_columnconfigure(
                i, 
                minsize=config["minsize"], 
                weight=config["weight"]
            )
        
        # Header-Labels im Hauptcontainer
        name_header = ctk.CTkLabel(
            table_container,
            text="Name",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        name_header.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        role_header = ctk.CTkLabel(
            table_container,
            text="Rolle",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        role_header.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        hours_header = ctk.CTkLabel(
            table_container,
            text="Stunden",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="e"
        )
        hours_header.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        
        capacity_header = ctk.CTkLabel(
            table_container,
            text="Kapazität",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="e"
        )
        capacity_header.grid(row=0, column=3, padx=5, pady=5, sticky="e")
        
        # Trennlinie unter den Überschriften
        separator = ctk.CTkFrame(table_container, height=1, fg_color=("gray70", "gray30"))
        separator.grid(row=1, column=0, columnspan=4, sticky="ew", padx=5, pady=2)
        
        # Team-Mitglieder anzeigen
        for i, member in enumerate(team_members):
            row_idx = i + 2  # +2 wegen Header und Separator
            member_id = member.get("id", "")
            member_name = member.get("name", "Unbekannt")
            member_role = member.get("rolle", "")
            
            # Kapazität für diesen Mitarbeiter
            hours = capacity_data.get(member_id, 0)
            capacity = kapa_data.get(member_id, 0)
            
            # Hintergrundfarbe für die Zeile
            bg_color = ("gray90", "gray20") if i % 2 == 1 else None
            
            # Zeilencontainer
            row_frame = ctk.CTkFrame(table_container, fg_color=bg_color, corner_radius=0)
            row_frame.grid(row=row_idx, column=0, columnspan=4, sticky="ew", padx=0, pady=0)
            
            # Die gleiche Grid-Konfiguration für die Zeile
            for j, config in enumerate(col_configs):
                row_frame.grid_columnconfigure(
                    j, 
                    minsize=config["minsize"], 
                    weight=config["weight"]
                )
            
            # Name (linksbündig)
            name_label = ctk.CTkLabel(
                row_frame,
                text=member_name,
                anchor="w"
            )
            name_label.grid(row=0, column=0, padx=5, pady=3, sticky="ew")
            
            # Rolle (linksbündig)
            role_label = ctk.CTkLabel(
                row_frame,
                text=member_role,
                anchor="w"
            )
            role_label.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
            
            # Stunden (rechtsbündig)
            hours_label = ctk.CTkLabel(
                row_frame,
                text=f"{hours:.1f}h",
                anchor="e"
            )
            hours_label.grid(row=0, column=2, padx=5, pady=3, sticky="e")
            
            # Kapazität (rechtsbündig)
            capacity_label = ctk.CTkLabel(
                row_frame,
                text=f"{capacity:.2f}",
                anchor="e"
            )
            capacity_label.grid(row=0, column=3, padx=5, pady=3, sticky="e")

    def _update_calendar(self):
        """Aktualisiert den Tageskalender für den ausgewählten Sprint"""
        # Bestehenden Kalender leeren
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        # Überprüfen, ob ein Sprint ausgewählt ist
        if not self.selected_project or not self.selected_sprint:
            no_calendar_label = ctk.CTkLabel(
                self.calendar_frame,
                text="Kein Sprint ausgewählt",
                font=ctk.CTkFont(size=12, slant="italic")
            )
            no_calendar_label.pack(pady=20)
            return
        
        # Status aktualisieren
        self.show_message("Kalender wird aktualisiert...")
        
        # Sprint-Zeitraum ermitteln
        start_date = self._parse_date(self.selected_sprint.get("start", "01.01.2000"))
        end_date = self._parse_date(self.selected_sprint.get("ende", "01.01.2000"))
        
        # Alle Tage zwischen Start und Ende berechnen
        current_date = start_date
        sprint_days = []
        while current_date <= end_date:
            sprint_days.append(current_date)
            current_date += timedelta(days=1)
        
        # Team-Mitglieder
        team_members = self.selected_project.get("teilnehmer", [])
        
        if not team_members or not sprint_days:
            no_calendar_label = ctk.CTkLabel(
                self.calendar_frame,
                text="Keine Daten für den Kalender verfügbar",
                font=ctk.CTkFont(size=12, slant="italic")
            )
            no_calendar_label.pack(pady=20)
            return
        
        # Liste der Mitarbeiter-IDs im Projektteam erstellen
        team_member_ids = set()
        for member in team_members:
            member_id = member.get("id", "")
            if member_id:  # Nur gültige IDs hinzufügen
                team_member_ids.add(member_id)
        
        # Kapazitätsdaten nach Datum und Mitarbeiter organisieren
        daily_capacity = {}
        for member_id, datum_str, hours, capacity in self.kapa_data:
            # Nur Einträge für Teammitglieder berücksichtigen
            if member_id not in team_member_ids:
                continue
                
            try:
                datum = self._parse_date(datum_str)
                
                # Prüfen, ob der Tag im Sprint-Zeitraum liegt
                if start_date <= datum <= end_date:
                    key = (member_id, datum.strftime('%d.%m.%Y'))
                    daily_capacity[key] = capacity
            except:
                # Bei Fehlern beim Datum-Parsing überspringen
                continue
        
        # Scrollbaren Container für den Kalender erstellen
        # Nutzt die gesamte verfügbare Fläche
        calendar_scroll = ctk.CTkScrollableFrame(
            self.calendar_frame
        )
        calendar_scroll.pack(fill="both", expand=True)
        
        # Tabelle erstellen
        calendar_table = ctk.CTkFrame(calendar_scroll)
        calendar_table.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Konfiguriere die Spalten, damit sie sich der Fenstergröße anpassen
        # Die erste Spalte (Mitarbeiternamen) sollte eine feste Breite haben
        calendar_table.grid_columnconfigure(0, weight=0, minsize=150)
        
        # Die restlichen Spalten (Tage) sollen sich gleichmäßig den verbleibenden Platz teilen
        for col in range(1, len(sprint_days) + 1):
            calendar_table.grid_columnconfigure(col, weight=1)
        
        # Zellen-Höhe
        cell_height = 30
        
        # Spalten-Header (Tage)
        # Erste Zelle ist leer (Ecke oben links)
        empty_header = ctk.CTkLabel(
            calendar_table,
            text="",
            height=cell_height
        )
        empty_header.grid(row=0, column=0, padx=1, pady=1, sticky="nsew")
        
        # Tagesnamen und Daten als Header
        for col, day in enumerate(sprint_days, 1):
            # Prüfen, ob der Tag ein Wochenende ist
            is_weekend = day.weekday() >= 5  # 5=Samstag, 6=Sonntag
            
            # Zellenfarbe für Wochenenden
            bg_color = ("gray80", "gray40") if is_weekend else None
            
            # Tagesnamen erstellen (z.B. "Mo 01.01.")
            day_name = day.strftime('%a %d.%m.')
            
            # Header-Label
            day_header = ctk.CTkLabel(
                calendar_table,
                text=day_name,
                height=cell_height,
                fg_color=bg_color
            )
            day_header.grid(row=0, column=col, padx=1, pady=1, sticky="nsew")
        
        # Zeilen-Header (Mitarbeiter) und Kapazitätszellen
        for row, member in enumerate(team_members, 1):
            member_id = member.get("id", "")
            member_name = member.get("name", "Unbekannt")
            
            # Mitarbeiter-Namen
            name_header = ctk.CTkLabel(
                calendar_table,
                text=member_name,
                height=cell_height,
                anchor="w",
                font=ctk.CTkFont(weight="bold")
            )
            name_header.grid(row=row, column=0, padx=1, pady=1, sticky="nsew")
            
            # Kapazitätszellen für jeden Tag
            for col, day in enumerate(sprint_days, 1):
                # Prüfen, ob der Tag ein Wochenende ist
                is_weekend = day.weekday() >= 5  # 5=Samstag, 6=Sonntag
                
                # Wenn Wochenende, dann graue Zelle
                if is_weekend:
                    cell = ctk.CTkLabel(
                        calendar_table,
                        text="WE",
                        height=cell_height,
                        fg_color=("gray80", "gray40")
                    )
                    cell.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
                    continue
                
                # Kapazität für diesen Tag und Mitarbeiter suchen
                key = (member_id, day.strftime('%d.%m.%Y'))
                capacity = daily_capacity.get(key, None)
                
                # Farbe basierend auf Kapazität
                if capacity is None:
                    # Kein Eintrag: Rot
                    bg_color = ("red", "dark red")
                    text = "X"
                elif capacity >= 1.0:
                    # Kapazität 1 oder mehr: Grün
                    bg_color = ("green", "dark green")
                    text = f"{capacity:.1f}"
                else:
                    # Kapazität unter 1: Gelb
                    bg_color = ("yellow", "#B8860B")  # Dunkleres Gelb für dunklen Modus
                    text = f"{capacity:.1f}"
                
                # Zelle erstellen
                cell = ctk.CTkLabel(
                    calendar_table,
                    text=text,
                    height=cell_height,
                    fg_color=bg_color
                )
                cell.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
        
        # Status aktualisieren
        self.show_message(f"Kalender für Sprint '{self.selected_sprint.get('sprint_name', 'Unbenannt')}' aktualisiert")

    
    def _get_sprint_capacity_data(self):
        """
        Ermittelt die Kapazitätsdaten für den ausgewählten Sprint,
        aber nur für Mitglieder des aktuellen Projektteams
        
        Returns:
            list: Liste von Tupeln (member_id, hours, capacity)
        """
        if not self.selected_sprint or not self.selected_project:
            return []
        
        # Sprint-Zeitraum
        start_date = self._parse_date(self.selected_sprint.get("start", "01.01.2000"))
        end_date = self._parse_date(self.selected_sprint.get("ende", "01.01.2000"))
        
        # Liste der Mitarbeiter-IDs im Projektteam erstellen
        team_member_ids = set()
        for member in self.selected_project.get("teilnehmer", []):
            member_id = member.get("id", "")
            if member_id:  # Nur gültige IDs hinzufügen
                team_member_ids.add(member_id)
        
        # Kapazitätsdaten filtern - nur für Teammitglieder
        member_hours = {}
        member_capacity = {}
        
        for member_id, datum_str, hours, capacity in self.kapa_data:
            # Nur Einträge für Teammitglieder berücksichtigen
            if member_id not in team_member_ids:
                continue
                
            try:
                # Datumstring in Datum umwandeln
                datum = self._parse_date(datum_str)
                
                # Prüfen, ob das Datum im Sprint-Zeitraum liegt
                if start_date <= datum <= end_date:
                    # Stunden für Mitarbeiter summieren
                    if member_id in member_hours:
                        member_hours[member_id] += hours
                        member_capacity[member_id] += capacity
                    else:
                        member_hours[member_id] = hours
                        member_capacity[member_id] = capacity
            except:
                # Bei Fehlern beim Datum-Parsing überspringen
                continue
        
        # Als Liste von Tupeln zurückgeben
        return [(member_id, member_hours[member_id], member_capacity[member_id]) 
                for member_id in member_hours.keys()]
    
    def _clear_sprint_details(self):
        """Leert die Sprint-Details"""
        self.sprint_details.configure(state="normal")
        self.sprint_details.delete("1.0", "end")
        self.sprint_details.insert("1.0", "Kein Sprint ausgewählt")
        self.sprint_details.configure(state="disabled")
    
    def _clear_team_list(self):
        """Leert die Team-Liste"""
        for widget in self.team_list_frame.winfo_children():
            widget.destroy()
        
        no_team_label = ctk.CTkLabel(
            self.team_list_frame,
            text="Kein Sprint ausgewählt",
            font=ctk.CTkFont(size=12, slant="italic")
        )
        no_team_label.pack(pady=20)
    
    def _clear_calendar(self):
        """Leert den Kalender"""
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
    
    def _parse_date(self, date_str):
        """
        Wandelt einen Datums-String in ein datetime-Objekt um
        
        Args:
            date_str: Datums-String im Format DD.MM.YYYY
            
        Returns:
            datetime: Das entsprechende datetime-Objekt
        """
        try:
            if not date_str:
                return datetime.now()
            
            day, month, year = map(int, date_str.split('.'))
            return datetime(year, month, day)
        except:
            # Fallback: Aktuelles Datum
            return datetime.now()
    
    def refresh_data(self):
        """Aktualisiert alle Daten"""
        # Projektdaten neu laden
        self.projects_data = self._load_project_data()
        
        # Kapazitätsdaten neu laden
        self.kapa_data = self._load_kapa_data()
        
        # Projekt-Dropdown aktualisieren
        project_names = [p.get("name", "") for p in self.projects_data.get("projekte", [])]
        self.project_dropdown.configure(values=project_names if project_names else ["Kein Projekt"])
        
        if project_names:
            current_project = self.project_var.get()
            if current_project not in project_names:
                # Aktuelles Projekt nicht mehr vorhanden
                self.project_var.set(project_names[0])
                self._on_project_selected(project_names[0])
            else:
                # Aktuelles Projekt beibehalten, aber Daten aktualisieren
                self._on_project_selected(current_project)
        else:
            # Keine Projekte vorhanden
            self.project_var.set("Kein Projekt")
            self._clear_sprint_details()
            self._clear_team_list()
            self._clear_calendar()
        
        # Status aktualisieren
        self.show_message("Daten aktualisiert")
    
    def show_message(self, message: str):
        """
        Zeigt eine Nachricht in der Statusleiste an
        
        Args:
            message: Die anzuzeigende Nachricht
        """
        self.status_label.configure(text=message)

    def _create_card(self, parent, title, value, description=None):
        """Erstellt eine Dashboard-Karte mit Titel, Wert und optionaler Beschreibung"""
        card = ctk.CTkFrame(parent, corner_radius=10)
        
        # Titel
        title_label = ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        title_label.pack(anchor="n", pady=(15, 5))
        
        # Wert (groß)
        value_label = ctk.CTkLabel(
            card,
            text=value,
            font=ctk.CTkFont(size=24, weight="bold")
        )
        value_label.pack(pady=10)
        
        # Beschreibung (optional)
        if description:
            desc_label = ctk.CTkLabel(
                card,
                text=description,
                font=ctk.CTkFont(size=10),
                text_color=("gray60", "gray70")
            )
            desc_label.pack(pady=(0, 15))
        
        return card