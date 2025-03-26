"""
Hauptanwendungsklasse für CAMP
Koordiniert alle Komponenten und enthält die Hauptlogik
"""
import os
import customtkinter as ctk
from src.ui.components.toolbar import Toolbar
from src.ui.views.main_view import MainView
from src.ui.dialogs.data_modal import DataModal
from src.ui.dialogs.camp_manager_modal import CAMPManagerModal
from src.ui.dialogs.import_modal import ImportModal
from src.ui.dialogs.markdown_modal import MarkdownModal
from datetime import datetime


class CAMPApp:
    """
    Hauptanwendungsklasse für die CAMP-Anwendung
    """
    def __init__(self, root):
        """
        Initialisiert die Anwendung
        
        Args:
            root: Das Hauptfenster (CTk-Root)
        """
        self.root = root
        self.root.title("CAMP - Capacity Analysis & Management Planner")
        self.root.geometry("1200x800")
        
        # Konfiguration
        self.csv_path = os.path.join("data", "kapa_data.csv")
        
        # Stelle sicher, dass das Datenverzeichnis existiert
        self._ensure_data_directory()
        
        # Erstelle die Komponenten
        self._create_ui_components()
        
        # Richte das Layout ein
        self._setup_layout()
        
        # Registriere Event-Handler
        self._register_event_handlers()
    
    def _ensure_data_directory(self):
        """Stellt sicher, dass das Datenverzeichnis existiert"""
        os.makedirs("data", exist_ok=True)
    
    def _create_ui_components(self):
        """Erstellt alle UI-Komponenten"""
        # Erstelle die Toolbar mit Callbacks
        self.toolbar = Toolbar(self.root, self._get_toolbar_callbacks())
        
        # Erstelle die Hauptansicht
        self.main_view = MainView(self.root)
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Komponenten ein"""
        # Toolbar am oberen Rand
        self.toolbar.pack(fill="x", padx=10, pady=10)
        
        # Hauptansicht darunter, nimmt den restlichen Platz ein
        self.main_view.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def _register_event_handlers(self):
        """Registriert Event-Handler für die Anwendung"""
        # Fenster-Schließen-Event
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # Tastatur-Shortcuts
        self.root.bind("<Control-s>", lambda e: self._save_project())
        self.root.bind("<Control-o>", lambda e: self._open_project())
        self.root.bind("<F5>", lambda e: self._refresh_data())
    
    def _get_toolbar_callbacks(self):
        """
        Erstellt ein Dictionary mit den Callback-Funktionen für die Toolbar
        
        Returns:
            dict: Dictionary mit Callback-Funktionen
        """
        return {
            # Button-Callbacks
            "show_data": self._show_data,
            "import_data": self._show_import_data,
            "create_markup": self._create_markup,
            "camp_manager": self._show_camp_manager
        }
    
    # Event-Handler-Methoden
    
    def _on_close(self):
        """Wird aufgerufen, wenn das Fenster geschlossen wird"""
        self.root.destroy()
    
    def _show_data(self):
        """Zeigt die Daten aus der CSV-Datei in einem Modal-Fenster an"""
        # Erstelle und zeige das Daten-Modal
        data_modal = DataModal(self.root, self.csv_path)
        
        # Warte, bis das Modal geschlossen wird
        self.root.wait_window(data_modal)
        
        # Aktualisiere die Statusanzeige in der Hauptansicht
        self.main_view.show_message("Daten wurden angezeigt")
    
    def _create_markup(self):
        """Generiert Markdown-Text basierend auf den aktuellen Daten und zeigt ihn in einem Modal an"""
        # Prüfen, ob ein Sprint in der Hauptansicht ausgewählt ist
        if not hasattr(self.main_view, 'selected_sprint') or not self.main_view.selected_sprint:
            self.main_view.show_message("Bitte wählen Sie zuerst einen Sprint aus")
            return
            
        # Generiere Markdown-Text
        markdown_text = self._generate_markdown()
        
        # Erstelle und zeige das Markdown-Modal
        markdown_modal = MarkdownModal(
            self.root, 
            markdown_text, 
            title="Sprint Markdown"
        )
        
        # Warte, bis das Modal geschlossen wird
        self.root.wait_window(markdown_modal)
        
        # Aktualisiere die Statusanzeige in der Hauptansicht
        self.main_view.show_message("Markdown wurde erstellt")
    
    def _generate_markdown(self):
        """
        Generiert Markdown-Text mit Sprint-Details, Team-Auslastung und Sprint-Kalender
        
        Returns:
            str: Der generierte Markdown-Text
        """
        sprint = self.main_view.selected_sprint
        project = self.main_view.selected_project
        
        if not sprint or not project:
            return "Kein Sprint oder Projekt ausgewählt."
        
        # Header mit Projekt und Sprint
        markdown = f"# {project.get('name', 'Projekt')} - {sprint.get('sprint_name', 'Sprint')}\n\n"
        
        # Aktuelles Datum
        markdown += f"*Erstellt am: {datetime.now().strftime('%d.%m.%Y')}*\n\n"
        
        # Sprint Details
        markdown += "## Sprint Details\n\n"
        markdown += f"* **Zeitraum:** {sprint.get('start', '')} - {sprint.get('ende', '')}\n"
        
        # Geplante und gelieferte Story Points
        if 'confimed_story_points' in sprint:
            markdown += f"* **Geplante Story Points:** {sprint.get('confimed_story_points', 0)}\n"
        if 'delivered_story_points' in sprint:
            markdown += f"* **Gelieferte Story Points:** {sprint.get('delivered_story_points', 0)}\n"
        
        # Sprint-Status basierend auf Datum
        start_date = self.main_view._parse_date(sprint.get('start', ''))
        end_date = self.main_view._parse_date(sprint.get('ende', ''))
        now = datetime.now()
        
        if now < start_date:
            status = "Geplant"
            days_to_start = (start_date - now).days
            markdown += f"* **Status:** {status} (startet in {days_to_start} Tagen)\n"
        elif start_date <= now <= end_date:
            status = "Aktiv"
            days_left = (end_date - now).days
            markdown += f"* **Status:** {status} (noch {days_left} Tage)\n"
        else:
            status = "Abgeschlossen"
            days_since = (now - end_date).days
            markdown += f"* **Status:** {status} (vor {days_since} Tagen)\n"
        
        # Sprint-Dauer in Tagen
        duration = (end_date - start_date).days + 1
        markdown += f"* **Dauer:** {duration} Tage\n\n"
        
        # Team Auslastung
        markdown += "## Team Auslastung\n\n"
        
        # Header der Tabelle
        markdown += "| Name | Rolle | Stunden | Kapazität |\n"
        markdown += "|------|-------|---------|----------|\n"
        
        # Kapazitätsdaten für den Sprint laden
        team_members = project.get('teilnehmer', [])
        
        # Sprint-Kapazitätsdaten abrufen (wenn vorhanden in der MainView)
        sprint_capacity_data = self.main_view._get_sprint_capacity_data() if hasattr(self.main_view, '_get_sprint_capacity_data') else []
        
        # Dictionary für schnelleren Zugriff
        capacity_dict = {}
        for member_id, hours, capacity in sprint_capacity_data:
            capacity_dict[member_id] = (hours, capacity)
        
        # Teammitglieder zur Tabelle hinzufügen
        for member in team_members:
            member_id = member.get('id', '')
            member_name = member.get('name', 'Unbekannt')
            member_role = member.get('rolle', '')
            
            # Kapazität und Stunden für dieses Mitglied
            hours, capacity = capacity_dict.get(member_id, (0, 0))
            
            # Zeile zur Tabelle hinzufügen
            markdown += f"| {member_name} | {member_role} | {hours:.1f} | {capacity:.2f} |\n"
        
        markdown += "\n"
        
        # Sprint-Kalender
        markdown += "## Sprint-Kalender\n\n"
        markdown += "Für detaillierten Sprint-Kalender bitte den CAMP Manager verwenden.\n\n"
        
        # Zusammenfassung
        markdown += "## Zusammenfassung\n\n"
        
        # Berechne die Gesamtkapazität und geschätzte Story Points
        total_capacity = sum(capacity for _, _, capacity in sprint_capacity_data)
        umrechnungsfaktor = 1.4  # Fester Umrechnungsfaktor
        berechnete_sp = total_capacity / umrechnungsfaktor if total_capacity > 0 else 0
        
        markdown += f"* **Gesamtkapazität:** {total_capacity:.2f}\n"
        markdown += f"* **Umrechnungsfaktor:** {umrechnungsfaktor}\n"
        markdown += f"* **Berechnete Story Points:** {berechnete_sp:.1f}\n"
        
        if 'confimed_story_points' in sprint and sprint.get('confimed_story_points', 0) > 0:
            # Differenz zwischen geplanten und berechneten SPs
            diff = sprint.get('confimed_story_points', 0) - berechnete_sp
            markdown += f"* **Differenz zu geplanten Story Points:** {diff:.1f}\n"
        
        return markdown
    
    

    def _show_camp_manager(self):
        """Zeigt das CAMP Manager Modal an"""
        # Erstelle und zeige das CAMP Manager-Modal
        manager_modal = CAMPManagerModal(self.root, self._get_toolbar_callbacks())
        
        # Warte, bis das Modal geschlossen wird
        self.root.wait_window(manager_modal)
        
        # Aktualisiere die Statusanzeige in der Hauptansicht
        self.main_view.show_message("CAMP Manager wurde geschlossen")

    def _show_import_data(self):
        """Zeigt das Import-Daten-Modal an"""
        # Erstelle und zeige das Import-Modal
        import_modal = ImportModal(self.root, self._get_toolbar_callbacks())
        
        # Warte, bis das Modal geschlossen wird
        self.root.wait_window(import_modal)
        
        # Aktualisiere die Statusanzeige in der Hauptansicht
        self.main_view.show_message("Daten-Import wurde geschlossen")