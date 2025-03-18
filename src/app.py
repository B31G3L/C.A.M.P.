"""
Hauptanwendungsklasse für CAMP
Koordiniert alle Komponenten und enthält die Hauptlogik
"""
import os
import customtkinter as ctk
from src.ui.toolbar import Toolbar
from src.ui.main_view import MainView
from src.ui.data_modal import DataModal

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
            "create_markup": self._create_markup,
            
            # Dropdown-Callbacks für Projekt
            "new_project": self._new_project,
            "open_project": self._open_project,
            "save_project": self._save_project,
            "delete_project": self._delete_project,
            
            # Dropdown-Callbacks für Sprint
            "new_sprint": self._new_sprint,
            "edit_sprint": self._edit_sprint,
            "complete_sprint": self._complete_sprint
        }
    
    # Event-Handler-Methoden
    
    def _on_close(self):
        """Wird aufgerufen, wenn das Fenster geschlossen wird"""
        # Hier könnte geprüft werden, ob ungespeicherte Änderungen vorliegen
        self.root.destroy()
    
    def _show_data(self):
        """Zeigt die Daten aus der CSV-Datei in einem Modal-Fenster an"""
        print("Daten")
        
        # Erstelle und zeige das Daten-Modal
        data_modal = DataModal(self.root, self.csv_path)
        
        # Warte, bis das Modal geschlossen wird
        self.root.wait_window(data_modal)
        
        # Aktualisiere die Statusanzeige in der Hauptansicht
        self.main_view.show_message("Daten wurden angezeigt")
    
    def _create_markup(self):
        """Erstellt Markup-Text und kopiert ihn in die Zwischenablage"""
        print("Markup erstellen/kopieren")
        # Hier später den Markup-Text generieren und in die Zwischenablage kopieren
        self.main_view.show_message("Markup wurde erstellt und kopiert")
    
    def _new_project(self):
        """Erstellt ein neues Projekt"""
        print("Neues Projekt erstellen")
        # Hier später einen Dialog zum Erstellen eines neuen Projekts öffnen
        self.main_view.show_message("Neues Projekt erstellt")
    
    def _open_project(self):
        """Öffnet ein bestehendes Projekt"""
        print("Projekt öffnen")
        # Hier später einen Dialog zum Öffnen eines Projekts anzeigen
        self.main_view.show_message("Projekt geöffnet")
    
    def _save_project(self):
        """Speichert das aktuelle Projekt"""
        print("Projekt speichern")
        # Hier später das aktuelle Projekt speichern
        self.main_view.show_message("Projekt gespeichert")
    
    def _delete_project(self):
        """Löscht das aktuelle Projekt"""
        print("Projekt löschen")
        # Hier später einen Dialog zur Bestätigung anzeigen und dann das Projekt löschen
        self.main_view.show_message("Projekt gelöscht")
    
    def _new_sprint(self):
        """Erstellt einen neuen Sprint"""
        print("Neuen Sprint erstellen")
        # Hier später einen Dialog zum Erstellen eines neuen Sprints öffnen
        self.main_view.show_message("Neuer Sprint erstellt")
    
    def _edit_sprint(self):
        """Bearbeitet den aktuellen Sprint"""
        print("Sprint bearbeiten")
        # Hier später einen Dialog zum Bearbeiten des aktuellen Sprints öffnen
        self.main_view.show_message("Sprint bearbeitet")
    
    def _complete_sprint(self):
        """Schließt den aktuellen Sprint ab"""
        print("Sprint abschließen")
        # Hier später einen Dialog zum Abschließen des Sprints anzeigen
        self.main_view.show_message("Sprint abgeschlossen")
    
    def _refresh_data(self):
        """Aktualisiert die angezeigten Daten"""
        print("Daten aktualisieren")
        # Hier später die Daten neu laden und anzeigen
        self.main_view.show_message("Daten aktualisiert")