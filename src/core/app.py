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
from src.services.markdown_service import MarkdownService
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
        self.root.geometry("1280x960")
        
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

    def _create_markup(self):
        """Generiert Confluence Wiki-Text basierend auf den aktuellen Daten und zeigt ihn in einem Modal an"""
        # Import der WikiExporter-Klasse
        from src.utils.wiki_exporter import WikiExporter
        
        # Prüfen, ob ein Sprint in der Hauptansicht ausgewählt ist
        if not hasattr(self.main_view, 'selected_sprint') or not self.main_view.selected_sprint:
            self.main_view.show_message("Bitte wählen Sie zuerst einen Sprint aus")
            return
            
        # Generiere Confluence Wiki-Text
        confluence_text = WikiExporter.generate_confluence_wiki(self.main_view)
        
        # Erstelle und zeige das Markdown-Modal
        markdown_modal = MarkdownModal(
            self.root, 
            confluence_text, 
            title="Sprint Confluence Wiki"
        )
        
        # Warte, bis das Modal geschlossen wird
        self.root.wait_window(markdown_modal)
        
        # Aktualisiere die Statusanzeige in der Hauptansicht
        self.main_view.show_message("Confluence Wiki wurde erstellt")
            
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