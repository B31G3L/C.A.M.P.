"""
Toolbar-Komponente für CAMP
Enthält Buttons und Dropdowns für die Hauptfunktionen
"""
import customtkinter as ctk
from typing import Dict, Callable, Any, Optional

class Toolbar(ctk.CTkFrame):
    """
    Toolbar mit Buttons und Dropdowns für die wichtigsten Funktionen
    """
    def __init__(self, master, callbacks: Optional[Dict[str, Callable]] = None):
        """
        Initialisiert die Toolbar
        
        Args:
            master: Das übergeordnete Widget
            callbacks: Dictionary mit Callback-Funktionen für die Toolbar-Elemente
        """
        super().__init__(master, height=50)
        
        # Speichere Callbacks
        self.callbacks = callbacks or {}
        
        # Erstelle die UI-Elemente
        self._create_widgets()
        
        # Richte das Layout ein
        self._setup_layout()
    

    def _create_widgets(self):
        """Erstellt alle UI-Elemente für die Toolbar"""
        # 1. Button: "Daten anzeigen"
        self.show_data_btn = ctk.CTkButton(
            self, 
            text="Daten",
            command=self._on_show_data
        )
        
        # 2. Button: "Import Daten"
        self.import_data_btn = ctk.CTkButton(
            self, 
            text="Import Daten",
            command=self._on_import_data
        )
        
        # 3. Button: "CAMP Manager" 
        self.camp_manager_btn = ctk.CTkButton(
            self, 
            text="CAMP Manager",
            command=self._on_camp_manager
        )
        
        # 4. Button: "Create/Copy Markup"
        self.markup_btn = ctk.CTkButton(
            self, 
            text="Create/Copy Markup",
            command=self._on_create_markup
        )
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Alle Elemente nebeneinander anordnen
        self.show_data_btn.pack(side="left", padx=(10, 5), pady=5)
        self.import_data_btn.pack(side="left", padx=5, pady=5)
        self.camp_manager_btn.pack(side="left", padx=5, pady=5)
        self.markup_btn.pack(side="left", padx=5, pady=5)

    def _on_camp_manager(self):
        """Wird aufgerufen, wenn der 'CAMP Manager' Button geklickt wird"""
        if "camp_manager" in self.callbacks:
            self.callbacks["camp_manager"]()

    def _on_import_data(self):
        """Wird aufgerufen, wenn der 'Import Daten' Button geklickt wird"""
        if "import_data" in self.callbacks:
            self.callbacks["import_data"]()

    # Event-Handler-Methoden
    
    def _on_show_data(self):
        """Wird aufgerufen, wenn der 'Daten anzeigen' Button geklickt wird"""
        if "show_data" in self.callbacks:
            self.callbacks["show_data"]()
    
    def _on_project_option(self, option: str):
        """
        Wird aufgerufen, wenn eine Option im 'Projekt bearbeiten' Dropdown ausgewählt wird
        
        Args:
            option: Die ausgewählte Option
        """
        # Dropdown zurücksetzen
        self.project_var.set("Projekt bearbeiten")
        
        # Bestimme den passenden Callback-Schlüssel
        callback_key = ""
        if option == "Neues Projekt":
            callback_key = "new_project"
        elif option == "Projekt öffnen":
            callback_key = "open_project"
        elif option == "Projekt speichern":
            callback_key = "save_project"
        elif option == "Projekt löschen":
            callback_key = "delete_project"
        
        # Rufe den passenden Callback auf, falls vorhanden
        if callback_key and callback_key in self.callbacks:
            self.callbacks[callback_key]()
    
    def _on_sprint_option(self, option: str):
        """
        Wird aufgerufen, wenn eine Option im 'Sprints bearbeiten' Dropdown ausgewählt wird
        
        Args:
            option: Die ausgewählte Option
        """
        # Dropdown zurücksetzen
        self.sprint_var.set("Sprints bearbeiten")
        
        # Bestimme den passenden Callback-Schlüssel
        callback_key = ""
        if option == "Neuer Sprint":
            callback_key = "new_sprint"
        elif option == "Sprint bearbeiten":
            callback_key = "edit_sprint"
        elif option == "Sprint abschließen":
            callback_key = "complete_sprint"
        
        # Rufe den passenden Callback auf, falls vorhanden
        if callback_key and callback_key in self.callbacks:
            self.callbacks[callback_key]()
    
    def _on_create_markup(self):
        """Wird aufgerufen, wenn der 'Create/Copy Markup' Button geklickt wird"""
        if "create_markup" in self.callbacks:
            self.callbacks["create_markup"]()

    def _get_toolbar_callbacks(self):
        """
        Erstellt ein Dictionary mit den Callback-Funktionen für die Toolbar
        
        Returns:
            dict: Dictionary mit Callback-Funktionen
        """
        return {
            # Button-Callbacks
            "show_data": self._show_data,
            "import_data": self._import_data,
            "create_markup": self._show_import_data,
            "camp_manager": self._show_camp_manager,
            
            # Die alten Callbacks können bleiben, sie werden vom Manager-Modal verwendet
            "new_project": self._new_project,
            "open_project": self._open_project,
            "save_project": self._save_project,
            "delete_project": self._delete_project,
            "new_sprint": self._new_sprint,
            "edit_sprint": self._edit_sprint,
            "complete_sprint": self._complete_sprint
        }
