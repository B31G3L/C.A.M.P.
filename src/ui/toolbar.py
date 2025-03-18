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
            text="Daten anzeigen",
            command=self._on_show_data
        )
        
        # 2. Dropdown: "Projekt bearbeiten"
        self.project_options = ["Neues Projekt", "Projekt öffnen", "Projekt speichern", "Projekt löschen"]
        self.project_var = ctk.StringVar(value="Projekt bearbeiten")
        
        self.project_menu = ctk.CTkOptionMenu(
            self,
            values=self.project_options,
            variable=self.project_var,
            command=self._on_project_option
        )
        
        # 3. Dropdown: "Sprints bearbeiten"
        self.sprint_options = ["Neuer Sprint", "Sprint bearbeiten", "Sprint abschließen"]
        self.sprint_var = ctk.StringVar(value="Sprints bearbeiten")
        
        self.sprint_menu = ctk.CTkOptionMenu(
            self,
            values=self.sprint_options,
            variable=self.sprint_var,
            command=self._on_sprint_option
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
        self.project_menu.pack(side="left", padx=5, pady=5)
        self.sprint_menu.pack(side="left", padx=5, pady=5)
        self.markup_btn.pack(side="left", padx=5, pady=5)
    
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