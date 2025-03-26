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
        
        # 4. Button: "Create Markdown" (umbenannt von "Create/Copy Markup")
        self.markup_btn = ctk.CTkButton(
            self, 
            text="Create Markdown",
            command=self._on_create_markup
        )
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Linke Seite für die meisten Buttons
        self.show_data_btn.pack(side="left", padx=(10, 5), pady=5)
        self.import_data_btn.pack(side="left", padx=5, pady=5)
        self.camp_manager_btn.pack(side="left", padx=5, pady=5)
        
        # "Create Markdown" Button ganz nach rechts
        self.markup_btn.pack(side="right", padx=(5, 10), pady=5)

    # Event-Handler-Methoden
    
    def _on_show_data(self):
        """Wird aufgerufen, wenn der 'Daten anzeigen' Button geklickt wird"""
        if "show_data" in self.callbacks:
            self.callbacks["show_data"]()
    
    def _on_import_data(self):
        """Wird aufgerufen, wenn der 'Import Daten' Button geklickt wird"""
        if "import_data" in self.callbacks:
            self.callbacks["import_data"]()
            
    def _on_camp_manager(self):
        """Wird aufgerufen, wenn der 'CAMP Manager' Button geklickt wird"""
        if "camp_manager" in self.callbacks:
            self.callbacks["camp_manager"]()
    
    def _on_create_markup(self):
        """Wird aufgerufen, wenn der 'Create Markdown' Button geklickt wird"""
        if "markdown_manager" in self.callbacks:
            self.callbacks["markdown_manager"]()