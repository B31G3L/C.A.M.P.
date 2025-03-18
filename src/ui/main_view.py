"""
Hauptansicht für CAMP
Zeigt die Daten und Visualisierungen an
"""
import customtkinter as ctk

class MainView(ctk.CTkFrame):
    """
    Hauptansicht der Anwendung, zeigt die Daten und Visualisierungen an
    """
    def __init__(self, master):
        """
        Initialisiert die Hauptansicht
        
        Args:
            master: Das übergeordnete Widget
        """
        super().__init__(master)
        
        # Erstelle die UI-Elemente
        self._create_widgets()
        
        # Richte das Layout ein
        self._setup_layout()
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für die Hauptansicht"""
        # Platzhalter-Label für jetzt
        self.placeholder_label = ctk.CTkLabel(
            self, 
            text="Hauptansicht wird hier angezeigt",
            font=ctk.CTkFont(size=16)
        )
        
        # Statusleiste am unteren Rand
        self.status_bar = ctk.CTkFrame(self, height=25)
        self.status_label = ctk.CTkLabel(self.status_bar, text="Bereit", anchor="w")
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Platzhalter-Label zentriert
        self.placeholder_label.pack(expand=True)
        
        # Statusleiste am unteren Rand
        self.status_bar.pack(side="bottom", fill="x")
        self.status_label.pack(side="left", padx=10, pady=2, fill="x", expand=True)
    
    def show_message(self, message: str):
        """
        Zeigt eine Nachricht in der Statusleiste an
        
        Args:
            message: Die anzuzeigende Nachricht
        """
        self.status_label.configure(text=message)
    
    def clear(self):
        """Löscht den Inhalt der Hauptansicht"""
        # Hier später den Inhalt der Hauptansicht löschen
        pass
    
    def show_project_data(self, project_data):
        """
        Zeigt die Daten eines Projekts an
        
        Args:
            project_data: Die anzuzeigenden Projektdaten
        """
        # Hier später die Projektdaten anzeigen
        self.placeholder_label.configure(text=f"Projektdaten werden angezeigt: {project_data}")
    
    def show_sprint_data(self, sprint_data):
        """
        Zeigt die Daten eines Sprints an
        
        Args:
            sprint_data: Die anzuzeigenden Sprint-Daten
        """
        # Hier später die Sprint-Daten anzeigen
        self.placeholder_label.configure(text=f"Sprint-Daten werden angezeigt: {sprint_data}")