"""
Projektmodell für CAMP
Definiert die Datenstruktur für Projekte
"""

class Project:
    """
    Repräsentiert ein Projekt in CAMP
    """
    def __init__(self, name, teilnehmer=None, sprints=None):
        """
        Initialisiert ein Projekt
        
        Args:
            name (str): Name des Projekts
            teilnehmer (list): Liste der Teammitglieder
            sprints (list): Liste der Sprints
        """
        self.name = name
        self.teilnehmer = teilnehmer or []
        self.sprints = sprints or []
    
    @classmethod
    def from_dict(cls, data):
        """
        Erstellt ein Projektobjekt aus einem Dictionary
        
        Args:
            data (dict): Dictionary mit Projektdaten
            
        Returns:
            Project: Das erstellte Projektobjekt
        """
        return cls(
            name=data.get("name", ""),
            teilnehmer=data.get("teilnehmer", []),
            sprints=data.get("sprints", [])
        )
    
    def to_dict(self):
        """
        Konvertiert das Projekt in ein Dictionary
        
        Returns:
            dict: Dictionary-Repräsentation des Projekts
        """
        return {
            "name": self.name,
            "teilnehmer": self.teilnehmer,
            "sprints": self.sprints
        }