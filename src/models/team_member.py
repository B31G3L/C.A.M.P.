"""
Teammitglied-Modell für CAMP
Definiert die Datenstruktur für Teammitglieder
"""

class TeamMember:
    """
    Repräsentiert ein Teammitglied in CAMP
    """
    def __init__(self, id, name, rolle, fte=1.0):
        """
        Initialisiert ein Teammitglied
        
        Args:
            id (str): ID des Teammitglieds
            name (str): Name des Teammitglieds
            rolle (str): Rolle des Teammitglieds
            fte (float): Full Time Equivalent (0.0-1.0)
        """
        self.id = id
        self.name = name
        self.rolle = rolle
        self.fte = fte
    
    @classmethod
    def from_dict(cls, data):
        """
        Erstellt ein Teammitglied-Objekt aus einem Dictionary
        
        Args:
            data (dict): Dictionary mit Teammitglied-Daten
            
        Returns:
            TeamMember: Das erstellte Teammitglied-Objekt
        """
        return cls(
            id=data.get("id", ""),
            name=data.get("name", ""),
            rolle=data.get("rolle", ""),
            fte=data.get("fte", 1.0)
        )
    
    def to_dict(self):
        """
        Konvertiert das Teammitglied in ein Dictionary
        
        Returns:
            dict: Dictionary-Repräsentation des Teammitglieds
        """
        return {
            "id": self.id,
            "name": self.name,
            "rolle": self.rolle,
            "fte": self.fte
        }