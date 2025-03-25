"""
Sprint-Modell für CAMP
Definiert die Datenstruktur für Sprints
"""
from datetime import datetime

class Sprint:
    """
    Repräsentiert einen Sprint in CAMP
    """
    def __init__(self, sprint_name, start, ende, confimed_story_points=0, delivered_story_points=0):
        """
        Initialisiert einen Sprint
        
        Args:
            sprint_name (str): Name des Sprints
            start (str): Startdatum (Format: "DD.MM.YYYY")
            ende (str): Enddatum (Format: "DD.MM.YYYY")
            confimed_story_points (int): Bestätigte Story Points
            delivered_story_points (int): Gelieferte Story Points
        """
        self.sprint_name = sprint_name
        self.start = start
        self.ende = ende
        self.confimed_story_points = confimed_story_points
        self.delivered_story_points = delivered_story_points
    
    @classmethod
    def from_dict(cls, data):
        """
        Erstellt ein Sprint-Objekt aus einem Dictionary
        
        Args:
            data (dict): Dictionary mit Sprint-Daten
            
        Returns:
            Sprint: Das erstellte Sprint-Objekt
        """
        return cls(
            sprint_name=data.get("sprint_name", ""),
            start=data.get("start", ""),
            ende=data.get("ende", ""),
            confimed_story_points=data.get("confimed_story_points", 0),
            delivered_story_points=data.get("delivered_story_points", 0)
        )
    
    def to_dict(self):
        """
        Konvertiert den Sprint in ein Dictionary
        
        Returns:
            dict: Dictionary-Repräsentation des Sprints
        """
        return {
            "sprint_name": self.sprint_name,
            "start": self.start,
            "ende": self.ende,
            "confimed_story_points": self.confimed_story_points,
            "delivered_story_points": self.delivered_story_points
        }
    
    def is_active(self):
        """
        Prüft, ob der Sprint aktuell aktiv ist
        
        Returns:
            bool: True wenn der Sprint aktiv ist, sonst False
        """
        now = datetime.now()
        start = self._parse_date(self.start)
        ende = self._parse_date(self.ende)
        
        return start <= now <= ende
    
    def _parse_date(self, date_str):
        """
        Wandelt einen Datums-String in ein datetime-Objekt um
        
        Args:
            date_str (str): Datums-String im Format "DD.MM.YYYY"
            
        Returns:
            datetime: Das entsprechende datetime-Objekt
        """
        try:
            day, month, year = map(int, date_str.split('.'))
            return datetime(year, month, day)
        except (ValueError, TypeError):
            # Fallback: Aktuelles Datum zurückgeben
            return datetime.now()