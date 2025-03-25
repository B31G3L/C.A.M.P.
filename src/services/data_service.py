"""
Datenservice für CAMP
Verwaltet den Zugriff auf Projektdaten und Kapazitätsdaten
"""
import os
import json
import csv
from src.models.project import Project
from src.models.sprint import Sprint
from src.models.team_member import TeamMember

class DataService:
    """
    Service für den Zugriff auf die Projektdaten und Kapazitätsdaten
    """
    def __init__(self, project_path="data/project.json", kapa_path="data/kapa_data.csv"):
        """
        Initialisiert den Datenservice
        
        Args:
            project_path (str): Pfad zur project.json-Datei
            kapa_path (str): Pfad zur kapa_data.csv-Datei
        """
        self.project_path = project_path
        self.kapa_path = kapa_path
    
    def load_projects(self):
        """
        Lädt alle Projekte aus der project.json-Datei
        
        Returns:
            list: Liste der Projekte
        """
        try:
            if os.path.exists(self.project_path):
                with open(self.project_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    
                projects = []
                for project_data in data.get("projekte", []):
                    projects.append(Project.from_dict(project_data))
                    
                return projects
            return []
        except Exception as e:
            print(f"Fehler beim Laden der Projektdaten: {e}")
            return []
    
    def save_projects(self, projects):
        """
        Speichert alle Projekte in die project.json-Datei
        
        Args:
            projects (list): Liste der Projekte
            
        Returns:
            bool: True wenn erfolgreich, False wenn ein Fehler aufgetreten ist
        """
        try:
            # Stelle sicher, dass das Verzeichnis existiert
            os.makedirs(os.path.dirname(self.project_path), exist_ok=True)
            
            # Konvertiere Projekte in Dict-Format
            data = {
                "projekte": [project.to_dict() for project in projects]
            }
            
            # Speichere die Daten
            with open(self.project_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            
            return True
        except Exception as e:
            print(f"Fehler beim Speichern der Projektdaten: {e}")
            return False
    
    def load_kapa_data(self):
        """
        Lädt die Kapazitätsdaten aus der kapa_data.csv-Datei
        
        Returns:
            list: Liste von Tupeln (id, datum, stunden, kapazität)
        """
        data = []
        try:
            if os.path.exists(self.kapa_path):
                with open(self.kapa_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=';')
                    for row in reader:
                        if len(row) >= 3:
                            # Stunden als Float
                            stunden = float(row[2].replace(',', '.'))
                            
                            # Kapazität aus der Datei lesen oder berechnen
                            if len(row) >= 4:
                                kapazitaet = float(row[3].replace(',', '.'))
                            else:
                                # Kapazität berechnen: Stunden >= 8 entsprechen Kapazität 1
                                kapazitaet = 1.0 if stunden >= 8 else round(stunden / 8, 2)
                            
                            data.append((row[0], row[1], stunden, kapazitaet))
        except Exception as e:
            print(f"Fehler beim Laden der Kapazitätsdaten: {e}")
        
        return data
    
    def save_kapa_data(self, kapa_data, create_backup=True):
        """
        Speichert die Kapazitätsdaten in die kapa_data.csv-Datei
        
        Args:
            kapa_data (list): Liste von Tupeln (id, datum, stunden, kapazität)
            create_backup (bool): Ob ein Backup erstellt werden soll
            
        Returns:
            bool: True wenn erfolgreich, False wenn ein Fehler aufgetreten ist
        """
        try:
            # Stelle sicher, dass das Verzeichnis existiert
            os.makedirs(os.path.dirname(self.kapa_path), exist_ok=True)
            
            # Erstelle ein Backup, falls gewünscht
            if create_backup and os.path.exists(self.kapa_path):
                import shutil
                backup_path = f"{self.kapa_path}.bak"
                shutil.copy2(self.kapa_path, backup_path)
            
            # Speichere die Daten
            with open(self.kapa_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                for row in kapa_data:
                    writer.writerow(row)
            
            return True
        except Exception as e:
            print(f"Fehler beim Speichern der Kapazitätsdaten: {e}")
            return False