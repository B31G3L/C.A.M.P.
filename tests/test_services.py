"""
Tests für die Serviceklassen
"""
import unittest
import os
import tempfile
import shutil
from src.services.data_service import DataService
from src.models.project import Project

class TestDataService(unittest.TestCase):
    def setUp(self):
        """Testumgebung vorbereiten"""
        # Temporäres Verzeichnis für Testdaten
        self.test_dir = tempfile.mkdtemp()
        self.project_path = os.path.join(self.test_dir, "project.json")
        self.kapa_path = os.path.join(self.test_dir, "kapa_data.csv")
        
        # Initialisiere Service mit Testpfaden
        self.service = DataService(
            project_path=self.project_path,
            kapa_path=self.kapa_path
        )
    
    def tearDown(self):
        """Testumgebung aufräumen"""
        # Temporäres Verzeichnis entfernen
        shutil.rmtree(self.test_dir)
    
    def test_save_and_load_projects(self):
        """Test des Speicherns und Ladens von Projekten"""
        # Projekte erstellen
        project1 = Project("Projekt 1")
        project2 = Project("Projekt 2")
        
        # Projekte speichern
        result = self.service.save_projects([project1, project2])
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.project_path))
        
        # Projekte laden
        loaded_projects = self.service.load_projects()
        self.assertEqual(len(loaded_projects), 2)
        self.assertEqual(loaded_projects[0].name, "Projekt 1")
        self.assertEqual(loaded_projects[1].name, "Projekt 2")
    
    def test_save_and_load_kapa_data(self):
        """Test des Speicherns und Ladens von Kapazitätsdaten"""
        # Kapazitätsdaten erstellen
        kapa_data = [
            ("Mitarbeiter1", "01.01.2025", 8.0, 1.0),
            ("Mitarbeiter2", "01.01.2025", 4.0, 0.5)
        ]
        
        # Daten speichern
        result = self.service.save_kapa_data(kapa_data)
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.kapa_path))
        
        # Daten laden
        loaded_data = self.service.load_kapa_data()
        self.assertEqual(len(loaded_data), 2)
        self.assertEqual(loaded_data[0][0], "Mitarbeiter1")
        self.assertEqual(loaded_data[1][0], "Mitarbeiter2")

if __name__ == '__main__':
    unittest.main()