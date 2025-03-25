"""
Tests für die Modelklassen
"""
import unittest
from src.models.project import Project
from src.models.sprint import Sprint
from src.models.team_member import TeamMember

class TestProject(unittest.TestCase):
    def test_project_creation(self):
        """Test der Projekterstellung"""
        project = Project("Test Projekt")
        self.assertEqual(project.name, "Test Projekt")
        self.assertEqual(project.teilnehmer, [])
        self.assertEqual(project.sprints, [])
    
    def test_project_from_dict(self):
        """Test der Erstellung aus Dictionary"""
        data = {
            "name": "Test Projekt",
            "teilnehmer": [{"id": "1", "name": "Test Person"}],
            "sprints": [{"sprint_name": "Sprint 1"}]
        }
        project = Project.from_dict(data)
        self.assertEqual(project.name, "Test Projekt")
        self.assertEqual(len(project.teilnehmer), 1)
        self.assertEqual(len(project.sprints), 1)
    
    def test_project_to_dict(self):
        """Test der Konvertierung zu Dictionary"""
        project = Project("Test Projekt", [{"id": "1"}], [{"sprint_name": "Sprint 1"}])
        data = project.to_dict()
        self.assertEqual(data["name"], "Test Projekt")
        self.assertEqual(len(data["teilnehmer"]), 1)
        self.assertEqual(len(data["sprints"]), 1)

class TestSprint(unittest.TestCase):
    def test_sprint_creation(self):
        """Test der Sprinterstellung"""
        sprint = Sprint("Sprint 1", "01.01.2025", "15.01.2025", 50, 45)
        self.assertEqual(sprint.sprint_name, "Sprint 1")
        self.assertEqual(sprint.start, "01.01.2025")
        self.assertEqual(sprint.ende, "15.01.2025")
        self.assertEqual(sprint.confimed_story_points, 50)
        self.assertEqual(sprint.delivered_story_points, 45)

class TestTeamMember(unittest.TestCase):
    def test_team_member_creation(self):
        """Test der Teammitgliederstellung"""
        member = TeamMember("1", "Test Person", "DEV", 0.8)
        self.assertEqual(member.id, "1")
        self.assertEqual(member.name, "Test Person")
        self.assertEqual(member.rolle, "DEV")
        self.assertEqual(member.fte, 0.8)

if __name__ == '__main__':
    unittest.main()