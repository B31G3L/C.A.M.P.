from PyQt5 import QtWidgets
import sys
import json
from content.content import Content
from dialogs.edit_project_modal import EditProjectModal
from dialogs.manage_data_modal import ManageDataModal  # Import new file

class CampGUI(QtWidgets.QMainWindow):
    def __init__(self, config_file):
        super().__init__()
        self.setWindowTitle("CAMP - Capacity Analysis & Management Planner")
        self.setGeometry(100, 100, 1200, 800)

        self.config_file = config_file
        self.projects = {}

        # Toolbar with buttons and dropdown
        self.toolbar = self.addToolBar("Toolbar")

        self.edit_button = QtWidgets.QPushButton("Edit Project")
        self.edit_button.clicked.connect(self.open_edit_modal)
        self.toolbar.addWidget(self.edit_button)

        self.manage_data_button = QtWidgets.QPushButton("Manage Data")
        self.manage_data_button.clicked.connect(self.open_manage_data_modal)
        self.toolbar.addWidget(self.manage_data_button)

        self.toolbar.addSeparator()

        # Spacer, um das Dropdown nach rechts zu schieben
        spacer = QtWidgets.QWidget()
        spacer.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Preferred)
        self.toolbar.addWidget(spacer)

        self.project_dropdown = QtWidgets.QComboBox()
        self.project_dropdown.currentTextChanged.connect(self.update_project)
        self.toolbar.addWidget(self.project_dropdown)

        self.toolbar.setStyleSheet("QToolBar { spacing: 10px; }")

        # Central widget & layout
        self.central_widget = QtWidgets.QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QtWidgets.QVBoxLayout(self.central_widget)

        # SprintsTab
        self.sprints_tab = Content()
        self.layout.addWidget(self.sprints_tab)

        self.load_config()
    def load_config(self):
        try:
            with open(self.config_file, 'r', encoding='utf-8') as file:
                config = json.load(file)
            self.projects = {p['name']: p for p in config.get('projects', [])}
            self.project_dropdown.clear()
            self.project_dropdown.addItems(self.projects.keys())
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Error loading the configuration file: {e}")

    def update_project(self):
        project_name = self.project_dropdown.currentText()
        if project_name in self.projects:
            self.sprints_tab.update_sprints(self.projects[project_name])

    def open_edit_modal(self):
        dialog = EditProjectModal(self.config_file, self)
        dialog.exec_()

    def open_manage_data_modal(self):
        dialog = ManageDataModal(self)  # New modal for "Manage Data"
        dialog.exec_()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = CampGUI("data/project.json")
    window.show()
    sys.exit(app.exec_())
