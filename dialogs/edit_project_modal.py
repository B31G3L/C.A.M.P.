from PyQt5 import QtWidgets
import json

class EditProjectModal(QtWidgets.QDialog):
    def __init__(self, config_file, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Project JSON")
        self.setGeometry(200, 200, 600, 400)
        self.config_file = config_file
        self.parent = parent

        layout = QtWidgets.QVBoxLayout()
        self.text_edit = QtWidgets.QPlainTextEdit()

        # Set monospace font
        font = self.text_edit.font()
        font.setFamily("Monospace")  # Alternatively "Monospace" or "Consolas"
        font.setPointSize(10)
        self.text_edit.setFont(font)

        layout.addWidget(self.text_edit)
        self.load_config()

        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.validate_and_save)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def load_config(self):
        try:
            with open(self.config_file, 'r', encoding='utf-8') as file:
                self.text_edit.setPlainText(file.read())
        except FileNotFoundError:
            QtWidgets.QMessageBox.warning(self, "File Not Found", "The configuration file does not exist. A new file will be created.")
            self.text_edit.setPlainText("{\n    \"projects\": []\n}")  # Default value for an empty file
        except json.JSONDecodeError as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Error loading configuration file: Invalid JSON\n{e}")

    def validate_and_save(self):
        try:
            # Try to parse the JSON
            json_data = json.loads(self.text_edit.toPlainText())
            # Check if the root element is a dictionary with the key "projects"
            if not isinstance(json_data, dict) or "projects" not in json_data:
                raise ValueError("Invalid JSON format. The root element must be an object with a 'projects' key.")
            
            # Save the modified data back to the file
            with open(self.config_file, 'w', encoding='utf-8') as file:
                json.dump(json_data, file, indent=4, ensure_ascii=False)
            
            # Show a confirmation message
            QtWidgets.QMessageBox.information(self, "Saved", "The changes were successfully saved.")  
            self.accept()

            # Reload the configuration in the parent, if available
            if self.parent:
                self.parent.load_config()
        except (json.JSONDecodeError, ValueError) as e:
            # Error parsing or validating the JSON
            QtWidgets.QMessageBox.critical(self, "Error", f"Invalid JSON: {e}")
