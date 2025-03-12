from PyQt5 import QtWidgets, QtGui, QtCore
import csv
import os
from datetime import datetime, timedelta

class Content(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        
        self.layout = QtWidgets.QVBoxLayout(self)  # Main layout is vertical

        # The upper part of the layout, taking up 50% of the height
        self.top_layout = QtWidgets.QHBoxLayout()  # Horizontal layout for the upper half
        
        # Sprint List (Left)
        self.left_layout = QtWidgets.QVBoxLayout()
        self.sprint_list = QtWidgets.QListWidget()
        self.sprint_list.itemClicked.connect(self.display_sprint_details)
        self.left_layout.addWidget(self.sprint_list)
        self.top_layout.addLayout(self.left_layout, 1)  # Left takes 25%

        # Sprint Information (Middle)
        self.sprint_info_layout = QtWidgets.QVBoxLayout()

        self.sprint_title = QtWidgets.QLabel("Sprint Name")
        self.sprint_title.setFont(QtGui.QFont('Arial', 16, QtGui.QFont.Bold))

        self.sprint_info_layout.addWidget(self.sprint_title)

        # Start and end date
        self.sprint_dates = QtWidgets.QLabel("Start: 01.01.2025, End: 14.01.2025")

        self.sprint_info_layout.addWidget(self.sprint_dates)

        # Add Divider
        self.divider1 = QtWidgets.QFrame()
        self.divider1.setFrameShape(QtWidgets.QFrame.HLine)
        self.sprint_info_layout.addWidget(self.divider1)

        # Total Capacity
        self.total_capacity_label = QtWidgets.QLabel("Total Capacity: 1000")
        self.sprint_info_layout.addWidget(self.total_capacity_label)

        # Factor 1.4
        self.factor_label = QtWidgets.QLabel("Factor: 1.4")
        self.sprint_info_layout.addWidget(self.factor_label)

        # Divider 2
        self.divider2 = QtWidgets.QFrame()
        self.divider2.setFrameShape(QtWidgets.QFrame.HLine)
        self.sprint_info_layout.addWidget(self.divider2)

        # Result (Total Capacity / 1.4)
        self.result_label = QtWidgets.QLabel("Result: 714.29")
        self.result_label.setStyleSheet("font-weight: bold;")  # Highlighting
        self.sprint_info_layout.addWidget(self.result_label)
        self.sprint_info_layout.setSizeConstraint(QtWidgets.QLayout.SetMinAndMaxSize)

        # The sprint info layout goes into the center (50% of the width)
        self.top_layout.addLayout(self.sprint_info_layout, 2)  # Center takes 50%

        # Capacity Table (Right)
        self.right_layout = QtWidgets.QVBoxLayout()
        self.capacity_table = QtWidgets.QTableWidget()
        self.capacity_table.setColumnCount(2)
        self.capacity_table.setHorizontalHeaderLabels(["Participant", "Capacity"])
        self.capacity_table.verticalHeader().setVisible(False)

        self.right_layout.addWidget(self.capacity_table)
        self.top_layout.addLayout(self.right_layout, 1)  # Right takes 25%

        # The lower part of the layout, taking up 50% of the height
        self.bottom_layout = QtWidgets.QVBoxLayout()  # Vertical layout for the bottom part

        # Daily Table
        self.daily_table = QtWidgets.QTableWidget()
        self.daily_table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.daily_table.horizontalHeader().setStretchLastSection(True)

        self.daily_table.verticalHeader().setVisible(False)

        self.bottom_layout.addWidget(self.daily_table)

        # Add top and bottom layouts to the main layout
        self.layout.addLayout(self.top_layout, 1)  # Top part takes 50% of the tab
        self.layout.addLayout(self.bottom_layout, 1)  # Bottom part takes 50% of the tab
        
        # Add Export Button
        self.export_button = QtWidgets.QPushButton("Export to Confluence")
        self.export_button.clicked.connect(self.export_to_confluence)
        self.layout.addWidget(self.export_button)

    def update_sprints(self, project):
        self.sprint_list.clear()
        self.current_project = project
        current_date = datetime.now()  # Das aktuelle Datum
        
        # Sprints nach Startdatum sortieren, neueste/kommende oben
        sorted_sprints = sorted(
            project.get('sprints', []),
            key=lambda x: datetime.strptime(x['start_date'], "%d.%m.%Y"),
            reverse=True  # Sortierung in absteigender Reihenfolge
        )
        
        # Einen Divider hinzufügen
        divider_added = False
        
        for sprint in sorted_sprints:
            sprint_start_date = datetime.strptime(sprint['start_date'], "%d.%m.%Y")
            sprint_end_date = datetime.strptime(sprint['end_date'], "%d.%m.%Y")

            
            # Wenn der aktuelle Sprint erreicht ist
            if sprint_end_date <= current_date and not divider_added:
                self.sprint_list.addItem('--- Vergangene Sprints ---')  # Divider für vergangene Sprints
                divider_added = True
            
            # Den aktuellen Sprint speziell markieren
            if sprint_start_date <= current_date and current_date <= sprint_end_date:
                self.sprint_list.addItem(f"{sprint['sprint_name']} (active)")  # Aktueller Sprint dick
            else:
                self.sprint_list.addItem(sprint['sprint_name'])


    def display_sprint_details(self, item):
        sprint_name = item.text()

        # Entferne den Zusatz "(active)", falls vorhanden, für den Vergleich
        pure_sprint_name = sprint_name.replace(" (active)", "")

        if not self.current_project:
            print("No current project found.")
            return

        # Suche nach dem Sprint ohne den "(active)" Zusatz
        sprint = next((s for s in self.current_project.get('sprints', []) if s['sprint_name'] == pure_sprint_name), None)
        if sprint:
            print(f"Displaying details for sprint: {sprint_name}")

            # Lade die Kapazitätsdaten (Dummy-Implementierung)
            capacity_data, daily_availability = self.load_capacity_data(self.current_project, sprint_name, sprint['start_date'], sprint['end_date'])

            # Debug-Ausgabe für Kapazitätsdaten und tägliche Verfügbarkeit
            print("Loaded capacity data:", capacity_data)
            print("Loaded daily availability:", daily_availability)

            self.display_capacity(capacity_data)
            self.display_daily_availability(daily_availability, sprint['start_date'], sprint['end_date'])

            # Update der Sprint-Informationen im neuen Layout
            self.sprint_title.setText(sprint_name)
            self.sprint_dates.setText(f"Start: {sprint['start_date']}, End: {sprint['end_date']}")

            # Berechne die Gesamt-Kapazität (stelle sicher, dass es ein int ist)
            total_capacity = 0
            for name, capacity in capacity_data.items():
                print(f"Processing participant: {name}, Capacity: {capacity}")
                try:
                    # Nur addieren, wenn der Wert eine gültige Zahl ist
                    total_capacity += float(capacity) if str(capacity).replace('.', '', 1).isdigit() else 0
                except ValueError:
                    print(f"Invalid capacity value for {name}: {capacity}, skipping.")

            self.total_capacity_label.setText(f"Total Capacity: {total_capacity}")

            # Berechne und zeige das Ergebnis (Gesamt-Kapazität / 1.4)
            result = total_capacity / 1.4 if total_capacity else 0
            self.result_label.setText(f"Result: {result:.2f}")
            self.capacity_table.resizeColumnsToContents()
            self.daily_table.resizeColumnsToContents()
            self.daily_table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            self.daily_table.horizontalHeader().setStretchLastSection(True)

            self.daily_table.verticalHeader().setVisible(False)

        else:
            print(f"Sprint with name {sprint_name} not found in the project.")



    def load_capacity_data(self, project, sprint_name, start_date, end_date):
        print(f"Loading capacity data for project: {project.get('name', 'Unknown')} - Sprint: {sprint_name}")
        
        capacity_data = {m['name']: 0 for m in project.get('participants', [])}
        daily_availability = {m['name']: {} for m in project.get('participants', [])}

        csv_path = "data/kapa_data.csv" # Angepasst: kapa_data.csv
        if not os.path.exists(csv_path):
            print(f"Error: CSV file not found at {csv_path}")
            return capacity_data, daily_availability

        sprint_start = datetime.strptime(start_date, "%d.%m.%Y")
        sprint_end = datetime.strptime(end_date, "%d.%m.%Y")
        sprint_days = [(sprint_start + timedelta(days=i)).strftime("%d.%m.%Y") for i in range((sprint_end - sprint_start).days + 1)]

        for name in daily_availability:
            daily_availability[name] = {day: {'available': False, 'hours': 0} for day in sprint_days}

        with open(csv_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            next(reader, None)  # Skip header row, if present

            for row in reader:
                try:
                    if len(row) < 3: # Angepasst: 3 Spalten in kapa_data.csv
                        print(f"Skipping invalid row (not enough columns): {row}")
                        continue
                                
                    name = next((m['name'] for m in project.get('participants', []) if str(m['id']) == row[0]), None)

                    date_str = row[1] # Angepasst: Datum ist die zweite Spalte
                    hours = float(row[2].replace(',', '.')) if row[2] else 0 # Angepasst: Stunden ist die dritte Spalte

                    try:
                        entry_date = datetime.strptime(date_str, "%d.%m.%Y")
                    except (ValueError, IndexError):
                        print(f"Invalid date format in row: {date_str}")
                        continue
                    
                    print(name)
                    print(date_str)
                    print(hours)
                    # Only consider data within the sprint period
                    if sprint_start <= entry_date <= sprint_end:
                        capacity_data[name] += hours / 8
                        daily_availability[name][entry_date.strftime("%d.%m.%Y")]['available'] = True
                        daily_availability[name][entry_date.strftime("%d.%m.%Y")]['hours'] += hours

                except Exception as e:
                    print(f"Unexpected error while processing row {row}: {e}")

        print(f"Capacity data loaded: {capacity_data}")
        print(f"Daily availability sample: {list(daily_availability.items())[:3]}")  # Shows only a few values

        return capacity_data, daily_availability

    def display_capacity(self, capacity_data):
        self.capacity_table.setRowCount(len(capacity_data))
        for row, (name, capacity) in enumerate(capacity_data.items()):
            # Formatierung: Wenn Kapazität eine ganze Zahl ist, keine Dezimalstellen, sonst mit einer Dezimalstelle
            formatted_capacity = "{:.0f}".format(capacity) if capacity.is_integer() else "{:.2f}".format(capacity)

            # Zelle für den Namen erstellen und Schriftfarbe auf Weiß setzen
            cell_name = QtWidgets.QTableWidgetItem(name)
            cell_name.setFlags(cell_name.flags() & ~QtCore.Qt.ItemIsSelectable & ~QtCore.Qt.ItemIsEnabled)
            cell_name.setForeground(QtGui.QColor(255, 255, 255))  # Weiß
            self.capacity_table.setItem(row, 0, cell_name)

            # Zelle für die formatierte Kapazität erstellen und Schriftfarbe auf Weiß setzen
            cell_formatted_capacity = QtWidgets.QTableWidgetItem(formatted_capacity)
            cell_formatted_capacity.setFlags(cell_formatted_capacity.flags() & ~QtCore.Qt.ItemIsSelectable & ~QtCore.Qt.ItemIsEnabled)
            cell_formatted_capacity.setForeground(QtGui.QColor(255, 255, 255))  # Weiß
            self.capacity_table.setItem(row, 1, cell_formatted_capacity)

        



    def get_color_for_hours(self, hours):
        """Function to determine the background color based on the number of hours."""
        if hours >= 8:
            return QtGui.QColor("green")
        elif hours == 0:
            return QtGui.QColor("red")
        else:
            return QtGui.QColor("yellow")


    def display_daily_availability(self, daily_availability, start_date, end_date):
        try:
            # Convert the start and end dates
            sprint_start = datetime.strptime(start_date, "%d.%m.%Y")
            sprint_end = datetime.strptime(end_date, "%d.%m.%Y")

            # List all sprint days
            sprint_days = [(sprint_start + timedelta(days=i)).strftime("%d.%m.%Y") 
                        for i in range((sprint_end - sprint_start).days + 1)]

            # Update table structure
            self.daily_table.setColumnCount(len(sprint_days) + 1)
            self.daily_table.setHorizontalHeaderLabels(["Name"] + sprint_days)
            self.daily_table.setRowCount(len(daily_availability))

            for row, (name, days) in enumerate(daily_availability.items()):
                self.daily_table.setItem(row, 0, QtWidgets.QTableWidgetItem(name))
                
                for col, day in enumerate(sprint_days, start=1):
                    cell = QtWidgets.QTableWidgetItem()
                    
                    try:
                        # Stunden für den aktuellen Tag abrufen
                        hours = days.get(day, {}).get('hours', 0)  # Verhindert KeyError
                        
                        # Prüfen, ob das Datum ein Wochenende ist
                        day_date = datetime.strptime(day, "%d.%m.%Y")
                        if day_date.weekday() >= 5:  # 5 = Samstag, 6 = Sonntag
                            cell.setBackground(QtGui.QColor("gray"))  # Wochenenden grau
                        else:
                            if hasattr(self, 'get_color_for_hours'):
                                cell.setBackground(self.get_color_for_hours(hours))
                            else:
                                print(f"Warnung: Methode get_color_for_hours fehlt!")
                        
                        # Tooltip setzen
                        cell.setToolTip(f"{hours} Stunden")

                        # Zellen anklickbar machen
                        cell.setFlags(cell.flags() & ~QtCore.Qt.ItemIsSelectable & ~QtCore.Qt.ItemIsEnabled)

                        # Zelle in die Tabelle einfügen
                        self.daily_table.setItem(row, col, cell)

                    except Exception as e:
                        print(f"Fehler in Zelle ({row}, {col}): {e}")

        except Exception as e:
            print(f"Error displaying daily availability: {e}")




    def export_to_confluence(self):
        """Exports sprint data as HTML or Wiki Markup for Confluence"""
        sprint_name = self.sprint_title.text()
        if not self.current_project:
            print("No current project found.")
            return

        sprint = next((s for s in self.current_project.get('sprints', []) if s['sprint_name'] == sprint_name), None)
        if sprint:
            # Load capacity data and daily availability
            capacity_data, daily_availability = self.load_capacity_data(self.current_project, sprint_name, sprint['start_date'], sprint['end_date'])

            # Create HTML or Wiki Markup
            confluence_content = self.create_confluence_content(sprint, capacity_data, daily_availability)

            # Save or output the content
            file_path = os.path.join(os.getcwd(), f"{sprint_name}_export.html")
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(confluence_content)
            print(f"Export successful. File saved as {file_path}")

    def create_confluence_content(self, sprint, capacity_data, daily_availability):
        """
        Creates the Confluence Wiki Markup content that matches the GUI layout and colors.
        Adds a checkbox field (as [ ]) next to each participant's capacity in the capacity table.
        """
        # Calculate total capacity and result (sum of capacity / 1.4)
        total_capacity = sum(float(capacity) for capacity in capacity_data.values())
        result = total_capacity / 1.4 if total_capacity else 0

        # Create list of all sprint days
        sprint_start_date = datetime.strptime(sprint['start_date'], "%d.%m.%Y")
        sprint_end_date = datetime.strptime(sprint['end_date'], "%d.%m.%Y")
        sprint_days = [
            (sprint_start_date + timedelta(days=i)).strftime("%d.%m.%Y")
            for i in range((sprint_end_date - sprint_start_date).days + 1)
        ]

        def get_bg_color(hours, day_str):
            """Determines the background color similar to the GUI logic:
            - Gray on weekends,
            - Green if hours >= 8,
            - Red if hours == 0,
            - Otherwise yellow.
            """
            day_date = datetime.strptime(day_str, "%d.%m.%Y")
            if day_date.weekday() >= 5:  # Saturday or Sunday
                return "gray"
            else:
                if hours >= 8:
                    return "green"
                elif hours == 0:
                    return "red"
                else:
                    return "yellow"

        # Create the Wiki Markup content
        content = f"h1. {sprint['sprint_name']}\n\n"
        content += f"*Start:* {sprint['start_date']} &nbsp;&nbsp; *End:* {sprint['end_date']}\n\n"

        content += "h2. Sprint Information\n"
        content += f"* Total Capacity: {total_capacity:.2f}\n"
        content += "* Factor: 1.4\n"
        content += f"* Result: {result:.2f}\n\n"

        content += "h2. Capacity Table\n"
        # Add an additional column for the checkbox (e.g., as "Done")
        content += "|| Participant || Capacity || Done ||\n"
        for name, capacity in capacity_data.items():
            # The checkbox is output here as [ ].
            content += f"| {name} | {capacity:.2f} | [ ] |\n"
        content += "\n"

        content += "h2. Daily Availability\n"
        # Table header with a header for each sprint day
        content += "|| Name || " + " || ".join(sprint_days) + " ||\n"
        for name, days in daily_availability.items():
            content += f"| {name} "
            for day in sprint_days:
                # Get the hours (default: 0 if no entry)
                hours = days.get(day, {}).get('hours', 0)
                bg_color = get_bg_color(hours, day)
                # Use the panel macro to color the background
                cell_content = f"{{panel:bgColor={bg_color}|borderStyle=none|borderWidth=0}}{hours}{{panel}}"
                content += f" | {cell_content} "
            content += "|\n"

        return content
