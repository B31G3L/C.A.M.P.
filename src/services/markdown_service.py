"""
Markdown-Service für CAMP
Generiert Markdown-Inhalte basierend auf Projekt- und Sprint-Daten
"""
from datetime import datetime, timedelta

class MarkdownService:
    """
    Service zur Generierung von Markdown-Inhalten
    """
    
    @staticmethod
    def generate_sprint_markdown(project, sprint, team_capacity_data=None, kapa_data=None):
        """
        Generiert Markdown-Text für einen Sprint mit Teamauslastung und Kalender
        
        Args:
            project (dict): Das Projekt, zu dem der Sprint gehört
            sprint (dict): Der Sprint, für den Markdown generiert werden soll
            team_capacity_data (list): Liste mit Tupeln (member_id, hours, capacity)
            kapa_data (list): Liste mit Tupeln (id, datum, stunden, kapazität)
            
        Returns:
            str: Der generierte Markdown-Text
        """
        if not project or not sprint:
            return "Kein Sprint oder Projekt ausgewählt."
        
        # Header mit Projekt und Sprint
        markdown = f"# {project.get('name', 'Projekt')} - {sprint.get('sprint_name', 'Sprint')}\n\n"
        
        # Sprint Details
        markdown += "## Sprint Details\n\n"
        markdown += f"* **Zeitraum:** {sprint.get('start', '')} - {sprint.get('ende', '')}\n"
        
        # Kapazitäts- und Story-Point-Berechnungen
        if team_capacity_data:
            total_capacity = sum(capacity for _, _, capacity in team_capacity_data)
            umrechnungsfaktor = 1.5  # Fester Umrechnungsfaktor
            berechnete_sp = total_capacity / umrechnungsfaktor if total_capacity > 0 else 0
            
            markdown += f"* **Gesamtkapazität:** {total_capacity:.2f}\n"
            markdown += f"* **Umrechnungsfaktor:** {umrechnungsfaktor}\n"
            markdown += f"* **Berechnete Story Points:** {berechnete_sp:.1f}\n\n"
        
        # Team Auslastung mit Checkbox für Bestätigung
        markdown += "## Team Auslastung\n\n"
        
        # Header der Tabelle
        markdown += "| Name | Kapazität | Bestätigt |\n"
        markdown += "|------|-----------|----------|\n"
        
        # Teammitglieder zur Tabelle hinzufügen
        team_members = project.get('teilnehmer', [])
        
        # Dictionary für schnelleren Zugriff auf Kapazitätsdaten
        capacity_dict = {}
        if team_capacity_data:
            for member_id, _, capacity in team_capacity_data:
                capacity_dict[member_id] = capacity
        
        # Teammitglieder zur Tabelle hinzufügen
        for member in team_members:
            member_id = member.get('id', '')
            member_name = member.get('name', 'Unbekannt')
            
            # Kapazität für dieses Mitglied
            capacity = capacity_dict.get(member_id, 0)
            
            # Checkbox für die Bestätigung
            checkbox = "[ ]"
            
            # Zeile zur Tabelle hinzufügen
            markdown += f"| {member_name} | {capacity:.2f} | {checkbox} |\n"
        
        markdown += "\n"
        
        # Sprint-Kalender
        markdown += "## Sprint-Kalender\n\n"
        markdown += "```\n"
        
        # Kalender erstellen, wenn Daten vorhanden sind
        if sprint and team_members and kapa_data:
            # Sprint-Zeitraum ermitteln
            start_date = MarkdownService._parse_date(sprint.get('start', '13.03.2025'))
            end_date = MarkdownService._parse_date(sprint.get('ende', '08.04.2025'))
            
            # Alle Tage zwischen Start und Ende berechnen
            current_date = start_date
            sprint_days = []
            while current_date <= end_date:
                sprint_days.append(current_date)
                current_date += timedelta(days=1)
            
            # Liste der Mitarbeiter-IDs im Projektteam erstellen
            team_member_ids = set()
            for member in team_members:
                member_id = member.get('id', '')
                if member_id:  # Nur gültige IDs hinzufügen
                    team_member_ids.add(member_id)
            
            # Kapazitätsdaten nach Datum und Mitarbeiter organisieren
            daily_capacity = {}
            for member_id, datum_str, hours, capacity in kapa_data:
                # Nur Einträge für Teammitglieder berücksichtigen
                if member_id not in team_member_ids:
                    continue
                    
                try:
                    datum = MarkdownService._parse_date(datum_str)
                    
                    # Prüfen, ob der Tag im Sprint-Zeitraum liegt
                    if start_date <= datum <= end_date:
                        key = (member_id, datum.strftime('%d.%m.%Y'))
                        daily_capacity[key] = capacity
                except:
                    # Bei Fehlern beim Datum-Parsing überspringen
                    continue
            
            # Kalender-Header (Tage)
            header = "Mitarbeiter "
            for day in sprint_days:
                day_name = day.strftime('%a %d.%m.')
                header += f"| {day_name:<10} "
            
            markdown += header + "\n"
            markdown += "-" * len(header) + "\n"
            
            # Kalender-Zeilen (Mitarbeiter und Kapazitäten)
            for member in team_members:
                member_id = member.get('id', '')
                member_name = member.get('name', 'Unbekannt')
                
                # Kürze den Namen, wenn er zu lang ist
                if len(member_name) > 10:
                    short_name = member_name[:9] + "."
                else:
                    short_name = member_name
                
                row = f"{short_name:<10} "
                
                for day in sprint_days:
                    # Prüfen, ob der Tag ein Wochenende ist
                    is_weekend = day.weekday() >= 5  # 5=Samstag, 6=Sonntag
                    
                    if is_weekend:
                        # Wochenende
                        row += f"| {'WE':<10} "
                    else:
                        # Kapazität für diesen Tag und Mitarbeiter suchen
                        key = (member_id, day.strftime('%d.%m.%Y'))
                        capacity = daily_capacity.get(key, None)
                        
                        if capacity is None:
                            # Kein Eintrag
                            row += f"| {'X':<10} "
                        elif capacity >= 1.0:
                            # Volle Kapazität
                            row += f"| {'1.0':<10} "
                        else:
                            # Teilweise Kapazität
                            row += f"| {capacity:.1f} ".ljust(12)
                
                markdown += row + "\n"
        else:
            markdown += "Keine Daten für den Kalender verfügbar.\n"
            
        markdown += "```\n\n"
        
        # Legende für den Kalender
        markdown += "**Legende:**\n"
        markdown += "* **WE** - Wochenende\n"
        markdown += "* **X** - Keine Kapazitätsdaten verfügbar\n"
        markdown += "* **1.0** - Volle Kapazität\n"
        markdown += "* **0.1-0.9** - Teilweise Kapazität\n"
        
        return markdown
    
    @staticmethod
    def _parse_date(date_str):
        """
        Wandelt einen Datums-String in ein datetime-Objekt um
        
        Args:
            date_str: Datums-String im Format DD.MM.YYYY
            
        Returns:
            datetime: Das entsprechende datetime-Objekt
        """
        try:
            if not date_str:
                return datetime.now()
            
            day, month, year = map(int, date_str.split('.'))
            return datetime(year, month, day)
        except:
            # Fallback: Aktuelles Datum
            return datetime.now()