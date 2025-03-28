"""
Export-Funktionen für Confluence Wiki und andere Formate
"""
from datetime import datetime, timedelta

class WikiExporter:
    """
    Klasse zum Exportieren von Daten in verschiedene Wiki-Formate
    """
    
    @staticmethod
    def generate_confluence_wiki(main_view):
        """
        Generiert Confluence Wiki-Text mit Sprint-Details, Team-Auslastung und Sprint-Kalender
        
        Args:
            main_view: Die MainView-Instanz mit den anzuzeigenden Daten
            
        Returns:
            str: Der generierte Confluence Wiki-Text
        """
        sprint = main_view.selected_sprint
        project = main_view.selected_project
        
        if not sprint or not project:
            return "Kein Sprint oder Projekt ausgewählt."
        
        # Header mit Projekt und Sprint
        confluence = f"h1. {project.get('name', 'Projekt')} - {sprint.get('sprint_name', 'Sprint')}\n\n"
        
        # Sprint-Kapazitätsdaten abrufen
        sprint_capacity_data = main_view._get_sprint_capacity_data() if hasattr(main_view, '_get_sprint_capacity_data') else []
        
        # Berechne die Gesamtkapazität und geschätzte Story Points
        total_capacity = sum(capacity for _, _, capacity in sprint_capacity_data)
        umrechnungsfaktor = 1.5  # Fester Umrechnungsfaktor
        berechnete_sp = total_capacity / umrechnungsfaktor if total_capacity > 0 else 0
        
        # Sprint Details mit den gewünschten Informationen
        confluence += "h2. Sprint Details\n\n"
        confluence += f"* *Zeitraum:* {sprint.get('start', '')} - {sprint.get('ende', '')}\n"
        confluence += f"* *Gesamtkapazität:* {total_capacity:.2f}\n"
        confluence += f"* *Umrechnungsfaktor:* {umrechnungsfaktor}\n"
        confluence += f"* *Berechnete Story Points:* {berechnete_sp:.1f}\n\n"
        
        # Team Auslastung
        confluence += "h2. Team Auslastung\n\n"
        
        # Tabelle mit Team-Auslastung - einfache Confluence Wiki-Tabellensyntax
        # Header der Tabelle
        confluence += "|| Name || Kapazität || Bestätigt ||\n"
        
        # Kapazitätsdaten für den Sprint laden
        team_members = project.get('teilnehmer', [])
        
        # Dictionary für schnelleren Zugriff
        capacity_dict = {}
        for member_id, hours, capacity in sprint_capacity_data:
            capacity_dict[member_id] = (hours, capacity)
        
        # Teammitglieder zur Tabelle hinzufügen
        for member in team_members:
            member_id = member.get('id', '')
            member_name = member.get('name', 'Unbekannt')
            
            # Kapazität für dieses Mitglied
            hours, capacity = capacity_dict.get(member_id, (0, 0))
            
            # Zeile zur Tabelle hinzufügen (ohne spezielle Formatierung)
            confluence += f"| {member_name} | {capacity:.2f} | [ ] |\n"
        
        confluence += "\n"
        
        # Sprint-Kalender
        confluence += "h2. Sprint-Kalender\n\n"
        
        # Kalender implementieren
        
        # Sprint-Zeitraum ermitteln
        start_date = WikiExporter._parse_date(sprint.get('start', '01.01.2000'))
        end_date = WikiExporter._parse_date(sprint.get('ende', '01.01.2000'))
        
        # Alle Tage zwischen Start und Ende berechnen
        current_date = start_date
        sprint_days = []
        while current_date <= end_date:
            sprint_days.append(current_date)
            current_date += timedelta(days=1)
        
        # Liste der Mitarbeiter-IDs im Projektteam erstellen
        team_member_ids = {}
        for member in team_members:
            member_id = member.get('id', '')
            if member_id:  # Nur gültige IDs hinzufügen
                team_member_ids[member_id] = member.get('name', 'Unbekannt')
        
        # Kapazitätsdaten nach Datum und Mitarbeiter organisieren
        daily_capacity = {}
        for member_id, datum_str, hours, capacity in main_view.kapa_data:
            # Nur Einträge für Teammitglieder berücksichtigen
            if member_id not in team_member_ids:
                continue
                
            try:
                datum = WikiExporter._parse_date(datum_str)
                
                # Prüfen, ob der Tag im Sprint-Zeitraum liegt
                if start_date <= datum <= end_date:
                    key = (member_id, datum.strftime('%d.%m.%Y'))
                    daily_capacity[key] = capacity
            except:
                # Bei Fehlern beim Datum-Parsing überspringen
                continue
        
        # Kalender-Tabelle erstellen
        # Header mit Mitarbeitern und Tagen
        confluence += "|| Mitarbeiter"
        
        # Tagesnamen als Header
        for day in sprint_days:
            day_name = day.strftime('%a %d.%m.')
            confluence += f" || {day_name}"
        
        confluence += " ||\n"
        
        # Zeilen für jeden Mitarbeiter
        for member_id, member_name in team_member_ids.items():
            confluence += f"| {member_name}"
            
            # Spalten für jeden Tag
            for day in sprint_days:
                # Prüfen, ob der Tag ein Wochenende ist
                is_weekend = day.weekday() >= 5  # 5=Samstag, 6=Sonntag
                
                # Kapazität für diesen Tag und Mitarbeiter suchen
                key = (member_id, day.strftime('%d.%m.%Y'))
                capacity = daily_capacity.get(key, None)
                
                if is_weekend:
                    confluence += " | {panel:bgColor=#E0E0E0}WE{panel}"
                elif capacity is None:
                    confluence += " | {panel:bgColor=#FFD2D2}X{panel}"
                elif capacity >= 1.0:
                    confluence += f" | {{panel:bgColor=#D5F5D5}}{capacity:.1f}{{panel}}"
                else:
                    confluence += f" | {{panel:bgColor=#FFF2CC}}{capacity:.1f}{{panel}}"
            
            confluence += " |\n"
        
        return confluence
    
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