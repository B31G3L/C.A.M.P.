"""
Hilfsfunktionen für Datumsoperationen
"""
from datetime import datetime, timedelta

def parse_date(date_str):
    """
    Wandelt einen Datums-String in ein datetime-Objekt um
    
    Args:
        date_str (str): Datums-String im Format "DD.MM.YYYY"
        
    Returns:
        datetime: Das entsprechende datetime-Objekt
    """
    try:
        if not date_str:
            return None
            
        day, month, year = map(int, date_str.split('.'))
        return datetime(year, month, day)
    except (ValueError, TypeError):
        return None

def format_date(date_obj):
    """
    Formatiert ein datetime-Objekt als String
    
    Args:
        date_obj (datetime): Das zu formatierende datetime-Objekt
        
    Returns:
        str: Der formatierte String im Format "DD.MM.YYYY"
    """
    if not date_obj:
        return ""
        
    return date_obj.strftime("%d.%m.%Y")

def get_date_range(start_date, end_date):
    """
    Gibt alle Daten zwischen start_date und end_date zurück
    
    Args:
        start_date (datetime): Startdatum
        end_date (datetime): Enddatum
        
    Returns:
        list: Liste der datetime-Objekte im Bereich
    """
    if not start_date or not end_date:
        return []
        
    dates = []
    current_date = start_date
    
    while current_date <= end_date:
        dates.append(current_date)
        current_date += timedelta(days=1)
    
    return dates

def is_weekend(date_obj):
    """
    Prüft, ob ein Datum ein Wochenende ist
    
    Args:
        date_obj (datetime): Das zu prüfende Datum
        
    Returns:
        bool: True wenn Wochenende, sonst False
    """
    if not date_obj:
        return False
        
    return date_obj.weekday() >= 5  # 5=Samstag, 6=Sonntag