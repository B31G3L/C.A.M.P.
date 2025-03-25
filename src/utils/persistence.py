"""
Persistenz-Hilfsfunktionen für CAMP
"""
import os
import json
import csv
import shutil

def ensure_directory(directory):
    """
    Stellt sicher, dass ein Verzeichnis existiert
    
    Args:
        directory (str): Pfad zum Verzeichnis
    """
    os.makedirs(directory, exist_ok=True)

def save_json(data, file_path, create_backup=True):
    """
    Speichert Daten in eine JSON-Datei
    
    Args:
        data: Die zu speichernden Daten
        file_path (str): Pfad zur Zieldatei
        create_backup (bool): Ob ein Backup erstellt werden soll
        
    Returns:
        bool: True wenn erfolgreich, False wenn ein Fehler aufgetreten ist
    """
    try:
        # Stelle sicher, dass das Verzeichnis existiert
        ensure_directory(os.path.dirname(file_path))
        
        # Erstelle ein Backup, falls gewünscht
        if create_backup and os.path.exists(file_path):
            backup_path = f"{file_path}.bak"
            shutil.copy2(file_path, backup_path)
        
        # Speichere die Daten
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        
        return True
    except Exception as e:
        print(f"Fehler beim Speichern der JSON-Datei: {e}")
        return False

def load_json(file_path, default=None):
    """
    Lädt Daten aus einer JSON-Datei
    
    Args:
        file_path (str): Pfad zur Quelldatei
        default: Standardwert, wenn die Datei nicht existiert
        
    Returns:
        Die geladenen Daten oder der Standardwert
    """
    try:
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return default
    except Exception as e:
        print(f"Fehler beim Laden der JSON-Datei: {e}")
        return default

def save_csv(data, file_path, delimiter=';', create_backup=True):
    """
    Speichert Daten in eine CSV-Datei
    
    Args:
        data: Die zu speichernden Daten (Liste von Listen)
        file_path (str): Pfad zur Zieldatei
        delimiter (str): Trennzeichen
        create_backup (bool): Ob ein Backup erstellt werden soll
        
    Returns:
        bool: True wenn erfolgreich, False wenn ein Fehler aufgetreten ist
    """
    try:
        # Stelle sicher, dass das Verzeichnis existiert
        ensure_directory(os.path.dirname(file_path))
        
        # Erstelle ein Backup, falls gewünscht
        if create_backup and os.path.exists(file_path):
            backup_path = f"{file_path}.bak"
            shutil.copy2(file_path, backup_path)
        
        # Speichere die Daten
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=delimiter)
            writer.writerows(data)
        
        return True
    except Exception as e:
        print(f"Fehler beim Speichern der CSV-Datei: {e}")
        return False

def load_csv(file_path, delimiter=';', has_header=False):
    """
    Lädt Daten aus einer CSV-Datei
    
    Args:
        file_path (str): Pfad zur Quelldatei
        delimiter (str): Trennzeichen
        has_header (bool): Ob die Datei eine Kopfzeile hat
        
    Returns:
        tuple: (Daten, Header) oder (Daten, None) wenn has_header=False
    """
    try:
        if os.path.exists(file_path):
            with open(file_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=delimiter)
                
                if has_header:
                    header = next(reader)
                    data = list(reader)
                    return data, header
                else:
                    data = list(reader)
                    return data, None
        return [], None
    except Exception as e:
        print(f"Fehler beim Laden der CSV-Datei: {e}")
        return [], None