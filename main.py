"""
CAMP - Capacity Analysis & Management Planner
Einstiegspunkt der Anwendung
"""
import os
import sys
import csv  # CSV-Modul importieren

import customtkinter as ctk
from src.utils.config_loader import load_config

# Konfiguration laden
config = load_config()

# Anwendungseinstellungen
ctk.set_appearance_mode(config.DEFAULT_THEME)
ctk.set_default_color_theme(config.DEFAULT_COLOR_SCHEME)

# Stelle sicher, dass die Datenverzeichnisse existieren
os.makedirs(config.DATA_DIR, exist_ok=True)
os.makedirs(config.BACKUP_DIR, exist_ok=True)
os.makedirs(config.EXPORT_DIR, exist_ok=True)
kapa_csv_path = os.path.join(config.DATA_DIR, "kapa_data.csv")
if not os.path.exists(kapa_csv_path):
    with open(kapa_csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(["ID", "DATUM", "STUNDEN", "KAPAZITÄT"])

# Importiere die Hauptanwendungsklasse
from src.core.app import CAMPApp

if __name__ == "__main__":
    # Erstelle das Hauptfenster
    root = ctk.CTk()
    app = CAMPApp(root)
    
    # Verzögerte Ausführung der Startsequenz nach dem Start der Anwendungsschleife
    # Dies zeigt das Info-Fenster an und aktualisiert die Daten
    root.after(100, app.show_startup_info)
    
    # Starte die Hauptschleife
    root.mainloop()