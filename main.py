"""
CAMP - Capacity Analysis & Management Planner
Einstiegspunkt der Anwendung
"""
import os
import sys
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

# Importiere die Hauptanwendungsklasse
from src.core.app import CAMPApp

if __name__ == "__main__":
    # Erstelle das Hauptfenster
    root = ctk.CTk()
    app = CAMPApp(root)
    
    # Starte die Hauptschleife
    root.mainloop()