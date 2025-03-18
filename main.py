"""
CAMP - Capacity Analysis & Management Planner
Einstiegspunkt der Anwendung
"""
import os
import sys
import customtkinter as ctk

# Füge den Projektpfad zum Systempfad hinzu
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

from src.app import CAMPApp

if __name__ == "__main__":
    # Anwendungseinstellungen
    ctk.set_appearance_mode("Light")  # "System", "Dark" oder "Light"
    ctk.set_default_color_theme("blue")  # "blue", "dark-blue", "green"
    
    # Erstelle das Hauptfenster
    root = ctk.CTk()
    app = CAMPApp(root)
    
    # Starte die Hauptschleife
    root.mainloop()