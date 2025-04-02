"""
CAMP - Capacity Analysis & Management Planner
Einstiegspunkt der Anwendung
"""
import os
import sys
import csv  # CSV-Modul importieren

import customtkinter as ctk
from src.utils.config_loader import load_config
from src.utils.version import __version__, __app_name__

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

# Prüfe auf Single-Instance (nur eine Instanz der Anwendung erlauben)
def is_app_already_running():
    """Prüft, ob bereits eine Instanz der Anwendung läuft"""
    import tempfile
    import platform
    import socket
    import os
    
    temp_dir = tempfile.gettempdir()
    lock_file_path = os.path.join(temp_dir, "camp_app.lock")
    
    # Unterschiedliche Implementierung je nach Betriebssystem
    system = platform.system().lower()
    
    if system == "windows":
        # Windows-spezifische Implementierung mit Named Mutex
        try:
            import win32event
            import win32api
            import winerror
            
            mutex_name = "Global\\CAMP_Application_Mutex"
            mutex = win32event.CreateMutex(None, 1, mutex_name)
            if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
                win32api.CloseHandle(mutex)
                return True
            return False
        except ImportError:
            # win32api nicht verfügbar, Fallback auf Socket
            pass
    
    # Für macOS und Linux (oder als Fallback)
    try:
        # Versuche, eine TCP-Socket-Bindung zu erstellen
        # Port 0 bedeutet, dass das Betriebssystem einen freien Port auswählt
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # SO_REUSEADDR erlaubt das Wiederverwenden des Ports nach einem Absturz
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(('localhost', 12345))  # Verwende einen spezifischen Port für die App
        # Socket bleibt offen, wird beim Beenden geschlossen
        return False
    except socket.error:
        # Port wird bereits verwendet, vermutlich von einer anderen Instanz
        return True

# Importiere die Hauptanwendungsklasse
from src.core.app import CAMPApp

def main():
    """Hauptfunktion der Anwendung"""
    # Prüfe, ob bereits eine Instanz läuft
    if is_app_already_running():
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "CAMP bereits gestartet",
            "Eine Instanz von CAMP läuft bereits. Es kann nur eine Instanz gleichzeitig ausgeführt werden."
        )
        sys.exit(0)
    
    # Erstelle das Hauptfenster
    root = ctk.CTk()
    app = CAMPApp(root)
    
    # Verzögerte Ausführung der Startsequenz nach dem Start der Anwendungsschleife
    # Dies zeigt das Info-Fenster an und aktualisiert die Daten
    root.after(100, app.show_startup_info)
    
    # Starte die Hauptschleife
    root.mainloop()
    
    # Bereinigung beim Beenden
    try:
        import tempfile
        socket_file = os.path.join(tempfile.gettempdir(), "camp_app_socket")
        if os.path.exists(socket_file):
            os.unlink(socket_file)
    except:
        pass

if __name__ == "__main__":
    # Versionsinformationen anzeigen
    print(f"{__app_name__} v{__version__}")
    print("Starting application...")
    
    # Hauptfunktion aufrufen
    main()