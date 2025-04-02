"""
About-Dialog für CAMP - Zeigt Informationen über die Anwendung an
"""
import customtkinter as ctk
import webbrowser

class AboutDialog(ctk.CTkToplevel):
    """
    Modal-Fenster zum Anzeigen von Informationen über die Anwendung
    """
    def __init__(self, master, app_name="CAMP", version="0.1.0", release_date="2025-04-02"):
        """
        Initialisiert den About-Dialog
        
        Args:
            master: Das übergeordnete Widget
            app_name: Name der Anwendung
            version: Versionsnummer
            release_date: Release-Datum
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("Über CAMP")
        self.geometry("400x350")
        self.resizable(False, False)
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.app_name = app_name
        self.version = version
        self.release_date = release_date
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für den Dialog"""
        # Logo-Frame (kann später ein Logo enthalten)
        self.logo_frame = ctk.CTkFrame(self, height=100)
        
        # Logo-Label
        self.logo_label = ctk.CTkLabel(
            self.logo_frame,
            text="CAMP",
            font=ctk.CTkFont(size=36, weight="bold")
        )
        
        # Info-Frame
        self.info_frame = ctk.CTkFrame(self)
        
        # App-Name
        self.app_name_label = ctk.CTkLabel(
            self.info_frame,
            text=self.app_name,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        # Version und Datum
        self.version_label = ctk.CTkLabel(
            self.info_frame,
            text=f"Version {self.version} ({self.release_date})"
        )
        
        # Beschreibung
        self.description_label = ctk.CTkLabel(
            self.info_frame,
            text="Capacity Analysis & Management Planner",
            wraplength=350
        )
        
        # Copyright
        self.copyright_label = ctk.CTkLabel(
            self.info_frame,
            text="© 2025 Your Company",
            font=ctk.CTkFont(size=10),
            text_color=("gray60", "gray70")
        )
        
        # Zusätzliche Informationen
        self.additional_info = ctk.CTkLabel(
            self.info_frame,
            text="Diese Anwendung dient zur Verwaltung und Analyse von Projektkapazitäten und Sprints.",
            wraplength=350,
            justify="left"
        )
        
        # Button-Frame
        self.button_frame = ctk.CTkFrame(self)
        
        # GitHub-Button
        self.github_button = ctk.CTkButton(
            self.button_frame,
            text="GitHub",
            command=self._open_github,
            width=120
        )
        
        # Schließen-Button
        self.close_button = ctk.CTkButton(
            self.button_frame,
            text="Schließen",
            command=self.destroy,
            width=120
        )
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Logo-Frame
        self.logo_frame.pack(fill="x", padx=20, pady=(20, 10))
        self.logo_label.pack(pady=20)
        
        # Info-Frame
        self.info_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.app_name_label.pack(pady=(10, 0))
        self.version_label.pack(pady=(0, 10))
        self.description_label.pack(pady=5)
        self.additional_info.pack(pady=(15, 5))
        self.copyright_label.pack(pady=(15, 5))
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=20, pady=(0, 20))
        self.github_button.pack(side="left", padx=10)
        self.close_button.pack(side="right", padx=10)
    
    def _open_github(self):
        """Öffnet die GitHub-Seite des Projekts"""
        webbrowser.open("https://github.com/YOURUSERNAME/CAMP")
    
    def center_window(self):
        """Zentriert das Fenster auf dem Bildschirm"""
        self.update_idletasks()
        
        # Fenstergröße
        width = self.winfo_width()
        height = self.winfo_height()
        
        # Bildschirmgröße
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # Position berechnen
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Fenster positionieren
        self.geometry(f"{width}x{height}+{x}+{y}")