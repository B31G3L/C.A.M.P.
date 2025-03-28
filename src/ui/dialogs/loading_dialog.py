"""
Modal-Fenster zum Anzeigen eines Ladezustands
"""
import customtkinter as ctk

class LoadingDialog(ctk.CTkToplevel):
    """
    Modal-Fenster, das angezeigt wird, während die Anwendung Daten lädt oder verarbeitet
    """
    def __init__(self, master, message="Bitte warten..."):
        """
        Initialisiert das Modal-Fenster
        
        Args:
            master: Das übergeordnete Widget
            message: Die anzuzeigende Nachricht
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("")
        self.geometry("300x100")
        self.resizable(False, False)
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.message = message
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Sofort anzeigen (ohne Verzögerung)
        self.deiconify()
        self.update()
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        # Container-Frame
        self.container = ctk.CTkFrame(self)
        
        # Nachricht
        self.message_label = ctk.CTkLabel(
            self.container,
            text=self.message,
            font=ctk.CTkFont(size=14)
        )
        
        # Fortschrittsanzeige
        self.progress_bar = ctk.CTkProgressBar(self.container)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Container
        self.container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Nachricht
        self.message_label.pack(pady=(0, 10))
        
        # Fortschrittsanzeige
        self.progress_bar.pack(fill="x")
    
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