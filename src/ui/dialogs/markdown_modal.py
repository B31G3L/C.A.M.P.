"""
Modal-Fenster zur Anzeige und zum Kopieren von Markdown-Inhalt
"""
import tkinter as tk
import customtkinter as ctk
from typing import Optional

class MarkdownModal(ctk.CTkToplevel):
    """
    Modal-Fenster zur Anzeige von generiertem Markdown-Inhalt
    mit Möglichkeit zum Kopieren in die Zwischenablage
    """
    def __init__(self, master, markdown_content: str, title: str = "Markdown"):
        """
        Initialisiert das Modal-Fenster
        
        Args:
            master: Das übergeordnete Widget
            markdown_content: Der anzuzeigende Markdown-Inhalt
            title: Der Titel des Fensters
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title(title)
        self.geometry("800x600")
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.markdown_content = markdown_content
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Inhalt einfügen
        self.markdown_textbox.insert("1.0", markdown_content)
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        # Titel-Frame
        self.title_frame = ctk.CTkFrame(self)
        self.title_label = ctk.CTkLabel(
            self.title_frame, 
            text="Markdown Preview",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        # Info-Label
        self.info_label = ctk.CTkLabel(
            self.title_frame,
            text="Dieser Markdown-Text kann in Confluence eingefügt werden",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray70")
        )
        
        # Textbox für Markdown-Inhalt
        self.markdown_textbox = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(family="Courier", size=12),
            wrap="none"  # Keine Zeilenumbrüche für bessere Formatierung
        )
        
        # Button-Frame
        self.button_frame = ctk.CTkFrame(self)
        
        # Copy-Button
        self.copy_button = ctk.CTkButton(
            self.button_frame,
            text="In Zwischenablage kopieren",
            command=self._copy_to_clipboard,
            width=200
        )
        
        # Schließen-Button
        self.close_button = ctk.CTkButton(
            self.button_frame,
            text="Schließen",
            command=self.destroy,
            width=100
        )
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Frame
        self.title_frame.pack(fill="x", padx=20, pady=(20, 0))
        self.title_label.pack(side="left", pady=10)
        self.info_label.pack(side="right", padx=10, pady=10)
        
        # Markdown-Textbox (nimmt den meisten Platz ein)
        self.markdown_textbox.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=20, pady=(0, 20))
        self.copy_button.pack(side="left", padx=10)
        self.close_button.pack(side="right", padx=10)
    
    def _copy_to_clipboard(self):
        """Kopiert den Markdown-Inhalt in die Zwischenablage"""
        self.clipboard_clear()
        self.clipboard_append(self.markdown_content)
        
        # Kurze Bestätigung anzeigen (temporär)
        original_text = self.copy_button.cget("text")
        self.copy_button.configure(text="✓ Kopiert!", fg_color=("green", "dark green"))
        
        # Nach 1,5 Sekunden zurücksetzen
        self.after(1500, lambda: self.copy_button.configure(
            text=original_text,
            fg_color=None  # Zurück zur Standardfarbe
        ))
    
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