"""
Automatisches Update-System für CAMP
"""
import os
import sys
import json
import platform
import subprocess
import tempfile
import threading
import time
import urllib.request
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from urllib.error import URLError

# Importiere die Versionsinformationen
try:
    from src.utils.version import __version__, UPDATE_URL, DOWNLOAD_BASE_URL
except ImportError:
    # Fallback für den Fall, dass das Modul nicht gefunden wird
    __version__ = "0.1.0"
    UPDATE_URL = "https://api.github.com/repos/YOURUSERNAME/CAMP/releases/latest"
    DOWNLOAD_BASE_URL = "https://github.com/YOURUSERNAME/CAMP/releases/download/"

class UpdateStatus:
    """Status-Konstanten für den Update-Prozess"""
    CHECKING = "checking"
    NO_UPDATE = "no_update"
    UPDATE_AVAILABLE = "update_available"
    DOWNLOADING = "downloading"
    READY_TO_INSTALL = "ready_to_install"
    ERROR = "error"

class UpdaterService:
    """Service für die Prüfung und Installation von Updates"""
    
    def __init__(self):
        self.current_version = __version__
        self.latest_version = None
        self.latest_release_url = None
        self.download_url = None
        self.release_notes = None
        self.status = UpdateStatus.CHECKING
        self.error_message = None
        self.downloaded_file = None
    
    def check_for_updates(self):
        """
        Prüft, ob Updates verfügbar sind
        
        Returns:
            bool: True wenn Updates verfügbar sind, False wenn nicht
        """
        try:
            self.status = UpdateStatus.CHECKING
            
            # GitHub API für die neueste Release abfragen
            req = urllib.request.Request(
                UPDATE_URL,
                headers={"Accept": "application/vnd.github.v3+json", "User-Agent": "CAMP-App"}
            )
            
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode())
                
                # Neueste Version extrahieren (ohne 'v' Präfix)
                self.latest_version = data["tag_name"].lstrip("v")
                self.release_notes = data["body"]
                
                # Prüfe, ob die neueste Version neuer ist als die aktuelle
                if self._is_newer_version(self.latest_version, self.current_version):
                    self.status = UpdateStatus.UPDATE_AVAILABLE
                    
                    # Download-URL für die passende Plattform ermitteln
                    platform_name = platform.system().lower()
                    
                    for asset in data["assets"]:
                        asset_name = asset["name"].lower()
                        if platform_name == "windows" and asset_name.endswith(".exe"):
                            self.download_url = asset["browser_download_url"]
                            break
                        elif platform_name == "darwin" and asset_name.endswith(".dmg"):
                            self.download_url = asset["browser_download_url"]
                            break
                    
                    if not self.download_url:
                        self.status = UpdateStatus.ERROR
                        self.error_message = "Keine kompatible Download-Version gefunden"
                        return False
                    
                    return True
                else:
                    self.status = UpdateStatus.NO_UPDATE
                    return False
        
        except URLError as e:
            self.status = UpdateStatus.ERROR
            self.error_message = f"Netzwerkfehler: {str(e)}"
            return False
        except Exception as e:
            self.status = UpdateStatus.ERROR
            self.error_message = f"Fehler beim Prüfen auf Updates: {str(e)}"
            return False
    
    def download_update(self, progress_callback=None):
        """
        Lädt das Update herunter
        
        Args:
            progress_callback: Optionale Callback-Funktion für Fortschrittsanzeige
            
        Returns:
            bool: True wenn der Download erfolgreich war, False wenn nicht
        """
        if self.status != UpdateStatus.UPDATE_AVAILABLE or not self.download_url:
            return False
        
        try:
            self.status = UpdateStatus.DOWNLOADING
            
            # Temporäre Datei erstellen
            temp_dir = tempfile.gettempdir()
            file_name = os.path.basename(self.download_url)
            self.downloaded_file = os.path.join(temp_dir, file_name)
            
            def report_progress(block_num, block_size, total_size):
                if progress_callback:
                    progress = block_num * block_size / total_size
                    progress_callback(min(progress, 1.0))
            
            # Download mit Fortschrittsanzeige
            urllib.request.urlretrieve(
                self.download_url,
                self.downloaded_file,
                reporthook=report_progress if progress_callback else None
            )
            
            self.status = UpdateStatus.READY_TO_INSTALL
            return True
            
        except Exception as e:
            self.status = UpdateStatus.ERROR
            self.error_message = f"Fehler beim Herunterladen des Updates: {str(e)}"
            return False
    
    def install_update(self):
        """
        Installiert das heruntergeladene Update
        
        Returns:
            bool: True wenn die Installation gestartet wurde, False wenn nicht
        """
        if self.status != UpdateStatus.READY_TO_INSTALL or not self.downloaded_file:
            return False
        
        try:
            # Je nach Plattform unterschiedlich starten
            platform_name = platform.system().lower()
            
            if platform_name == "windows":
                # Startet die .exe-Datei als neuen Prozess
                subprocess.Popen([self.downloaded_file], close_fds=True)
            elif platform_name == "darwin":
                # Öffnet die .dmg-Datei
                subprocess.Popen(["open", self.downloaded_file], close_fds=True)
            else:
                # Für andere Plattformen
                self.status = UpdateStatus.ERROR
                self.error_message = "Automatische Installation wird für diese Plattform nicht unterstützt"
                return False
            
            # Beende die aktuelle Anwendung
            time.sleep(1)  # Kurze Verzögerung, damit der Installationsprozess starten kann
            sys.exit(0)
            
            return True
            
        except Exception as e:
            self.status = UpdateStatus.ERROR
            self.error_message = f"Fehler beim Installieren des Updates: {str(e)}"
            return False
    
    def check_for_updates_async(self, callback=None):
        """
        Prüft asynchron auf Updates
        
        Args:
            callback: Callback-Funktion, die aufgerufen wird, wenn die Prüfung abgeschlossen ist
        """
        def check_thread():
            result = self.check_for_updates()
            if callback:
                callback(result)
        
        thread = threading.Thread(target=check_thread)
        thread.daemon = True
        thread.start()
    
    def _is_newer_version(self, version1, version2):
        """
        Prüft, ob version1 neuer ist als version2
        
        Args:
            version1: Erste Version (z.B. "1.2.3")
            version2: Zweite Version (z.B. "1.2.0")
            
        Returns:
            bool: True wenn version1 neuer ist als version2
        """
        v1_parts = list(map(int, version1.split('.')))
        v2_parts = list(map(int, version2.split('.')))
        
        # Stelle sicher, dass beide Listen gleich lang sind
        while len(v1_parts) < len(v2_parts):
            v1_parts.append(0)
        while len(v2_parts) < len(v1_parts):
            v2_parts.append(0)
        
        # Vergleiche die Versionen
        for i in range(len(v1_parts)):
            if v1_parts[i] > v2_parts[i]:
                return True
            elif v1_parts[i] < v2_parts[i]:
                return False
        
        # Gleiche Version
        return False


class UpdaterDialog(ctk.CTkToplevel):
    """Dialog für die Anzeige und Installation von Updates"""
    
    def __init__(self, master, updater_service):
        """
        Initialisiert den Update-Dialog
        
        Args:
            master: Das übergeordnete Widget
            updater_service: Der UpdaterService
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Attribute
        self.updater_service = updater_service
        self.progress_value = 0
        
        # Fenster-Konfiguration
        self.title("CAMP Update")
        self.geometry("500x400")
        self.resizable(False, False)
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        self.after(100, self.deiconify)
        
        # Status aktualisieren
        self._update_ui_status()
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für den Dialog"""
        # Titel-Label
        self.title_label = ctk.CTkLabel(
            self, 
            text="CAMP Update",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        
        # Info-Frame
        self.info_frame = ctk.CTkFrame(self)
        
        # Aktuelle Version
        self.current_version_label = ctk.CTkLabel(
            self.info_frame, 
            text=f"Installierte Version: {self.updater_service.current_version}",
            anchor="w"
        )
        
        # Neueste Version
        self.latest_version_label = ctk.CTkLabel(
            self.info_frame, 
            text="Neueste Version: Wird geprüft...",
            anchor="w"
        )
        
        # Status-Label
        self.status_label = ctk.CTkLabel(
            self, 
            text="Prüfe auf Updates...",
            font=ctk.CTkFont(size=14)
        )
        
        # Fortschrittsbalken
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.set(0)
        
        # Release Notes Frame
        self.notes_frame = ctk.CTkFrame(self)
        
        # Release Notes Label
        self.notes_label = ctk.CTkLabel(
            self.notes_frame, 
            text="Release Notes:",
            anchor="w",
            font=ctk.CTkFont(weight="bold")
        )
        
        # Release Notes Textbox
        self.notes_textbox = ctk.CTkTextbox(
            self.notes_frame,
            width=460,
            height=150
        )
        self.notes_textbox.configure(state="disabled")
        
        # Button-Frame
        self.button_frame = ctk.CTkFrame(self)
        
        # Schließen-Button
        self.close_button = ctk.CTkButton(
            self.button_frame, 
            text="Schließen",
            command=self.destroy,
            width=120
        )
        
        # Update-Button (standardmäßig ausgeblendet)
        self.update_button = ctk.CTkButton(
            self.button_frame, 
            text="Update installieren",
            command=self._download_and_install,
            width=160,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3")
        )
        
        # Automatisch auf Updates prüfen
        self.check_for_updates()
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Label
        self.title_label.pack(pady=(20, 10))
        
        # Info-Frame
        self.info_frame.pack(fill="x", padx=20, pady=10)
        self.current_version_label.pack(anchor="w", padx=10, pady=2)
        self.latest_version_label.pack(anchor="w", padx=10, pady=2)
        
        # Status-Label
        self.status_label.pack(pady=5)
        
        # Fortschrittsbalken
        self.progress_bar.pack(fill="x", padx=20, pady=5)
        
        # Release Notes Frame
        self.notes_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.notes_label.pack(anchor="w", padx=0, pady=(0, 5))
        self.notes_textbox.pack(fill="both", expand=True)
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=20, pady=15)
        self.close_button.pack(side="left", padx=10)
        # Update-Button wird erst angezeigt, wenn ein Update verfügbar ist
    
    def check_for_updates(self):
        """Prüft auf verfügbare Updates"""
        def on_check_complete(update_available):
            if update_available:
                # Update verfügbar
                self.latest_version_label.configure(
                    text=f"Neueste Version: {self.updater_service.latest_version}"
                )
                self.status_label.configure(
                    text="Ein Update ist verfügbar!",
                    text_color=("green3", "green4")
                )
                
                # Release Notes anzeigen
                if self.updater_service.release_notes:
                    self.notes_textbox.configure(state="normal")
                    self.notes_textbox.delete("1.0", "end")
                    self.notes_textbox.insert("1.0", self.updater_service.release_notes)
                    self.notes_textbox.configure(state="disabled")
                
                # Update-Button anzeigen
                self.update_button.pack(side="right", padx=10)
            
            elif self.updater_service.status == UpdateStatus.NO_UPDATE:
                # Kein Update verfügbar
                self.latest_version_label.configure(
                    text=f"Neueste Version: {self.updater_service.latest_version}"
                )
                self.status_label.configure(
                    text="Sie haben bereits die neueste Version!",
                    text_color=("gray60", "gray70")
                )
                
                # Release Notes löschen
                self.notes_textbox.configure(state="normal")
                self.notes_textbox.delete("1.0", "end")
                self.notes_textbox.insert("1.0", "Keine neuen Änderungen verfügbar.")
                self.notes_textbox.configure(state="disabled")
            
            else:
                # Fehler bei der Überprüfung
                self.latest_version_label.configure(
                    text="Neueste Version: Unbekannt"
                )
                self.status_label.configure(
                    text=f"Fehler: {self.updater_service.error_message}",
                    text_color=("red", "red")
                )
            
            # UI-Status aktualisieren
            self._update_ui_status()
        
        # Asynchron auf Updates prüfen
        self.updater_service.check_for_updates_async(on_check_complete)
    
    def _download_and_install(self):
        """Lädt das Update herunter und installiert es"""
        # UI-Elemente blockieren
        self.update_button.configure(state="disabled")
        self.close_button.configure(state="disabled")
        
        # Fortschrittsanzeige anzeigen
        self.progress_bar.set(0)
        self.status_label.configure(
            text="Update wird heruntergeladen...",
            text_color=("blue", "light blue")
        )
        
        def on_progress(progress):
            self.progress_bar.set(progress)
            if progress >= 1.0:
                self.status_label.configure(
                    text="Download abgeschlossen. Installation wird gestartet...",
                    text_color=("green3", "green4")
                )
        
        def download_thread():
            success = self.updater_service.download_update(progress_callback=on_progress)
            
            if success:
                # Kurze Verzögerung, dann installieren
                self.after(1000, self._install_update)
            else:
                # Fehler beim Download
                self.status_label.configure(
                    text=f"Fehler beim Download: {self.updater_service.error_message}",
                    text_color=("red", "red")
                )
                self.update_button.configure(state="normal")
                self.close_button.configure(state="normal")
        
        # Download-Thread starten
        thread = threading.Thread(target=download_thread)
        thread.daemon = True
        thread.start()
    
    def _install_update(self):
        """Installiert das heruntergeladene Update"""
        # Bestätigungsdialog anzeigen
        confirm = messagebox.askyesno(
            "CAMP Update",
            "Das Update wurde heruntergeladen und ist bereit zur Installation. "
            "Die Anwendung wird beendet und das Update installiert. Fortfahren?",
            parent=self
        )
        
        if confirm:
            self.status_label.configure(
                text="Update wird installiert...",
                text_color=("green3", "green4")
            )
            
            # Update installieren
            self.updater_service.install_update()
        else:
            self.status_label.configure(
                text="Update wurde heruntergeladen, aber die Installation wurde abgebrochen.",
                text_color=("orange", "dark orange")
            )
            self.update_button.configure(state="normal", text="Installation starten")
            self.close_button.configure(state="normal")
    
    def _update_ui_status(self):
        """Aktualisiert die UI basierend auf dem aktuellen Status"""
        status = self.updater_service.status
        
        if status == UpdateStatus.CHECKING:
            self.progress_bar.configure(mode="indeterminate")
            self.progress_bar.start()
        else:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
            
            if status == UpdateStatus.NO_UPDATE:
                self.progress_bar.set(1)
            elif status == UpdateStatus.ERROR:
                self.progress_bar.set(0)
            else:
                self.progress_bar.set(0)
    
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


def check_for_updates_on_startup(root, show_no_update=False):
    """
    Prüft beim Start der Anwendung auf Updates
    
    Args:
        root: Das Root-Widget der Anwendung
        show_no_update: Ob eine Meldung angezeigt werden soll, wenn kein Update verfügbar ist
    """
    updater_service = UpdaterService()
    
    def on_check_complete(update_available):
        if update_available:
            # Update verfügbar - Dialog anzeigen
            show_update_dialog(root, updater_service)
        elif show_no_update:
            # Kein Update verfügbar, aber Meldung gewünscht
            messagebox.showinfo(
                "CAMP Update",
                "Sie haben bereits die neueste Version!",
                parent=root
            )
    
    # Asynchron auf Updates prüfen
    updater_service.check_for_updates_async(on_check_complete)


def show_update_dialog(root, updater_service=None):
    """
    Zeigt den Update-Dialog an
    
    Args:
        root: Das Root-Widget der Anwendung
        updater_service: Optional - ein vorhandener UpdaterService
    """
    if updater_service is None:
        updater_service = UpdaterService()
    
    update_dialog = UpdaterDialog(root, updater_service)
    return update_dialog