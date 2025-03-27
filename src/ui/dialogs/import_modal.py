"""
Modal-Fenster zum Importieren von Daten mit Tabs für BW Tool Daten und Arbeitszeiterfassung
"""
import os
import customtkinter as ctk
from typing import Dict, Callable, Optional, List
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import json
import shutil
import datetime
import re


class ProjectMemberSelectionDialog(ctk.CTkToplevel):
    """Dialog zur Auswahl eines Projektmitglieds"""
    
    def __init__(self, master, projects_data, on_select=None):
        """
        Initialisiert den Dialog
        
        Args:
            master: Das übergeordnete Widget
            projects_data: Die Projektdaten aus der project.json
            on_select: Callback-Funktion, die aufgerufen wird, wenn ein Mitglied ausgewählt wird
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("Projektmitglied auswählen")
        self.geometry("500x400")
        self.resizable(False, False)
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.projects_data = projects_data
        self.on_select = on_select
        self.selected_member = None
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        self._load_projects()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen
        self.after(100, self.deiconify)
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für den Dialog"""
        # Titel-Label
        self.title_label = ctk.CTkLabel(
            self, 
            text="Wählen Sie ein Projektmitglied aus",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        
        # Projekt-Auswahl
        self.project_frame = ctk.CTkFrame(self)
        self.project_label = ctk.CTkLabel(
            self.project_frame,
            text="Projekt:",
            width=100
        )
        self.project_var = ctk.StringVar(value="")
        self.project_dropdown = ctk.CTkOptionMenu(
            self.project_frame,
            values=[""],
            variable=self.project_var,
            command=self._on_project_selected,
            width=300
        )
        
        # Mitglieder-Listen-Frame
        self.members_label = ctk.CTkLabel(
            self,
            text="Projektmitglieder:",
            anchor="w"
        )
        
        # Scrollbare Liste für Mitglieder
        self.members_list_frame = ctk.CTkScrollableFrame(
            self,
            width=450,
            height=250
        )
        
        # Info-Label
        self.info_label = ctk.CTkLabel(
            self,
            text="",
            text_color=("gray60", "gray70")
        )
        
        # Button-Frame
        self.button_frame = ctk.CTkFrame(self)
        
        # Abbrechen-Button
        self.cancel_button = ctk.CTkButton(
            self.button_frame,
            text="Abbrechen",
            command=self.destroy,
            width=120,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40")
        )
        
        # Auswählen-Button (initial deaktiviert)
        self.select_button = ctk.CTkButton(
            self.button_frame,
            text="Auswählen",
            command=self._on_select_clicked,
            width=120,
            state="disabled"
        )
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Label
        self.title_label.pack(pady=(20, 15))
        
        # Projekt-Auswahl
        self.project_frame.pack(fill="x", padx=20, pady=(0, 15))
        self.project_label.pack(side="left")
        self.project_dropdown.pack(side="left", padx=(10, 0))
        
        # Mitglieder-Label
        self.members_label.pack(anchor="w", padx=20, pady=(0, 5))
        
        # Mitglieder-Liste
        self.members_list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        # Info-Label
        self.info_label.pack(padx=20, pady=(0, 10))
        
        # Button-Frame
        self.button_frame.pack(fill="x", padx=20, pady=(0, 20))
        self.cancel_button.pack(side="left", padx=(0, 10))
        self.select_button.pack(side="right")
    
    def _load_projects(self):
        """Lädt die Projekte in das Dropdown"""
        if not self.projects_data or "projekte" not in self.projects_data:
            self.info_label.configure(
                text="Keine Projekte gefunden. Bitte importieren Sie zuerst Projektdaten.",
                text_color="red"
            )
            return
        
        # Projekte extrahieren
        projects = self.projects_data.get("projekte", [])
        project_names = [project.get("name", "") for project in projects if project.get("name")]
        
        if not project_names:
            self.info_label.configure(
                text="Keine gültigen Projekte gefunden.",
                text_color="red"
            )
            return
        
        # Dropdown aktualisieren
        self.project_var.set(project_names[0])
        self.project_dropdown.configure(values=project_names)
        
        # Zeige Mitglieder des ersten Projekts
        self._on_project_selected(project_names[0])
    
    def _on_project_selected(self, project_name):
        """Wird aufgerufen, wenn ein Projekt ausgewählt wird"""
        # Bisherige Auswahl zurücksetzen
        self.selected_member = None
        self.select_button.configure(state="disabled")
        
        # Bisherige Mitgliederliste löschen
        for widget in self.members_list_frame.winfo_children():
            widget.destroy()
        
        # Finde das ausgewählte Projekt
        selected_project = None
        for project in self.projects_data.get("projekte", []):
            if project.get("name") == project_name:
                selected_project = project
                break
        
        if not selected_project:
            self.info_label.configure(
                text=f"Projekt '{project_name}' nicht gefunden.",
                text_color="red"
            )
            return
        
        # Mitglieder des Projekts extrahieren
        members = selected_project.get("teilnehmer", [])
        
        if not members:
            self.info_label.configure(
                text=f"Keine Mitglieder im Projekt '{project_name}' gefunden.",
                text_color="red"
            )
            return
        
        # Info-Label aktualisieren
        self.info_label.configure(
            text=f"{len(members)} Mitglieder gefunden. Bitte wählen Sie ein Mitglied aus.",
            text_color=("gray60", "gray70")
        )
        
        # Mitglieder anzeigen
        for i, member in enumerate(members):
            member_id = member.get("id", "")
            member_name = member.get("name", "")
            member_role = member.get("rolle", "")
            
            # Erstelle ein Frame für den Mitgliedseintrag
            member_frame = ctk.CTkFrame(self.members_list_frame)
            member_frame.pack(fill="x", padx=5, pady=2)
            
            # Hintergrundfarbe abwechseln
            if i % 2 == 0:
                member_frame.configure(fg_color=("gray90", "gray20"))
            
            # Speichere die Mitgliedsdaten im Frame
            member_frame.member_data = member
            
            # Erstelle Labels für Name und Rolle
            name_label = ctk.CTkLabel(
                member_frame,
                text=member_name,
                anchor="w",
                width=200
            )
            
            role_label = ctk.CTkLabel(
                member_frame,
                text=member_role,
                anchor="w",
                width=150
            )
            
            id_label = ctk.CTkLabel(
                member_frame,
                text=member_id,
                anchor="w",
                width=100
            )
            
            
            # Platziere die Labels
            name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
            role_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
            id_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
            
            member_frame.bind("<Button-1>", lambda e, f=member_frame: self._on_member_clicked(f))
            # Doppelklick zur direkten Auswahl
            member_frame.bind("<Double-Button-1>", lambda e, f=member_frame: self._on_member_double_click(f))
            name_label.bind("<Button-1>", lambda e, f=member_frame: self._on_member_clicked(f))
            name_label.bind("<Double-Button-1>", lambda e, f=member_frame: self._on_member_double_click(f))
            role_label.bind("<Button-1>", lambda e, f=member_frame: self._on_member_clicked(f))
            role_label.bind("<Double-Button-1>", lambda e, f=member_frame: self._on_member_double_click(f))
            id_label.bind("<Button-1>", lambda e, f=member_frame: self._on_member_clicked(f))
            id_label.bind("<Double-Button-1>", lambda e, f=member_frame: self._on_member_double_click(f))

    def _on_member_double_click(self, member_frame):
        """Wird aufgerufen, wenn auf ein Mitglied doppelgeklickt wird"""
        # Wähle das Mitglied aus
        self._on_member_clicked(member_frame)
        
        # Bestätige die Auswahl sofort, als würde der Auswählen-Button geklickt
        if self.selected_member:
            self._on_select_clicked()     
    
    def _on_member_clicked(self, member_frame):
        """Wird aufgerufen, wenn ein Mitglied angeklickt wird"""
        # Bisherige Auswahl zurücksetzen
        for widget in self.members_list_frame.winfo_children():
            if hasattr(widget, 'member_data'):
                i = list(self.members_list_frame.winfo_children()).index(widget)
                if i % 2 == 0:
                    widget.configure(fg_color=("gray90", "gray20"))
                else:
                    widget.configure(fg_color="transparent")
        
        # Diese Zeile als ausgewählt markieren
        member_frame.configure(fg_color=("light blue", "dark blue"))
        
        # Mitgliedsdaten speichern
        self.selected_member = member_frame.member_data
        
        # Auswählen-Button aktivieren
        self.select_button.configure(state="normal")
        
        # Info-Label aktualisieren
        member_name = self.selected_member.get("name", "")
        self.info_label.configure(
            text=f"Ausgewählt: {member_name}",
            text_color=("gray60", "gray70")
        )
    
    def _on_select_clicked(self):
        """Wird aufgerufen, wenn der Auswählen-Button geklickt wird"""
        if self.selected_member and self.on_select:
            self.on_select(self.selected_member)
        
        # Dialog schließen
        self.destroy()
    
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


class ImportModal(ctk.CTkToplevel):
    """
    Modal-Fenster zum Importieren von Daten mit Tabs für BW Tool Daten und Arbeitszeiterfassung
    """
    def __init__(self, master, callbacks: Optional[Dict[str, Callable]] = None):
        """
        Initialisiert das Modal-Fenster
        
        Args:
            master: Das übergeordnete Widget
            callbacks: Dictionary mit Callback-Funktionen
        """
        super().__init__(master)
        
        # Fenster zunächst ausblenden
        self.withdraw()
        
        # Fenster-Konfiguration
        self.title("Daten importieren")
        self.geometry("700x550")
        self.minsize(600, 450)  # Minimale Fenstergröße festlegen
        self.transient(master)  # Macht das Fenster zum Modal-Fenster
        self.grab_set()         # Blockiert Interaktion mit dem Hauptfenster
        
        # Attribute
        self.callbacks = callbacks or {}
        self.selected_file = None
        self.import_status = None
        
        # Attribute für die Arbeitszeiterfassung
        self.selected_member = None
        
        # UI erstellen
        self._create_widgets()
        self._setup_layout()
        
        # Fenster zentrieren und dann anzeigen
        self.center_window()
        
        # Nach einer kurzen Verzögerung anzeigen (warten bis UI gerendert ist)
        self.after(100, self.deiconify)
        
        # Stelle sicher, dass das Fenster tatsächlich angezeigt wird
        self.attributes('-topmost', True)
        self.after(200, lambda: self.attributes('-topmost', False))
    
    def _create_widgets(self):
        """Erstellt alle UI-Elemente für das Modal"""
        # Titel-Frame
        self.title_frame = ctk.CTkFrame(self)
        self.title_label = ctk.CTkLabel(
            self.title_frame, 
            text="Daten importieren",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        
        # Status-Label
        self.status_frame = ctk.CTkFrame(self)
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray70")
        )
        
        # Tab-View für die verschiedenen Import-Optionen
        self.tab_view = ctk.CTkTabview(self)
        
        # Tabs erstellen
        self.bw_tool_tab = self.tab_view.add("BW Tool Daten")
        self.ar_calendar_tab = self.tab_view.add("Arbeitszeiterfassung")
        
        # Inhalte für BW Tool Daten-Tab
        self._create_bw_tool_tab()
        
        # Inhalte für Arbeitszeiterfassung-Tab
        self._create_timesheet_tab()
        
        # Button-Frame (unten)
        self.button_frame = ctk.CTkFrame(self)
        self.close_button = ctk.CTkButton(
            self.button_frame, 
            text="Schließen",
            command=self.destroy,
            width=120
        )
    
    def _create_bw_tool_tab(self):
        """Erstellt die Inhalte für den BW Tool Daten-Tab"""
        # Container-Frame für den Tab
        container = ctk.CTkFrame(self.bw_tool_tab)
        
        # Info-Text
        info_text = (
            "Importieren Sie BW Tool Daten aus einer Excel-Datei (.xls, .xlsx).\n\n"
            "Die Daten werden aus der Excel-Datei extrahiert und in folgendem Format gespeichert:\n"
            "MITARBEITER_ID;DATUM;STUNDEN\n\n"
            "Es werden nur Einträge berücksichtigt, die in Spalte 11 'FCAST', 'Forecast' oder 'FORECAST' enthalten."
        )
        
        info_label = ctk.CTkLabel(
            container,
            text=info_text,
            justify="left",
            anchor="w",
            wraplength=600
        )
        
        # Drag & Drop Frame
        self.drop_frame = ctk.CTkFrame(
            container,
            fg_color=("gray90", "gray20"),
            corner_radius=10,
            border_width=2,
            border_color=("gray70", "gray40")
        )
        
        # Text und Icon für den Drop-Bereich
        self.drop_text_label = ctk.CTkLabel(
            self.drop_frame,
            text="Datei hier hineinziehen oder klicken zum Auswählen",
            font=ctk.CTkFont(size=14)
        )
        self.drop_text_label.pack(expand=True, pady=(20, 10))
        
        self.drop_info_label = ctk.CTkLabel(
            self.drop_frame,
            text="Akzeptierte Dateitypen: .xls, .xlsx",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray70")
        )
        self.drop_info_label.pack(expand=True, pady=(0, 20))
        
        # Erhöhe den Frame so, dass er gut als Drop-Ziel funktioniert
        self.drop_frame.configure(height=150)
        
        # Ereignisbehandlung für Mausklicks im Drop-Frame
        self.drop_frame.bind("<Button-1>", self._browse_bw_tool_file)
        self.drop_text_label.bind("<Button-1>", self._browse_bw_tool_file)
        self.drop_info_label.bind("<Button-1>", self._browse_bw_tool_file)
        
        # Dateiauswahl-Frame (für Anzeige des ausgewählten Dateinamens)
        file_frame = ctk.CTkFrame(container)
        
        # Label für ausgewählte Datei
        file_label = ctk.CTkLabel(
            file_frame,
            text="Ausgewählte Datei:",
            font=ctk.CTkFont(weight="bold"),
            width=120
        )
        
        # Anzeige des Dateinamens
        self.file_name_label = ctk.CTkLabel(
            file_frame,
            text="Keine Datei ausgewählt",
            anchor="w"
        )
        
        # Optionen-Frame
        options_frame = ctk.CTkFrame(container)
        
        # Optionen für den Import
        self.bw_tool_overwrite_var = ctk.BooleanVar(value=True)
        self.bw_tool_overwrite_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Bestehende Einträge überschreiben (gleiche ID und Datum)",
            variable=self.bw_tool_overwrite_var
        )
        
        self.bw_tool_backup_var = ctk.BooleanVar(value=True)
        self.bw_tool_backup_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Backup erstellen",
            variable=self.bw_tool_backup_var
        )
        
        # Import-Button
        self.bw_tool_import_button = ctk.CTkButton(
            container,
            text="BW Tool Daten importieren",
            command=self._import_bw_tool_data,
            font=ctk.CTkFont(weight="bold"),
            height=40,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3"),
            state="disabled"  # Initial deaktiviert
        )
        
        # Layout für den BW Tool Daten-Tab
        container.pack(fill="both", expand=True, padx=20, pady=20)
        info_label.pack(anchor="w", pady=(0, 15))
        
        # Drag & Drop Frame
        self.drop_frame.pack(fill="x", pady=(0, 15))
        
        # Dateiauswahl-Frame
        file_frame.pack(fill="x", pady=(0, 15))
        file_label.pack(side="left")
        self.file_name_label.pack(side="left", padx=10, fill="x", expand=True)
        
        # Optionen-Frame
        options_frame.pack(fill="x", pady=(0, 20))
        self.bw_tool_overwrite_checkbox.pack(anchor="w", pady=5)
        self.bw_tool_backup_checkbox.pack(anchor="w", pady=5)
        
        # Import-Button
        self.bw_tool_import_button.pack(fill="x", pady=(10, 0))
    
    def _create_timesheet_tab(self):
        """Erstellt die Inhalte für den Arbeitszeiterfassung-Tab"""
        # Container-Frame für den Tab
        container = ctk.CTkFrame(self.ar_calendar_tab)
        
        # Info-Text
        info_text = (
            "Importieren Sie Arbeitszeiterfassungsdaten direkt aus der Zeiterfassung.\n\n"
            "Wählen Sie einen Mitarbeiter aus und fügen Sie die Daten im folgenden Format ein:\n"
            "Date StartTime EndTime JobCode Break TotalHours Description ... \n"
            "Die Spalten Date und TotalHours werden extrahiert und in die Kapazitätsdaten übernommen."
        )
        
        info_label = ctk.CTkLabel(
            container,
            text=info_text,
            justify="left",
            anchor="w",
            wraplength=600
        )
        
        # Mitarbeiter-Auswahl-Frame
        member_frame = ctk.CTkFrame(container)
        
        # Label für Mitarbeiterauswahl
        member_label = ctk.CTkLabel(
            member_frame,
            text="Mitarbeiter:",
            font=ctk.CTkFont(weight="bold"),
            width=120
        )
        
        # Anzeige des ausgewählten Mitarbeiters
        self.selected_member_label = ctk.CTkLabel(
            member_frame,
            text="Kein Mitarbeiter ausgewählt",
            anchor="w"
        )
        
        # Button für Mitarbeiterauswahl
        self.select_member_button = ctk.CTkButton(
            member_frame,
            text="Mitarbeiter auswählen",
            command=self._open_member_selection,
            width=170
        )
        
        # Textbox für die Eingabe der Arbeitszeitdaten
        time_data_frame = ctk.CTkFrame(container)
        time_data_label = ctk.CTkLabel(
            time_data_frame,
            text="Arbeitszeitdaten einfügen:",
            font=ctk.CTkFont(weight="bold"),
            anchor="w"
        )
        
        self.time_data_textbox = ctk.CTkTextbox(
            time_data_frame,
            height=200,
            font=ctk.CTkFont(family="Courier", size=11)
        )
        
        # Optionen-Frame
        options_frame = ctk.CTkFrame(container)
        
        # Optionen für den Import
        self.ar_overwrite_var = ctk.BooleanVar(value=True)
        self.ar_overwrite_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Bestehende Einträge überschreiben (gleiche ID und Datum)",
            variable=self.ar_overwrite_var
        )
        
        self.ar_backup_var = ctk.BooleanVar(value=True)
        self.ar_backup_checkbox = ctk.CTkCheckBox(
            options_frame,
            text="Backup erstellen",
            variable=self.ar_backup_var
        )
        
        # Import-Button
        self.ar_import_button = ctk.CTkButton(
            container,
            text="Arbeitszeitdaten importieren",
            command=self._import_ar_calendar_data,
            font=ctk.CTkFont(weight="bold"),
            height=40,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3"),
            state="disabled"  # Initial deaktiviert
        )
        
        # Layout für den Arbeitszeit-Tab
        container.pack(fill="both", expand=True, padx=20, pady=20)
        info_label.pack(anchor="w", pady=(0, 15))
        
        # Mitarbeiter-Auswahl-Frame
        member_frame.pack(fill="x", pady=(0, 15))
        member_label.pack(side="left")
        self.selected_member_label.pack(side="left", padx=10, fill="x", expand=True)
        self.select_member_button.pack(side="right")
        
        # Textbox für die Arbeitszeitdaten
        time_data_frame.pack(fill="both", expand=True, pady=(0, 15))
        time_data_label.pack(anchor="w", pady=(0, 5))
        self.time_data_textbox.pack(fill="both", expand=True)
        
        # Optionen-Frame
        options_frame.pack(fill="x", pady=(0, 20))
        self.ar_overwrite_checkbox.pack(anchor="w", pady=5)
        self.ar_backup_checkbox.pack(anchor="w", pady=5)
        
        # Import-Button
        self.ar_import_button.pack(fill="x", pady=(10, 0))
    
    def _setup_layout(self):
        """Richtet das Layout der UI-Elemente ein"""
        # Titel-Frame
        self.title_frame.pack(fill="x", padx=20, pady=(20, 10))
        self.title_label.pack(pady=10)
        
        # Status-Frame
        self.status_frame.pack(fill="x", padx=20, pady=(0, 10))
        self.status_label.pack(pady=5, fill="x")
        
        # Tab-View (nimmt den meisten Platz ein)
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Button-Frame (unten)
        self.button_frame.pack(fill="x", padx=20, pady=(5, 20))
        self.close_button.pack(side="right", padx=10)
    
    def _browse_bw_tool_file(self, event=None):
        """Öffnet einen Datei-Browser-Dialog für BW Tool Dateien"""
        filetypes = [("Excel-Dateien", "*.xls *.xlsx"), ("Alle Dateien", "*.*")]
        filename = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file = filename
            self.file_name_label.configure(text=os.path.basename(filename))
            self.bw_tool_import_button.configure(state="normal")
            
            # Status zurücksetzen
            self._show_status("", False)
    
    def _open_member_selection(self):
        """Öffnet den Dialog zur Auswahl eines Projektmitglieds"""
        try:
            # Lade die Projektdaten
            projects_data = self._load_project_data()
            
            if not projects_data:
                self._show_status("Keine Projektdaten gefunden. Bitte importieren Sie zuerst Projektdaten.", error=True)
                return
            
            # Zeige den Dialog
            dialog = ProjectMemberSelectionDialog(
                self,
                projects_data,
                on_select=self._on_member_selected
            )
            
            # Warte, bis der Dialog geschlossen wird
            self.wait_window(dialog)
        
        except Exception as e:
            self._show_status(f"Fehler beim Öffnen des Mitarbeiterauswahl-Dialogs: {e}", error=True)
    
    def _on_member_selected(self, member_data):
        """
        Wird aufgerufen, wenn ein Projektmitglied ausgewählt wurde
        
        Args:
            member_data: Die Daten des ausgewählten Mitglieds
        """
        self.selected_member = member_data
        
        # Zeige den Namen des ausgewählten Mitglieds an
        member_name = member_data.get("name", "")
        member_id = member_data.get("id", "")
        self.selected_member_label.configure(text=f"{member_name} (ID: {member_id})")
        
        # Prüfe, ob auch Daten in der Textbox sind
        self._update_ar_import_button()
    
    def _update_ar_import_button(self):
        """Aktualisiert den Status des Import-Buttons für die Arbeitszeiterfassung"""
        if self.selected_member and self.time_data_textbox.get("1.0", "end-1c").strip():
            self.ar_import_button.configure(state="normal")
        else:
            self.ar_import_button.configure(state="disabled")
            
        # Binde das Event an den TextBox-Inhalt
        self.time_data_textbox.bind("<KeyRelease>", lambda e: self._update_ar_import_button())
    
    def _import_bw_tool_data(self):
        """Importiert BW Tool Daten aus einer Excel-Datei"""
        # Prüfen, ob eine Datei ausgewählt wurde
        if not self.selected_file:
            self._show_status("Bitte wählen Sie eine Excel-Datei aus.", error=True)
            return
        
        # Prüfen, ob die Datei existiert
        if not os.path.exists(self.selected_file):
            self._show_status(f"Die Datei '{self.selected_file}' existiert nicht.", error=True)
            return
        
        try:
            # Excel-Datei einlesen
            self._show_status("Excel-Datei wird eingelesen...", error=False)
            
            # Ziel-Dateipfad
            target_path = os.path.join("data", "kapa_data.csv")
            
            # Backup erstellen, falls gewünscht
            if self.bw_tool_backup_var.get() and os.path.exists(target_path):
                backup_path = f"{target_path}.bak"
                shutil.copy2(target_path, backup_path)
            
            # Verarbeite die Excel-Datei
            processed_data = self._process_mime_file(self.selected_file)
            
            # Wenn keine Daten verarbeitet wurden
            if not processed_data:
                self._show_status("Keine gültigen Daten in der Datei gefunden.", error=True)
                return
            
            # Einträge überschreiben, falls vorhanden
            if self.bw_tool_overwrite_var.get() and os.path.exists(target_path):
                final_data = self._merge_with_existing_data(processed_data, target_path)
            else:
                # Nur neue Daten hinzufügen
                final_data = processed_data
                if os.path.exists(target_path):
                    with open(target_path, 'r', newline='', encoding='utf-8') as existing_file:
                        reader = csv.reader(existing_file, delimiter=';')
                        existing_data = list(reader)
                        # Nur Zeilen hinzufügen, die noch nicht existieren
                        id_date_pairs = {(row[0], row[1]) for row in existing_data if len(row) >= 2}
                        final_data = [row for row in existing_data] + [
                            row for row in processed_data if (row[0], row[1]) not in id_date_pairs
                        ]
            
            # Daten speichern
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, 'w', newline='', encoding='utf-8') as output_file:
                writer = csv.writer(output_file, delimiter=';')
                writer.writerows(final_data)
            
            # Erfolgsmeldung anzeigen
            self._show_status(f"Import erfolgreich: {len(processed_data)} Einträge verarbeitet", error=False)
        
        except Exception as e:
            # Fehlermeldung anzeigen
            self._show_status(f"Fehler beim Import: {str(e)}", error=True)
            import traceback
            traceback.print_exc()
    
    def _import_ar_calendar_data(self):
        """Importiert Daten aus der Arbeitszeiterfassung"""
        # Prüfen, ob ein Mitarbeiter ausgewählt wurde
        if not self.selected_member:
            self._show_status("Bitte wählen Sie einen Mitarbeiter aus.", error=True)
            return
        
        # Prüfen, ob Daten eingegeben wurden
        time_data = self.time_data_textbox.get("1.0", "end-1c").strip()
        if not time_data:
            self._show_status("Bitte fügen Sie Arbeitszeitdaten ein.", error=True)
            return
        
        try:
            # Mitarbeiter-ID extrahieren
            member_id = self.selected_member.get("id", "")
            
            # Anzeigen, dass die Verarbeitung beginnt
            self._show_status("Verarbeite Arbeitszeitdaten...", error=False)
            
            # Ziel-Dateipfad
            target_path = os.path.join("data", "kapa_data.csv")
            
            # Backup erstellen, falls gewünscht
            if self.ar_backup_var.get() and os.path.exists(target_path):
                backup_path = f"{target_path}.bak"
                shutil.copy2(target_path, backup_path)
            
            # Verarbeite die Arbeitszeitdaten
            processed_data = self._process_timesheet_data(time_data, member_id)
            
            # Wenn keine Daten verarbeitet wurden
            if not processed_data:
                self._show_status("Keine gültigen Daten gefunden oder Format nicht erkannt.", error=True)
                return
            
            # Einträge überschreiben, falls vorhanden
            if self.ar_overwrite_var.get() and os.path.exists(target_path):
                final_data = self._merge_with_existing_data(processed_data, target_path)
            else:
                # Nur neue Daten hinzufügen
                final_data = processed_data
                if os.path.exists(target_path):
                    with open(target_path, 'r', newline='', encoding='utf-8') as existing_file:
                        reader = csv.reader(existing_file, delimiter=';')
                        existing_data = list(reader)
                        # Nur Zeilen hinzufügen, die noch nicht existieren
                        id_date_pairs = {(row[0], row[1]) for row in existing_data if len(row) >= 2}
                        final_data = [row for row in existing_data] + [
                            row for row in processed_data if (row[0], row[1]) not in id_date_pairs
                        ]
            
            # Daten speichern
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, 'w', newline='', encoding='utf-8') as output_file:
                writer = csv.writer(output_file, delimiter=';')
                writer.writerows(final_data)
            
            # Erfolgsmeldung anzeigen
            self._show_status(f"Import erfolgreich: {len(processed_data)} Einträge verarbeitet", error=False)
            
            # Textbox leeren
            self.time_data_textbox.delete("1.0", "end")
            self.ar_import_button.configure(state="disabled")
            
        except Exception as e:
            # Fehlermeldung anzeigen
            self._show_status(f"Fehler beim Import: {str(e)}", error=True)
            import traceback
            traceback.print_exc()
    
    def _process_mime_file(self, file_path):
        """
        Verarbeitet eine MIME-formatierte Datei
        
        Args:
            file_path: Pfad zur MIME/HTML-Datei
            
        Returns:
            List[List[str]]: Verarbeitete Daten im Format [[ID, DATUM, STUNDEN, KAPAZITÄT], ...]
        """
        try:
            # Dateiinhalt mit geeigneter Kodierung lesen
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                try:
                    with open(file_path, 'r', encoding='latin1') as f:
                        content = f.read()
                except:
                    with open(file_path, 'r', encoding='cp1252') as f:
                        content = f.read()
            
            import re
            import datetime
            
            # Liste für verarbeitete Daten
            processed_data = []
            
            # Header für CSV-Datei
            header = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]
            processed_data.append(header)
            
            # Alle Tabellenzeilen extrahieren
            rows = re.findall(r'<tr.*?>(.*?)</tr>', content, re.DOTALL)
            
            # Zeilen verarbeiten
            for row in rows:
                # Alle Zellen in dieser Zeile extrahieren
                cells = re.findall(r'<td.*?>(.*?)</td>', row, re.DOTALL)
                
                # Zeilen mit unzureichenden Zellen überspringen
                if len(cells) < 13:
                    continue
                
                try:
                    # Prüfen, ob dies ein Forecast-Eintrag ist (Spalte 11, Index 10)
                    forecast_cell = cells[10] if len(cells) > 10 else ""
                    forecast_value = re.sub(r'<.*?>', '', forecast_cell).upper().strip()
                    
                    if "FCAST" not in forecast_value and "FORECAST" not in forecast_value:
                        continue
                    
                    # Name (ID) extrahieren
                    sixth_value = re.sub(r'<.*?>', '', cells[5]).replace("&#32;", " ").strip()
                    if sixth_value == "Fremdleistung OPS" and len(cells) > 6:
                        mitarbeiter_id = re.sub(r'<.*?>', '', cells[6]).replace("&#32;", " ").strip()
                    else:
                        mitarbeiter_id = sixth_value
                    
                    # Datum extrahieren
                    date_cell = cells[7]
                    date_value = re.sub(r'<.*?>', '', date_cell).strip()
                    try:
                        # Versuche, einen Excel-Datumswert zu konvertieren
                        excel_date = int(float(date_value.replace(",", ".")))
                        datetime_date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=excel_date)
                        datum = datetime_date.strftime('%d.%m.%Y')
                    except ValueError:
                        # Wenn es kein numerischer Wert ist, den Wert direkt verwenden
                        datum = date_value
                    
                    # Stunden extrahieren
                    hours_cell = cells[12]
                    hours_value = re.sub(r'<.*?>', '', hours_cell).strip()
                    try:
                        stunden_float = float(hours_value.replace(",", "."))
                        stunden = str(stunden_float)
                        
                        # Kapazität berechnen: Stunden >= 8 entsprechen einer Kapazität von 1
                        if stunden_float >= 8:
                            kapazitaet = "1.0"
                        else:
                            # Anteilige Kapazität berechnen (Stunden/8)
                            kapazitaet = str(round(stunden_float / 8, 2))
                        
                    except ValueError:
                        stunden = "0.0"  # Fallback, wenn keine gültige Zahl
                        kapazitaet = "0.0"
                    
                    # Daten zur Liste hinzufügen [ID, DATUM, STUNDEN, KAPAZITÄT]
                    processed_data.append([mitarbeiter_id, datum, stunden, kapazitaet])
                    
                except Exception:
                    continue
            
            return processed_data
        
        except Exception as e:
            self._show_status(f"MIME-Verarbeitung fehlgeschlagen: {e}", error=True)
            return []

    def _process_timesheet_data(self, data_text, member_id):
        """
        Verarbeitet die eingegebenen Arbeitszeitdaten
        
        Args:
            data_text: Der eingegebene Text mit den Arbeitszeitdaten
            member_id: Die ID des ausgewählten Mitarbeiters
            
        Returns:
            List[List[str]]: Verarbeitete Daten im Format [[ID, DATUM, STUNDEN, KAPAZITÄT], ...]
        """
        # Liste für die verarbeiteten Daten
        processed_data = []

        # Header für die CSV-Datei, falls noch nicht vorhanden
        csv_header = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]
        processed_data.append(csv_header)
        
        # Text in Zeilen aufteilen
        lines = data_text.strip().split("\n")
        
        # Mindestlänge für eine gültige Zeile
        min_columns = 6  # Date, StartTime, EndTime, JobCode, Break, TotalHours
        
        # Versuche zuerst, Kopfzeilen zu identifizieren
        header_line = None
        data_lines = []
        date_column = None
        hours_column = None
        
        for i, line in enumerate(lines):
            columns = line.split()
            if len(columns) >= min_columns:
                if header_line is None and ("Date" in columns or "TotalHours" in columns):
                    header_line = i
                    # Finde die Spalten-Indizes für Date und TotalHours
                    for j, col in enumerate(columns):
                        if col.lower() == "date":
                            date_column = j
                        elif col.lower() == "totalhours":
                            hours_column = j
                else:
                    data_lines.append(i)
        
        # Wenn kein expliziter Header gefunden wurde, prüfe ob wir ein bekanntes Format erkennen können
        if header_line is None:
            # Versuche, ein Datum im Format YYYY-MM-DD in der ersten Spalte zu erkennen
            first_line = lines[0].split() if lines else []
            if first_line and re.match(r'\d{4}-\d{2}-\d{2}', first_line[0]):
                date_column = 0
                # Suche nach Spalten mit Zahlen, die TotalHours sein könnten
                for j, col in enumerate(first_line):
                    if j > 0 and re.match(r'\d+(\.\d+)?', col):
                        # Wähle die 6. Spalte (Index 5) als TotalHours, wenn sie existiert
                        if j == 5:
                            hours_column = j
                            break
                
                # Fallback: Nimm die erste Zahl nach dem Datum
                if hours_column is None and len(first_line) > 1:
                    for j in range(1, min(len(first_line), 10)):
                        if first_line[j].replace('.', '').isdigit():
                            hours_column = j
                            break
                
                data_lines = list(range(len(lines)))
        
        # Wenn wir immer noch keine Spalten identifizieren konnten, versuche es mit Standardpositionen
        if date_column is None or hours_column is None:
            date_column = 0  # Erste Spalte ist meist Datum
            hours_column = 5  # Sechste Spalte ist oft TotalHours
            data_lines = list(range(len(lines)))
        
        self._show_status(f"Verarbeite {len(data_lines)} Zeilen, Date-Spalte: {date_column}, Hours-Spalte: {hours_column}", error=False)
        
        # Verarbeite die Datenzeilen
        for line_idx in data_lines:
            line = lines[line_idx]
            columns = line.split()
            
            # Überspringe Header oder zu kurze Zeilen
            if line_idx == header_line or len(columns) <= max(date_column, hours_column):
                continue
            
            try:
                # Datum extrahieren
                date_str = columns[date_column]
                # Versuche, das Datum zu normalisieren (YYYY-MM-DD -> DD.MM.YYYY)
                try:
                    if re.match(r'\d{4}-\d{2}-\d{2}', date_str):
                        year, month, day = date_str.split('-')
                        date_str = f"{day}.{month}.{year}"
                except:
                    pass
                
                # Stunden extrahieren
                hours_str = columns[hours_column]
                try:
                    hours = float(hours_str)
                    
                    # Kapazität berechnen: Stunden >= 8 entsprechen Kapazität 1
                    if hours >= 8:
                        capacity = 1.0
                    else:
                        capacity = round(hours / 8, 2)
                    
                    # Daten zur Liste hinzufügen
                    processed_data.append([member_id, date_str, str(hours), str(capacity)])
                except ValueError:
                    # Keine gültige Stundenzahl
                    continue
            except Exception as e:
                # Fehler beim Verarbeiten der Zeile
                print(f"Fehler bei Zeile {line_idx + 1}: {e}")
                continue
        
        return processed_data
    
    def _merge_with_existing_data(self, new_data, existing_file_path):
        """
        Führt neue Daten mit vorhandenen Daten zusammen, wobei vorhandene Einträge überschrieben werden
        
        Args:
            new_data: Neue Daten [[ID, DATUM, STUNDEN, KAPAZITÄT], ...]
            existing_file_path: Pfad zur bestehenden CSV-Datei
                
        Returns:
            List[List[str]]: Zusammengeführte Daten
        """
        if not os.path.exists(existing_file_path):
            # Wenn die Datei nicht existiert, gib einfach die neuen Daten zurück
            return new_data
        
        try:
            # Vorhandene Daten einlesen
            with open(existing_file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file, delimiter=';')
                existing_data = list(reader)
            
            if not existing_data:
                # Leere Datei, nur Header erstellen
                return new_data
            
            # Prüfe ob erste Zeile ein Header ist
            first_row = existing_data[0]
            has_header = False
            
            if first_row and len(first_row) >= 3 and (first_row[0] == "ID" or first_row[0] == "id"):
                has_header = True
                header = first_row
                data_rows = existing_data[1:]
            else:
                # Kein Header gefunden, die neuen Daten enthalten bereits einen Header
                data_rows = existing_data
            
            # Dictionary für schnellere Suche erstellen (ID+DATUM -> Zeilenindex)
            id_date_indices = {(row[0], row[1]): i for i, row in enumerate(data_rows) if len(row) >= 2}
            
            # Neue Daten (ohne Header) durchgehen
            merged_data = data_rows.copy()
            for row in new_data[1:]:  # Überspringe den Header in den neuen Daten
                if len(row) >= 2:
                    key = (row[0], row[1])
                    if key in id_date_indices:
                        # Vorhandenen Eintrag überschreiben
                        merged_data[id_date_indices[key]] = row
                    else:
                        # Neuen Eintrag hinzufügen
                        merged_data.append(row)
            
            # Header wieder hinzufügen
            return [header] + merged_data if has_header else merged_data
        
        except Exception as e:
            raise Exception(f"Fehler beim Zusammenführen der Daten: {str(e)}")
    
    def _show_status(self, message, error=False):
        """
        Zeigt eine Statusmeldung an
        
        Args:
            message: Die anzuzeigende Nachricht
            error: True für eine Fehlermeldung, False für eine Erfolgsmeldung
        """
        if not message:
            self.status_label.configure(text="")
            return
        
        # Nachricht anzeigen
        self.status_label.configure(
            text=message,
            text_color="red" if error else "green"
        )
    
    def _load_project_data(self):
        """
        Lädt die Projektdaten aus der project.json-Datei
        
        Returns:
            dict: Die geladenen Projektdaten oder None bei einem Fehler
        """
        try:
            # Pfad zur project.json
            json_path = os.path.join("data", "project.json")
            
            # Prüfe, ob die Datei existiert
            if not os.path.exists(json_path):
                print(f"Warnung: Datei {json_path} nicht gefunden.")
                return None
            
            # Lade die Daten
            with open(json_path, 'r', encoding='utf-8') as json_file:
                return json.load(json_file)
        
        except Exception as e:
            print(f"Fehler beim Laden der Projektdaten: {e}")
            return None
    
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