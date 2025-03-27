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
import pandas as pd
import re


class DragDropFrame(ctk.CTkFrame):
    """
    Ein Frame, das Drag & Drop von Dateien ermöglicht
    """
    def __init__(self, master, callback, accepted_files=None, **kwargs):
        """
        Initialisiert den Drag & Drop Frame
        
        Args:
            master: Das übergeordnete Widget
            callback: Funktion, die aufgerufen wird, wenn eine Datei fallen gelassen wird
            accepted_files: Liste der akzeptierten Dateierweiterungen (z.B. ['.xls', '.xlsx'])
        """
        super().__init__(master, **kwargs)
        
        self.callback = callback
        self.accepted_files = accepted_files or ['.xls', '.xlsx']
        
        # Aussehen des Frames
        self.configure(
            fg_color=("gray90", "gray20"),
            corner_radius=10,
            border_width=2,
            border_color=("gray70", "gray40")
        )
        
        # Text und Icon für den Drop-Bereich
        self.text_label = ctk.CTkLabel(
            self,
            text="Datei hier hineinziehen oder klicken zum Auswählen",
            font=ctk.CTkFont(size=14)
        )
        self.text_label.pack(expand=True, pady=(20, 10))
        
        self.info_label = ctk.CTkLabel(
            self,
            text=f"Akzeptierte Dateitypen: {', '.join(self.accepted_files)}",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray70")
        )
        self.info_label.pack(expand=True, pady=(0, 20))
        
        # Erhöhe den Frame so, dass er gut als Drop-Ziel funktioniert
        self.configure(height=150)
        
        # Ereignisbehandlung für Mausklicks
        self.bind("<Button-1>", self._on_click)
        self.text_label.bind("<Button-1>", self._on_click)
        self.info_label.bind("<Button-1>", self._on_click)
        
        # Richte Drag & Drop ein für Windows (TkDND)
        self._setup_tkdnd()
    
    def _setup_tkdnd(self):
        """Richtet TkDND für Drag & Drop-Funktionalität ein, falls verfügbar"""
        try:
            self.winfo_toplevel().tk.call('package', 'require', 'tkdnd')
            self.winfo_toplevel().tk.call('tkdnd::drop_target', 'register', self._get_widget_path(), self.accepted_files)
            self.bind('<<Drop>>', self._on_drop)
            
            # Visuelle Rückmeldung beim Überfahren
            self.bind('<<DropEnter>>', lambda e: self.configure(border_color=("blue", "light blue")))
            self.bind('<<DropLeave>>', lambda e: self.configure(border_color=("gray70", "gray40")))
        except tk.TclError:
            # TkDND ist nicht verfügbar, füge eine Info hinzu
            self.info_label.configure(text="Drag & Drop nicht verfügbar, bitte klicken zum Auswählen")
    
    def _get_widget_path(self):
        """Gibt den TK-Widget-Pfad für TkDND zurück"""
        return self.winfo_toplevel().register(self)
    
    def _on_click(self, event):
        """Wird aufgerufen, wenn auf den Frame geklickt wird"""
        filetypes = [("Excel-Dateien", "*.xls *.xlsx"), ("Alle Dateien", "*.*")]
        filename = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=filetypes
        )
        
        if filename:
            self.callback(filename)
    
    def _on_drop(self, event):
        """Wird aufgerufen, wenn eine Datei auf den Frame gezogen wird"""
        self.configure(border_color=("gray70", "gray40"))  # Zurücksetzen der Rahmenfarbe
        
        # Prüfe, ob die Datei einen der akzeptierten Dateitypen hat
        filename = event.data
        _, file_extension = os.path.splitext(filename)
        
        if file_extension.lower() in self.accepted_files:
            self.callback(filename)
        else:
            messagebox.showerror(
                "Ungültiger Dateityp",
                f"Die Datei {os.path.basename(filename)} hat einen ungültigen Dateityp. "
                f"Akzeptierte Typen: {', '.join(self.accepted_files)}",
                parent=self.winfo_toplevel()
            )


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
        
        # Neue Attribute für die Arbeitszeiterfassung
        self.ar_selected_file = None
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
        self.drop_frame = DragDropFrame(
            container,
            callback=self._on_file_selected,
            accepted_files=['.xls', '.xlsx']
        )
        
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
            "Importieren Sie Arbeitszeiterfassungsdaten aus einer Excel-Datei (.xlsx, .xls).\n\n"
            "Die Datei sollte Spalten für 'Date' und 'TotalHours' enthalten. "
            "Da die Datei keine Mitarbeiter-ID enthält, müssen Sie einen Mitarbeiter auswählen, "
            "dem diese Arbeitszeiten zugeordnet werden sollen."
        )
        
        info_label = ctk.CTkLabel(
            container,
            text=info_text,
            justify="left",
            anchor="w",
            wraplength=600
        )
        
        # Drag & Drop Frame
        self.ar_drop_frame = DragDropFrame(
            container,
            callback=self._on_ar_file_selected,
            accepted_files=['.xlsx', '.xls']
        )
        
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
        self.ar_file_name_label = ctk.CTkLabel(
            file_frame,
            text="Keine Datei ausgewählt",
            anchor="w"
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
        
        # Drag & Drop Frame
        self.ar_drop_frame.pack(fill="x", pady=(0, 15))
        
        # Dateiauswahl-Frame
        file_frame.pack(fill="x", pady=(0, 15))
        file_label.pack(side="left")
        self.ar_file_name_label.pack(side="left", padx=10, fill="x", expand=True)
        
        # Mitarbeiter-Auswahl-Frame
        member_frame.pack(fill="x", pady=(0, 15))
        member_label.pack(side="left")
        self.selected_member_label.pack(side="left", padx=10, fill="x", expand=True)
        self.select_member_button.pack(side="right")
        
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
    
    def _on_file_selected(self, filename):
        """
        Wird aufgerufen, wenn eine Datei ausgewählt oder fallen gelassen wurde
        
        Args:
            filename: Pfad zur ausgewählten Datei
        """
        self.selected_file = filename
        self.file_name_label.configure(text=os.path.basename(filename))
        self.bw_tool_import_button.configure(state="normal")
        
        # Status zurücksetzen
        self._show_status("", False)
    
    def _on_ar_file_selected(self, filename):
        """
        Wird aufgerufen, wenn eine Arbeitszeiterfassungsdatei ausgewählt wurde
        
        Args:
            filename: Pfad zur ausgewählten Datei
        """
        self.ar_selected_file = filename
        self.ar_file_name_label.configure(text=os.path.basename(filename))
        
        # Prüfe, ob auch ein Mitarbeiter ausgewählt wurde
        self._update_ar_import_button()
        
        # Status zurücksetzen
        self._show_status("", False)
    
    def _browse_file(self, file_type, entry_widget):
        """
        Öffnet einen Datei-Browser-Dialog
        
        Args:
            file_type: Der Dateityp (csv oder json)
            entry_widget: Das Eingabefeld, in das der Dateipfad eingetragen werden soll
        """
        if file_type.lower() == "csv":
            filetypes = [("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")]
            title = "CSV-Datei auswählen"
        elif file_type.lower() == "json":
            filetypes = [("JSON-Dateien", "*.json"), ("Alle Dateien", "*.*")]
            title = "JSON-Datei auswählen"
        else:
            filetypes = [("Alle Dateien", "*.*")]
            title = "Datei auswählen"
        
        # Datei-Dialog öffnen
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes
        )
        
        # Wenn eine Datei ausgewählt wurde, den Pfad in das Eingabefeld eintragen
        if filename:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, filename)
            
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
        
        # Prüfe, ob auch eine Datei ausgewählt wurde
        self._update_ar_import_button()
    
    def _update_ar_import_button(self):
        """Aktualisiert den Status des Import-Buttons für die Arbeitszeiterfassung"""
        if self.ar_selected_file and self.selected_member:
            self.ar_import_button.configure(state="normal")
        else:
            self.ar_import_button.configure(state="disabled")
    
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
                        id_date_pairs = {(row[0], row[1]) for row in existing_data}
                        final_data = [row for row in processed_data if (row[0], row[1]) not in id_date_pairs]
                        final_data = existing_data + final_data
            
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
        """Importiert Daten aus einer Arbeitszeiterfassungsdatei"""
        # Prüfen, ob eine Datei und ein Mitarbeiter ausgewählt wurden
        if not self.ar_selected_file:
            self._show_status("Bitte wählen Sie eine Excel-Datei aus.", error=True)
            return
        
        if not self.selected_member:
            self._show_status("Bitte wählen Sie einen Mitarbeiter aus.", error=True)
            return
        
        # Prüfen, ob die Datei existiert
        if not os.path.exists(self.ar_selected_file):
            self._show_status(f"Die Datei '{self.ar_selected_file}' existiert nicht.", error=True)
            return
        
        try:
            # Mitarbeiter-ID extrahieren
            member_id = self.selected_member.get("id", "")
            # Excel-Datei verarbeiten
            self._show_status("Arbeitszeitdaten werden verarbeitet...", error=False)
            
            # Ziel-Dateipfad
            target_path = os.path.join("data", "kapa_data.csv")
            
            # Backup erstellen, falls gewünscht
            if self.ar_backup_var.get() and os.path.exists(target_path):
                backup_path = f"{target_path}.bak"
                shutil.copy2(target_path, backup_path)
            
            # Verarbeite die Excel-Datei
            processed_data = self._process_ar_calendar_file(self.ar_selected_file, member_id)
            print(processed_data)
            # Wenn keine Daten verarbeitet wurden
            if not processed_data:
                self._show_status("Keine gültigen Daten in der Datei gefunden.", error=True)
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
                        id_date_pairs = {(row[0], row[1]) for row in existing_data}
                        final_data = [row for row in processed_data if (row[0], row[1]) not in id_date_pairs]
                        final_data = existing_data + final_data
            
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
            # Wenn die Datei nicht existiert, fügen wir einen Header hinzu
            header = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]
            return [header] + new_data
        
        try:
            # Vorhandene Daten einlesen
            with open(existing_file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file, delimiter=';')
                existing_data = list(reader)
            
            if not existing_data:
                # Leere Datei, nur Header erstellen
                header = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]
                return [header] + new_data
            
            # Prüfe ob erste Zeile ein Header ist
            first_row = existing_data[0]
            has_header = False
            
            if first_row and len(first_row) >= 3 and (first_row[0] == "ID" or first_row[0] == "id"):
                has_header = True
                header = first_row
                data_rows = existing_data[1:]
            else:
                # Kein Header gefunden, Standard verwenden
                header = ["ID", "DATUM", "STUNDEN", "KAPAZITÄT"]
                data_rows = existing_data
            
            # Dictionary für schnellere Suche erstellen (ID+DATUM -> Zeilenindex)
            id_date_indices = {(row[0], row[1]): i for i, row in enumerate(data_rows)}
            
            # Neue Daten durchgehen
            merged_data = data_rows.copy()
            for row in new_data:
                if len(row) >= 2:
                    key = (row[0], row[1])
                    if key in id_date_indices:
                        # Vorhandenen Eintrag überschreiben
                        merged_data[id_date_indices[key]] = row
                    else:
                        # Neuen Eintrag hinzufügen
                        merged_data.append(row)
            
            # Header wieder hinzufügen
            return [header] + merged_data
        
        except Exception as e:
            raise Exception(f"Fehler beim Zusammenführen der Daten: {str(e)}")
    
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

    def _process_ar_calendar_file(self, file_path, selected_member_id):
        """
        Robuste Excel-Import-Methode für problematische Dateien mit Formatierungsfehlern
        
        Args:
            file_path: Pfad zur Excel-Datei
            selected_member_id: ID des ausgewählten Projektmitglieds
            
        Returns:
            List[List[str]]: Verarbeitete Daten im Format [[ID, DATUM, STUNDEN, KAPAZITÄT], ...]
        """
        import pandas as pd
        import os
        import datetime
        import tempfile
        import csv
        
        # Liste für verarbeitete Daten
        processed_data = []
        temp_files = []
        
        # Debug-Funktion, die sowohl Konsolenausgabe als auch UI-Status aktualisiert
        def log_info(message, is_error=False):
            print(f"LOG: {message}")
            self._show_status(message, error=is_error)
        
        try:
            log_info(f"Verarbeite Datei: {os.path.basename(file_path)}")
            
            # ROBUSTES EXCEL-LESEN MIT MEHREREN FALLBACK-METHODEN
            df = None
            success = False
            
            # Versuch 1: Mit expliziter Engine-Angabe und ohne Formatierungen
            try:
                log_info("Methode 1: Lese Excel ohne Formatierungen...")
                df = pd.read_excel(file_path, engine='openpyxl', style_compression=True)
                success = True
                log_info("Methode 1 erfolgreich!")
            except Exception as e1:
                log_info(f"Methode 1 fehlgeschlagen: {str(e1)}")
            
            # Versuch 2: Mit alternativer Engine
            if not success:
                try:
                    log_info("Methode 2: Versuche alternative Engine...")
                    # Probiere xlrd für ältere Excel-Formate
                    df = pd.read_excel(file_path, engine='xlrd')
                    success = True
                    log_info("Methode 2 erfolgreich!")
                except Exception as e2:
                    log_info(f"Methode 2 fehlgeschlagen: {str(e2)}")
            
            # Versuch 3: Mit openpyxl und data_only
            if not success:
                try:
                    log_info("Methode 3: Verwende openpyxl mit data_only=True...")
                    from openpyxl import load_workbook
                    
                    # Excel zu temporärer CSV konvertieren
                    temp_csv = tempfile.NamedTemporaryFile(suffix='.csv', delete=False)
                    temp_csv.close()
                    temp_files.append(temp_csv.name)
                    
                    # Lade Workbook und ignoriere Formatierungen
                    wb = load_workbook(filename=file_path, read_only=True, data_only=True)
                    ws = wb.active
                    
                    # In CSV schreiben
                    with open(temp_csv.name, 'w', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        for row in ws.rows:
                            writer.writerow([cell.value for cell in row])
                    
                    # DataFrame aus CSV erstellen
                    df = pd.read_csv(temp_csv.name)
                    success = True
                    log_info("Methode 3 erfolgreich!")
                except Exception as e3:
                    log_info(f"Methode 3 fehlgeschlagen: {str(e3)}")
            
            # Versuch 4: Fallback auf csv-Leser für Excel-Dateien
            if not success:
                try:
                    log_info("Methode 4: Verwende PyExcelerate fallback...")
                    # Dieser Code könnte helfen, aber erfordert die Installation von pyexcelerate
                    # Für einfacheren Fallback verwenden wir eine direkte CSV-Konvertierung
                    
                    # Dateierweiterung prüfen
                    import subprocess
                    
                    # Temporäres CSV erstellen
                    temp_csv = tempfile.NamedTemporaryFile(suffix='.csv', delete=False)
                    temp_csv.close()
                    temp_files.append(temp_csv.name)
                    
                    # Hier würde man normalerweise externes Konvertierungstool aufrufen
                    # Da wir das nicht haben, zeigen wir eine Anleitung
                    log_info("Excel-Datei kann nicht automatisch verarbeitet werden.")
                    log_info("Bitte speichern Sie die Datei manuell als CSV und versuchen Sie es erneut.")
                    
                    # Versuche trotzdem zu lesen, vielleicht klappt's
                    try:
                        df = pd.read_csv(file_path, sep=None, engine='python')
                        success = True
                        log_info("Notfall-CSV-Lesung erfolgreich!")
                    except:
                        pass
                except Exception as e4:
                    log_info(f"Methode 4 fehlgeschlagen: {str(e4)}")
            
            # Wenn alle Methoden fehlgeschlagen sind
            if not success or df is None:
                log_info("Alle Lesemethoden fehlgeschlagen! Die Excel-Datei scheint beschädigt zu sein.", is_error=True)
                log_info("EMPFEHLUNG: Öffnen Sie die Datei in Excel und speichern Sie als CSV.", is_error=True)
                return []
            
            # Ab hier ist df verfügbar
            log_info(f"Excel erfolgreich geladen: {len(df.columns)} Spalten, {len(df)} Zeilen")
            
            # Zeige Spalten an
            log_info(f"Gefundene Spalten: {', '.join(str(col) for col in df.columns)}")
            
            # SCHRITT 2: Spalten identifizieren
            date_column = None
            hours_column = None
            desc_column = None
            
            # Listen mit möglichen Spaltennamen
            date_keywords = ["date", "datum", "tag", "day", "dates", "calendar", "termin", "zeit"]
            hours_keywords = ["total", "hours", "stunden", "zeit", "summe", "totalhours", "hour", "working", "hours_sum", "time"]
            desc_keywords = ["desc", "description", "beschreibung", "comment", "kommentar", "remarks", "note", "notes", "text"]
            
            # Suche nach Spalten (case-insensitive)
            for col in df.columns:
                col_str = str(col).lower().strip()
                
                # Für Datumspalte
                if not date_column:
                    if any(kw == col_str for kw in date_keywords) or any(kw in col_str for kw in date_keywords):
                        date_column = col
                        log_info(f"Datumsspalte gefunden: '{col}'")
                
                # Für Stundenspalte
                if not hours_column:
                    if any(kw == col_str for kw in hours_keywords) or any(kw in col_str for kw in hours_keywords):
                        hours_column = col
                        log_info(f"Stundenspalte gefunden: '{col}'")
                
                # Für Beschreibungsspalte
                if not desc_column:
                    if any(kw == col_str for kw in desc_keywords) or any(kw in col_str for kw in desc_keywords):
                        desc_column = col
                        log_info(f"Beschreibungsspalte gefunden: '{col}'")
            
            # Typ-basierte Erkennung als Fallback
            # Wenn keine Datumsspalte gefunden wurde
            if not date_column:
                for col in df.columns:
                    # Prüfe auf Datums-Datentyp
                    if pd.api.types.is_datetime64_any_dtype(df[col]):
                        date_column = col
                        log_info(f"Datumsspalte durch Datentyp gefunden: '{col}'")
                        break
            
            # Wenn keine Stundenspalte gefunden wurde
            if not hours_column:
                for col in df.columns:
                    if col != date_column and pd.api.types.is_numeric_dtype(df[col]):
                        # Nehme die erste numerische Spalte, die keine Datumsspalte ist
                        hours_column = col
                        log_info(f"Stundenspalte durch Datentyp gefunden: '{col}'")
                        break
            
            # Durchsuche alle Spalten nach FCAST/Forecast
            if not desc_column:
                for col in df.columns:
                    if col != date_column and col != hours_column:
                        try:
                            # Konvertiere zu String und suche nach FCAST/Forecast
                            col_data = df[col].astype(str).str.upper()
                            if any(term in val for val in col_data for term in ["FCAST", "FORECAST"]):
                                desc_column = col
                                log_info(f"Beschreibungsspalte durch FCAST/Forecast gefunden: '{col}'")
                                break
                        except:
                            continue
            
            # Wenn keine Spalten identifiziert wurden, nehmen wir die ersten verfügbaren
            if not date_column and len(df.columns) > 0:
                date_column = df.columns[0]
                log_info(f"Keine Datumsspalte erkannt! Verwende erste Spalte: '{date_column}'", is_error=True)
            
            if not hours_column and len(df.columns) > 1:
                for col in df.columns:
                    if col != date_column:
                        hours_column = col
                        log_info(f"Keine Stundenspalte erkannt! Verwende Spalte: '{hours_column}'", is_error=True)
                        break
            
            # Prüfen, ob wir weitermachen können
            if not date_column or not hours_column:
                log_info("Erforderliche Spalten nicht gefunden!", is_error=True)
                return []
            
            # SCHRITT 3: Daten filtern (wenn möglich)
            filtered_df = df
            if desc_column:
                try:
                    log_info(f"Filtere nach FCAST/Forecast in '{desc_column}'...")
                    
                    # Konvertiere zu String (falls nötig) und suche nach FCAST/Forecast
                    if not pd.api.types.is_string_dtype(df[desc_column]):
                        desc_data = df[desc_column].astype(str)
                    else:
                        desc_data = df[desc_column]
                    
                    # Case-insensitive Suche
                    mask = desc_data.str.upper().str.contains('FCAST|FORECAST', na=False)
                    filtered_df = df[mask]
                    
                    log_info(f"Nach Filterung verbleiben {len(filtered_df)} von {len(df)} Zeilen")
                    
                    if filtered_df.empty:
                        log_info("Keine Einträge mit FCAST/Forecast gefunden. Verwende alle Zeilen.", is_error=True)
                        filtered_df = df
                except Exception as filter_error:
                    log_info(f"Fehler bei Filterung: {str(filter_error)}", is_error=True)
                    filtered_df = df
            else:
                log_info("Keine Beschreibungsspalte für FCAST/Forecast-Filter gefunden.")
            
            # SCHRITT 4: Daten verarbeiten
            log_info("Verarbeite gefilterte Daten...")
            successful_rows = 0
            error_rows = 0
            
            for index, row in filtered_df.iterrows():
                try:
                    # Datum verarbeiten
                    date_value = row[date_column]
                    
                    # Verschiedene Datumsformate behandeln
                    if pd.isna(date_value):
                        error_rows += 1
                        continue
                    
                    if isinstance(date_value, (pd.Timestamp, datetime.datetime, datetime.date)):
                        date_str = date_value.strftime('%d.%m.%Y')
                    elif isinstance(date_value, str):
                        try:
                            # Versuche zu parsen
                            parsed_date = pd.to_datetime(date_value)
                            date_str = parsed_date.strftime('%d.%m.%Y')
                        except:
                            # Falls nicht parse-bar, prüfe auf gängige Formate
                            import re
                            # Mögliche Datumsformate: DD.MM.YYYY, MM/DD/YYYY, YYYY-MM-DD
                            if re.match(r'\d{1,2}[./-]\d{1,2}[./-]\d{2,4}', date_value):
                                # Plausibles Datumsformat, behalten
                                date_str = date_value
                            else:
                                # Kein plausibles Datum, überspringen
                                error_rows += 1
                                continue
                    else:
                        # Fallback für numerische Werte (Excel-Datum)
                        try:
                            excel_date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(date_value))
                            date_str = excel_date.strftime('%d.%m.%Y')
                        except:
                            # Kein konvertierbarer Wert
                            error_rows += 1
                            continue
                    
                    # Stunden verarbeiten mit mehr Fehlertoleranz
                    hours_value = row[hours_column]
                    
                    # Überspringe Zeilen ohne Stundenwerte
                    if pd.isna(hours_value):
                        error_rows += 1
                        continue
                    
                    # Konvertierung der Stunden mit mehr Fehlertoleranz
                    try:
                        # Bereinige und konvertiere
                        if isinstance(hours_value, str):
                            # Verschiedene Formate bereinigen
                            cleaned_hours = hours_value.replace(',', '.').strip()
                            # Entferne alles außer Ziffern und Punkt
                            import re
                            cleaned_hours = re.sub(r'[^\d.]', '', cleaned_hours)
                            stunden_float = float(cleaned_hours) if cleaned_hours else 0
                        else:
                            stunden_float = float(hours_value)
                        
                        # Ignoriere unplausible Werte
                        if stunden_float <= 0 or stunden_float > 24:
                            error_rows += 1
                            continue
                        
                        stunden = str(stunden_float)
                        
                        # Kapazität berechnen
                        if stunden_float >= 8:
                            kapazitaet = "1.0"
                        else:
                            kapazitaet = str(round(stunden_float / 8, 2))
                        
                        # Daten zur Liste hinzufügen
                        processed_data.append([selected_member_id, date_str, stunden, kapazitaet])
                        successful_rows += 1
                        
                    except (ValueError, TypeError) as e:
                        error_rows += 1
                        continue
                
                except Exception as row_error:
                    error_rows += 1
                    continue
            
            # Zusammenfassung
            if successful_rows > 0:
                log_info(f"Import erfolgreich: {successful_rows} Einträge verarbeitet, {error_rows} übersprungen")
                return processed_data
            else:
                log_info(f"Keine gültigen Daten gefunden! Alle {error_rows} Zeilen fehlerhaft.", is_error=True)
                return []
        
        except Exception as e:
            log_info(f"Kritischer Fehler: {str(e)}", is_error=True)
            return []
        
        finally:
            # Bereinigen: Temporäre Dateien löschen
            for temp_file in temp_files:
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
   
    
    
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