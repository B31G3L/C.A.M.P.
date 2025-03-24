# CAMP - Capacity Analysis & Management Planner

CAMP ist ein Tool zur Verwaltung und Analyse von Projektkapazitäten und Sprints.

## Installation

1. Klone dieses Repository:
```bash
git clone https://your-repository-url/camp.git
cd camp
```

2. Installiere die Abhängigkeiten:
```bash
pip install -r requirements.txt
```

## Verwendung

Starte die Anwendung mit:
```bash
python main.py
```

## Funktionen

- Verwaltung von Projekten und Sprints
- Anzeige von Kapazitätsdaten
- Erstellung von Markup für Dokumentation
- Import von Zeiterfassungs- und Projektdaten

## Projektstruktur

```
camp/
├── main.py              # Einstiegspunkt der Anwendung
├── requirements.txt     # Abhängigkeiten
├── data/                # Datenverzeichnis
└── src/                 # Quellcode
    ├── app.py           # Hauptanwendungsklasse
    ├── ui/              # UI-Komponenten
    │   ├── toolbar.py   # Toolbar-Definition
    │   ├── main_view.py # Hauptansicht
    │   ├── data_modal.py # Datendarstellung
    │   ├── camp_manager_modal.py # Manager für Projekte und Mitarbeiter
    │   └── import_modal.py # Import-Funktionalität
    ├── models/          # Datenmodelle
    │   ├── project.py   # Projekt-Modell
    │   └── sprint.py    # Sprint-Modell
    └── utils/           # Hilfsfunktionen
        └── persistence.py # Datenpersistenz
```

## Lizenz

[MIT](LICENSE)