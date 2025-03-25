# CAMP - Dokumentation

## Übersicht

CAMP (Capacity Analysis & Management Planner) ist ein Tool zur Verwaltung und Analyse von Projektkapazitäten und Sprints.

## Struktur

Die Anwendung ist in mehrere Module unterteilt:

### Core

Das Core-Modul enthält die Hauptanwendungsklasse (`app.py`), die als Einstiegspunkt für die Anwendungslogik dient.

### Models

Das Models-Modul enthält die Datenstrukturen für:
- Projekte (`project.py`)
- Sprints (`sprint.py`)
- Teammitglieder (`team_member.py`)

### Services

Services kapseln die Geschäftslogik und Datenzugriffe:
- `data_service.py`: Zugriff auf Projektdaten und Kapazitätsdaten
- `import_service.py`: Import von Daten aus externen Quellen
- `export_service.py`: Export von Daten in verschiedene Formate

### UI

Das UI-Modul ist in drei Untermodule unterteilt:
- `views`: Hauptansichten der Anwendung
- `components`: Wiederverwendbare UI-Komponenten
- `dialogs`: Modale Dialoge für verschiedene Funktionen

### Utils

Das Utils-Modul enthält allgemeine Hilfsfunktionen:
- `date_utils.py`: Funktionen für Datumsoperationen
- `persistence.py`: Funktionen für die Datenpersistenz
- `config_loader.py`: Laden von Konfigurationen

## Konfiguration

Die Anwendung unterstützt verschiedene Konfigurationen:
- `default.py`: Standardkonfiguration
- `dev.py`: Entwicklungskonfiguration

## Daten

Die Anwendung verwendet zwei Hauptdatenstrukturen:
- `project.json`: Enthält Projektdaten, Sprints und Teammitglieder
- `kapa_data.csv`: Enthält Kapazitätsdaten im Format "ID;DATUM;STUNDEN;KAPAZITÄT"