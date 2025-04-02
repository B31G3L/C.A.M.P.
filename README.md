# CAMP - Capacity Analysis & Management Planner

CAMP ist ein Tool zur Verwaltung und Analyse von Projektkapazitäten und Sprints.

## Funktionen

- Verwaltung von Projekten und Sprints
- Visualisierung von Team-Kapazitäten
- Sprint-Planung und Nachverfolgung
- Berechnung von Story Points basierend auf Team-Kapazität
- Export nach Confluence Wiki
- Automatische Updates

## Installation

### Vorkompilierte Versionen

Laden Sie die neueste Version von unserer [Releases-Seite](https://github.com/YOURUSERNAME/CAMP/releases) herunter:

- **Windows**: Laden Sie die `.zip`-Datei herunter, entpacken Sie sie und führen Sie `CAMP.exe` aus.
- **macOS**: Laden Sie die `.dmg`-Datei herunter, öffnen Sie sie und ziehen Sie die Anwendung in Ihren Programme-Ordner.

### Aus dem Quellcode

1. Klonen Sie dieses Repository:
```bash
git clone https://github.com/YOURUSERNAME/CAMP.git
cd CAMP
```

2. Installieren Sie die Abhängigkeiten:
```bash
pip install -r requirements.txt
```

3. Starten Sie die Anwendung:
```bash
python main.py
```

## Aus dem Quellcode bauen

### Voraussetzungen

- Python 3.10 oder höher
- PyInstaller (`pip install pyinstaller`)
- Alle Abhängigkeiten aus `requirements.txt`

### Build erstellen

1. Führen Sie das Build-Skript aus:
```bash
python build_app.py
```

2. Die kompilierten Dateien finden Sie im Verzeichnis `dist/`:
   - Windows: `dist/CAMP/CAMP.exe`
   - macOS: `dist/CAMP.app`

Für ein Distributionspaket wird automatisch eine Datei erstellt:
   - Windows: `CAMP-x.y.z-win64.zip`
   - macOS: `CAMP-x.y.z-macos.dmg`

### Icons erstellen (optional)

Falls Sie eigene Icons für die Anwendung erstellen möchten:

```bash
python assets/create_icons.py
```

## Automatische Updates

CAMP prüft beim Start automatisch auf neue Versionen. Wenn eine neue Version verfügbar ist, werden Sie benachrichtigt und können das Update direkt aus der Anwendung installieren.

Sie können auch manuell auf Updates prüfen:
1. Menü: **Hilfe > Auf Updates prüfen**

## Entwicklung

### Versionierung

Die Versionsnummer folgt dem Format `MAJOR.MINOR.PATCH` und wird in der Datei `src/utils/version.py` definiert.

### GitHub Actions

Dieses Projekt verwendet GitHub Actions für:
- Automatisierte Tests bei jedem Push
- Automatischer Build und Release bei Tags (`v*`)
- Automatische Versionsinkrementierung nach einem Release

### Neue Version erstellen

Um eine neue Version zu erstellen:

1. Erstellen Sie einen Git-Tag:
```bash
git tag -a v0.1.0 -m "Version 0.1.0"
git push origin v0.1.0
```

2. Der GitHub Actions Workflow erstellt automatisch Builds für Windows und macOS und veröffentlicht sie als Release.

3. Nach dem Release wird die Patch-Version automatisch inkrementiert für die weitere Entwicklung.

## Lizenz

[MIT](LICENSE)