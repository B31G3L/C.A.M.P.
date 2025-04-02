"""
Einfaches Skript zum Erstellen von App-Icons für CAMP
"""
import os
import sys

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("PIL/Pillow ist nicht installiert. Installieren Sie es mit 'pip install pillow'")
    sys.exit(1)

def create_directory(path):
    """Erstellt ein Verzeichnis, falls es nicht existiert"""
    if not os.path.exists(path):
        os.makedirs(path)

def create_icon(size, output_path, text="CAMP"):
    """Erstellt ein einfaches Icon mit Text"""
    # Erstelle ein quadratisches Bild mit transparentem Hintergrund
    img = Image.new('RGBA', (size, size), color=(0, 0, 0, 0))
    
    # Erstelle einen Draw-Kontext
    draw = ImageDraw.Draw(img)
    
    # Zeichne einen blauen Kreis
    circle_radius = size // 2
    draw.ellipse((0, 0, size, size), fill=(30, 120, 200, 255))
    
    # Füge Text hinzu
    try:
        # Versuche, eine Schriftart zu laden
        font_size = size // 3
        font = ImageFont.truetype("arial.ttf", font_size)
    except IOError:
        # Fallback auf die Standardschriftart
        font = ImageFont.load_default()
    
    # Text zentrieren
    text_width, text_height = draw.textsize(text, font=font) if hasattr(draw, 'textsize') else (font_size * len(text), font_size)
    position = ((size - text_width) // 2, (size - text_height) // 2)
    
    # Text zeichnen
    draw.text(position, text, fill=(255, 255, 255, 255), font=font)
    
    # Speichern
    img.save(output_path)
    print(f"Icon erstellt: {output_path}")

def create_windows_icon(base_dir):
    """Erstellt ein Windows-Icon (.ico)"""
    # Windows-Icon-Größen
    sizes = [16, 32, 48, 64, 128, 256]
    icon_path = os.path.join(base_dir, "camp_icon.ico")
    
    # Temporäre PNG-Dateien erstellen
    temp_images = []
    for size in sizes:
        temp_path = os.path.join(base_dir, f"temp_{size}.png")
        create_icon(size, temp_path)
        temp_images.append(Image.open(temp_path))
    
    # Multi-Size-Icon erstellen und speichern
    temp_images[0].save(
        icon_path, 
        format='ICO', 
        sizes=[(size, size) for size in sizes],
        append_images=temp_images[1:]
    )
    
    # Temporäre Dateien löschen
    for size in sizes:
        temp_path = os.path.join(base_dir, f"temp_{size}.png")
        if os.path.exists(temp_path):
            os.remove(temp_path)

def create_macos_icon(base_dir):
    """Erstellt ein macOS-Icon (.icns)"""
    # Hier würde normalerweise ein .icns-Icon erstellt
    # Das erfordert aber iconutil, das nur auf macOS verfügbar ist
    # Stattdessen erstellen wir einfach ein großes PNG als Platzhalter
    
    icon_path = os.path.join(base_dir, "camp_icon.png")
    create_icon(1024, icon_path)
    print("Hinweis: Um ein richtiges .icns-Icon zu erstellen, müsste das Skript auf macOS ausgeführt werden.")
    
    # Platzhalter .icns-Datei erstellen
    icns_path = os.path.join(base_dir, "camp_icon.icns")
    with open(icns_path, 'wb') as f:
        f.write(b'icns')  # Minimaler Dateityp-Header
    print(f"Platzhalter-Icon erstellt: {icns_path}")

def main():
    """Hauptfunktion"""
    # Basisverzeichnis
    script_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(script_dir)
    assets_dir = os.path.join(base_dir, "assets")
    
    # Verzeichnis erstellen
    create_directory(assets_dir)
    
    # Icons erstellen
    create_windows_icon(assets_dir)
    create_macos_icon(assets_dir)
    
    print("Icons erfolgreich erstellt!")

if __name__ == "__main__":
    main()