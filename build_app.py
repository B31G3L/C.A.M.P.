"""
Build-Skript für CAMP
Erstellt eigenständige ausführbare Dateien für Windows und macOS
"""
import os
import sys
import shutil
import platform
import subprocess
import argparse
from src.utils.version import __version__

def parse_args():
    """Kommandozeilenargumente parsen"""
    parser = argparse.ArgumentParser(description="Build CAMP application")
    parser.add_argument("--clean", action="store_true", help="Clean build directories before building")
    parser.add_argument("--skipbuild", action="store_true", help="Skip build and go straight to packaging")
    return parser.parse_args()

def clean_build_directories():
    """Bereinigt Build-Verzeichnisse"""
    print("Cleaning build directories...")
    
    directories = ["build", "dist"]
    for directory in directories:
        if os.path.exists(directory):
            print(f"Removing {directory}...")
            shutil.rmtree(directory)
    
    # Entferne .spec-Dateien
    for file in os.listdir("."):
        if file.endswith(".spec") and file != "build_app.spec":
            print(f"Removing {file}...")
            os.remove(file)

def build_application():
    """Baut die Anwendung mit PyInstaller"""
    print("Building application...")
    
    # PyInstaller-Befehl mit der .spec-Datei
    cmd = [
        "pyinstaller",
        "--clean",
        "build_app.spec"
    ]
    
    # Befehl ausführen
    subprocess.run(cmd, check=True)

def create_distribution_package():
    """Erstellt ein Distributionspaket je nach Plattform"""
    print("Creating distribution package...")
    
    system = platform.system().lower()
    version = __version__
    
    if system == "windows":
        # Windows: ZIP-Datei erstellen
        dist_dir = os.path.join("dist", "CAMP")
        output_zip = f"CAMP-{version}-win64.zip"
        
        if os.path.exists(output_zip):
            os.remove(output_zip)
        
        print(f"Creating ZIP file: {output_zip}...")
        shutil.make_archive(f"CAMP-{version}-win64", "zip", "dist", "CAMP")
        
        print(f"Distribution package created: {output_zip}")
    
    elif system == "darwin":
        # macOS: DMG-Datei erstellen
        app_path = os.path.join("dist", "CAMP.app")
        output_dmg = f"CAMP-{version}-macos.dmg"
        
        if os.path.exists(output_dmg):
            os.remove(output_dmg)
        
        print(f"Creating DMG file: {output_dmg}...")
        try:
            # Verwende create-dmg oder hdituil
            cmd = [
                "hdiutil", "create",
                "-volname", f"CAMP {version}",
                "-srcfolder", app_path,
                "-ov", output_dmg
            ]
            subprocess.run(cmd, check=True)
            print(f"Distribution package created: {output_dmg}")
        except:
            print("Could not create DMG file. Please install create-dmg or use hdiutil manually.")
            print(f"The application can be found at: {app_path}")
    
    else:
        # Linux: Tar.gz-Datei erstellen
        dist_dir = os.path.join("dist", "CAMP")
        output_tar = f"CAMP-{version}-linux.tar.gz"
        
        if os.path.exists(output_tar):
            os.remove(output_tar)
        
        print(f"Creating TAR.GZ file: {output_tar}...")
        shutil.make_archive(f"CAMP-{version}-linux", "gztar", "dist", "CAMP")
        
        print(f"Distribution package created: {output_tar}")

def main():
    """Hauptfunktion"""
    args = parse_args()
    
    if args.clean:
        clean_build_directories()
    
    if not args.skipbuild:
        build_application()
    
    create_distribution_package()
    
    print("Build completed successfully!")

if __name__ == "__main__":
    main()