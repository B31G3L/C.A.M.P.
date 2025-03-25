"""
Konfigurationsloader für CAMP
"""
import importlib
import os

def load_config(config_name="default"):
    """
    Lädt die Konfiguration
    
    Args:
        config_name (str): Name der Konfiguration (z.B. "default", "dev")
        
    Returns:
        module: Das geladene Konfigurationsmodul
    """
    try:
        return importlib.import_module(f"config.{config_name}")
    except ImportError:
        # Fallback zur Standardkonfiguration
        return importlib.import_module("config.default")