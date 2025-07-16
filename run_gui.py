#!/usr/bin/env python3
"""
Professor Evaluation Report Generator - GUI Launcher
====================================================

This script launches the graphical user interface for generating 
professor evaluation reports from Excel data.

Usage:
    python run_gui.py

Features:
- Browse and select Excel files
- Choose specific professor or generate reports for all professors
- Automatic file handling and PDF generation
- Reports saved directly to Downloads folder
"""

import sys
import os

# Add current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from gui_app import main
    
    if __name__ == "__main__":
        print("Starting Professor Evaluation Report Generator...")
        print("GUI interface loading...")
        main()
        
except ImportError as e:
    print(f"Error importing required modules: {e}")
    print("\nPlease make sure all dependencies are installed:")
    print("pip install pandas matplotlib reportlab Pillow")
    print("\nNote: tkinter should be included with Python by default")
    input("\nPress Enter to exit...")
    
except Exception as e:
    print(f"An error occurred: {e}")
    input("\nPress Enter to exit...")
