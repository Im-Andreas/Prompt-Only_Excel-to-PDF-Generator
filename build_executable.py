# Professor Report Generator - Build Script
# This script creates an executable (.exe) file for the GUI application

import subprocess
import sys
import os

def get_venv_python():
    """Get the Python executable from the virtual environment"""
    venv_paths = [
        "venv\\Scripts\\python.exe",  # Windows venv
        "venv\\python.exe",           # Alternative location
        ".venv\\Scripts\\python.exe", # .venv folder
        "env\\Scripts\\python.exe"    # env folder
    ]
    
    for path in venv_paths:
        if os.path.exists(path):
            print(f"✓ Found virtual environment Python: {path}")
            return os.path.abspath(path)
    
    print("⚠️  Virtual environment not found, using system Python")
    return sys.executable

def install_pyinstaller():
    """Install PyInstaller if not already installed"""
    python_exe = get_venv_python()
    
    try:
        # Check if PyInstaller is installed in the venv
        result = subprocess.run([python_exe, "-c", "import PyInstaller"], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print("✓ PyInstaller is already installed in virtual environment")
        else:
            raise ImportError
    except (ImportError, subprocess.CalledProcessError):
        print("Installing PyInstaller in virtual environment...")
        subprocess.check_call([python_exe, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed successfully in virtual environment")

def build_executable():
    """Build the executable using PyInstaller"""
    print("\nBuilding executable...")
    
    python_exe = get_venv_python()
    
    # Try to find pyinstaller in the virtual environment
    venv_scripts_dir = os.path.dirname(python_exe)
    pyinstaller_exe = os.path.join(venv_scripts_dir, "pyinstaller.exe")
    
    # PyInstaller command - prefer venv version
    if os.path.exists(pyinstaller_exe):
        print(f"✓ Using PyInstaller from virtual environment: {pyinstaller_exe}")
        pyinstaller_cmd = pyinstaller_exe
    else:
        print("✓ Using PyInstaller via Python module")
        pyinstaller_cmd = f"{python_exe} -m PyInstaller"
    
    # Build command arguments
    if pyinstaller_cmd.endswith(".exe"):
        cmd = [
            pyinstaller_cmd,
            "--onefile",                    # Create a single executable file
            "--windowed",                   # Hide console window (GUI app)
            "--name=Professor_Report_Generator",  # Name of the executable
            "--add-data=assets;assets",     # Include assets folder
            "--hidden-import=pandas",       # Explicitly include pandas
            "--hidden-import=matplotlib",   # Explicitly include matplotlib
            "--hidden-import=reportlab",    # Explicitly include reportlab
            "--hidden-import=PIL",          # Explicitly include Pillow
            "--hidden-import=tkinter",      # Explicitly include tkinter
            "--hidden-import=openpyxl",     # Explicitly include openpyxl
            "run_gui.py"                    # Main script to build
        ]
    else:
        cmd = [
            python_exe, "-m", "PyInstaller",
            "--onefile",                    # Create a single executable file
            "--windowed",                   # Hide console window (GUI app)
            "--name=Professor_Report_Generator",  # Name of the executable
            "--add-data=assets;assets",     # Include assets folder
            "--hidden-import=pandas",       # Explicitly include pandas
            "--hidden-import=matplotlib",   # Explicitly include matplotlib
            "--hidden-import=reportlab",    # Explicitly include reportlab
            "--hidden-import=PIL",          # Explicitly include Pillow
            "--hidden-import=tkinter",      # Explicitly include tkinter
            "--hidden-import=openpyxl",     # Explicitly include openpyxl
            "run_gui.py"                    # Main script to build
        ]
    
    # Add icon if it exists
    if os.path.exists("assets/app_icon.ico"):
        cmd.insert(-1, "--icon=assets/app_icon.ico")
    
    try:
        print(f"Running: {' '.join(cmd)}")
        subprocess.check_call(cmd)
        print("\n✓ Executable built successfully using virtual environment!")
        print("📁 Location: dist/Professor_Report_Generator.exe")
        print("\n🚀 You can now run the application by double-clicking the .exe file!")
        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Error building executable: {e}")
        return False
    
    return True

def create_batch_file():
    """Create a simple batch file to run the Python script"""
    python_exe = get_venv_python()
    
    batch_content = f"""@echo off
title Professor Report Generator
echo Starting Professor Report Generator...

REM Check if virtual environment Python exists
if exist "{python_exe}" (
    echo Using virtual environment Python: {python_exe}
    "{python_exe}" run_gui.py
) else (
    echo Virtual environment not found, trying system Python...
    python run_gui.py
)

if errorlevel 1 (
    echo.
    echo Error: Python or required packages not found.
    echo Please ensure the virtual environment is set up correctly.
    echo Or run: pip install -r requirements.txt
    echo.
    pause
)
"""
    
    with open("Start_Professor_Report_Generator.bat", "w") as f:
        f.write(batch_content)
    
    print("✓ Created Start_Professor_Report_Generator.bat (using virtual environment)")
    print("  You can double-click this .bat file to run the application")

if __name__ == "__main__":
    print("=" * 60)
    print("Professor Report Generator - Executable Builder")
    print("=" * 60)
    
    # Check if we're in the right directory
    if not os.path.exists("run_gui.py"):
        print("❌ Error: run_gui.py not found!")
        print("Please run this script from the project directory.")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Install PyInstaller
    install_pyinstaller()
    
    # Create batch file as backup
    create_batch_file()
    
    # Build executable
    success = build_executable()
    
    if success:
        print("\n" + "=" * 60)
        print("BUILD COMPLETE!")
        print("=" * 60)
        print("You now have two ways to run the application:")
        print("1. 📁 dist/Professor_Report_Generator.exe (Standalone executable)")
        print("2. 📄 Start_Professor_Report_Generator.bat (Batch file)")
        print("\nBoth files can be moved anywhere and will work independently!")
    else:
        print("\n" + "=" * 60)
        print("BUILD FAILED!")
        print("=" * 60)
        print("You can still use the batch file to run the application:")
        print("📄 Start_Professor_Report_Generator.bat")
    
    print("\n" + "=" * 60)
    input("Press Enter to exit...")
