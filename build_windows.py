import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'excel_monitor/main.py',
    '--name=ExcelMonitor',
    '--onefile',
    '--windowed',
    f'--add-data={os.path.join(current_dir, "requirements.txt")};.',  # Fixed syntax
    '--hidden-import=excel_monitor.gui',
    '--hidden-import=excel_monitor.monitor',
    '--hidden-import=excel_monitor.utils',
    '--clean',
    '--noconfirm',
    '--icon=NONE'  # You can add a .ico file for Windows icon
]) 