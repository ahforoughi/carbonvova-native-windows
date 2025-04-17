import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'excel_monitor/main.py',  # Main entry point
    '--name=ExcelMonitor',
    '--onefile',
    '--windowed',
    '--add-data=requirements.txt:.',
    '--hidden-import=excel_monitor.gui',
    '--hidden-import=excel_monitor.monitor',
    '--hidden-import=excel_monitor.utils',
    '--clean',
    '--noconfirm'
]) 