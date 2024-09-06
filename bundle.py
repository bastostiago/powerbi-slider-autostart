import subprocess
import shutil 
import os

APP_NAME = 'PB Slider Launcher'

# Remove existing build and dist directories
if os.path.exists('./build'):
    shutil.rmtree('./build')
if os.path.exists('./dist'):
    shutil.rmtree('./dist')

subprocess.run([
    "pyinstaller",
    "--onefile",
    "--icon=logo.ico",
    f"--name={APP_NAME}",    
    f"--add-data=logo.ico;.", 
    "--noconsole",
    "main.py"
], check=True)