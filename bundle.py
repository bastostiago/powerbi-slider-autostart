import subprocess
import shutil 
import os
from manifest import _MANIFEST

APP_NAME = _MANIFEST["app_name"]

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
    f"--add-data=screenshots;./screenshots", 
    "--noconsole",
    "main.py"
], check=True)