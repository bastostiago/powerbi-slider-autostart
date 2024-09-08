import tkinter as tk
from tkinter import messagebox
import os
import configparser
import webbrowser
import pyautogui
import time
import re
import sys
from manifest import _MANIFEST
import win32com.client
import logging
import pygetwindow as gw

_TOTAL_SCREEN_SHOTS = 9


def is_in_startup():
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    shortcut_path = os.path.join(startup_folder, f'{_MANIFEST["app_name"]}.lnk')
    return os.path.exists(shortcut_path)


def add_to_startup():
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    exe_path = os.path.join(os.getcwd(), f'{_MANIFEST["app_name"]}.exe')  
    shortcut_path = os.path.join(startup_folder, f'{_MANIFEST["app_name"]}.lnk')    
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = exe_path
    shortcut.Arguments = "noopenapp"  
    shortcut.WorkingDirectory = os.getcwd()
    shortcut.IconLocation = exe_path  
    shortcut.save()


def remove_from_startup():
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    shortcut_path = os.path.join(startup_folder, f'{_MANIFEST["app_name"]}.lnk')

    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)


def save_and_update_startup():    

    config_value = text_area.get("1.0", tk.END).strip()    
    with open(config_file, 'w') as config_file_object:
        config_file_object.write(f'[DEFAULT]\nurl={config_value}')

    if startup_var.get():
        add_to_startup()
    else:
        remove_from_startup()

    messagebox.showinfo("Success", "Data has been saved!")


def launcher():
    
    if app:
        url = text_area.get("1.0", tk.END).strip()
    else:
        url = existing_url

    if not url:
        logging.error("No configured URL.")
        return
    webbrowser.open(url.split('?')[0])

    time.sleep(7)    

    browser_window = None
    for window in gw.getAllTitles():
        if "Chrome" in window or "Edge" in window or "Firefox" in window:  
            browser_window = window
            break

    if browser_window:
        window = gw.getWindowsWithTitle(browser_window)[0]
        window.activate()  
        
        retries = 50
        while retries > 0:

            found = False
            image_index = 1
            while not found and image_index <= _TOTAL_SCREEN_SHOTS:
                try:
                    button_location = pyautogui.locateOnScreen(resource_path(f'screenshots/screen_shot_pm_{image_index}.png'), confidence=0.75)                
                    if button_location:
                        pyautogui.click(button_location)  # Simulates button click
                        found = True
                    else:
                        found = False
                except pyautogui.ImageNotFoundException:
                    found = False
                image_index += 1
            
            if found:
                retries = 0
            else:
                logging.error(f"Try {retries} - Fullscreen button not found.")
                time.sleep(5)
                retries -= 1
            
    else:
        logging.error("Browser window not found.")


def check_url(urlIn=None):
    error = None
    if urlIn:
        url = urlIn
    else:
        url = text_area.get("1.0", tk.END).strip()

    if re.match(r'^(https?://)?(localhost|powerbislider.com|www.powerbislider.com)', url):
        if launch_button:
            launch_button.config(state=tk.NORMAL, text="Test URL")         
    else:
        if launch_button:
            launch_button.config(state=tk.DISABLED, text="Invalid URL")
        error = "Invalid URL!"

    return error


def resource_path(relative_path):
    """ Get absolute path to resource, works for PyInstaller """
    try:
        
        base_path = sys._MEIPASS
    except AttributeError:
        
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    app = None

    _LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
    config_dir = os.path.expanduser("~\\AppData\\Local\\PB Slider Launcher")
    
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)

    log_file = os.path.join(config_dir, 'error.log')
    config_file = os.path.join(config_dir, 'config.ini')


    logging.basicConfig(filename=log_file, level="ERROR", format=_LOG_FORMAT)
    
    config = configparser.ConfigParser()
    if os.path.exists(config_file):
        config.read(config_file)
    existing_url = config['DEFAULT']['url'] if 'url' in config['DEFAULT'] else ''

    open_app = True
    launch_button = None

    if len(sys.argv) > 1 and sys.argv[1] == "noopenapp":
        msgError = check_url(existing_url)
        if not msgError:
            launcher()
        else:
            logging.error(msgError)
    else:
        app = tk.Tk()
        app.title(f"Settings - PB Slider Launcher (v{_MANIFEST["version"]})")
        
        app.geometry('400x265')  
        app.eval('tk::PlaceWindow . center')  
        app.resizable(False, False)

        icon_path = resource_path("logo.ico")
        app.iconbitmap(icon_path)
        
        tk.Label(app, text="Power BI Slider URL:").pack(pady=5)
        
        text_area = tk.Text(app, height=6, width=40)
        text_area.pack(pady=(0, 10))        
        text_area.insert(tk.END, existing_url)        
        text_area.bind("<KeyRelease>", lambda event: check_url())
        
        launch_button = tk.Button(app, text="Test URL", command=launcher, state=tk.DISABLED, width=10)
        launch_button.pack(padx=37, anchor='e')
        
        startup_var = tk.BooleanVar()        
        
        if is_in_startup():
            startup_var.set(True)
        
        startup_checkbox = tk.Checkbutton(app, text="Start on System Startup", variable=startup_var)
        startup_checkbox.pack(pady=5)
        
        tk.Button(app, text="Save", command=save_and_update_startup, width=35, height=2).pack(pady=5)
        
        check_url()

        app.mainloop()
