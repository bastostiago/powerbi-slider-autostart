import tkinter as tk
from tkinter import messagebox
import os
import configparser
import re
import sys
from manifest import _MANIFEST  # Certifique-se de que este módulo está corretamente configurado
import win32com.client
import logging
import ctypes
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import logging
import time

_CONFIG_DIR = os.path.expanduser("~\\AppData\\Local\\PB Slider Launcher")

# Tornar o aplicativo DPI-aware (Windows)
if sys.platform.startswith('win'):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception as e:
        print(f"Erro ao definir consciência de DPI: {e}")

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
    shortcut.Arguments = "launch"  
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

def create_driver(navegador="chrome"):
    navegador = navegador.lower()

    status_var.set(f"Checking driver for {navegador}...")
    app.update_idletasks()

    try:
        if navegador == "chrome":
            service = ChromeService(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument(f"--user-data-dir={os.path.join(_CONFIG_DIR, 'ChromeUserData')}")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            driver = webdriver.Chrome(service=service, options=options)
        elif navegador == "edge":
            service = EdgeService(EdgeChromiumDriverManager().install())
            options = webdriver.EdgeOptions()
            options.use_chromium = True
            options.add_argument('--start-maximized')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument(f"--user-data-dir={os.path.join(_CONFIG_DIR, 'EdgeUserData')}")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            options.add_argument("--profile-directory=Default")
            driver = webdriver.Edge(service=service, options=options)
        
        status_var.set(f"Driver for {navegador} has been loaded successfully.")
        app.update_idletasks()

        return driver
    except Exception as e:
        logging.error(f"Error to create driver {navegador}: {e}")
        status_var.set("Error by loading driver.")
        app.update_idletasks()
        raise

def automate(navegador="chrome", url=None):
    tempo_maximo_espera = 600  # Tempo máximo de espera em segundos
    driver = None

    try:
        driver = create_driver(navegador)
        status_var.set(f"Visiting URL ({navegador})...")
        app.update_idletasks()

        driver.get(url)

        wait = WebDriverWait(driver, tempo_maximo_espera)
        botao_apresentacao = wait.until(
            EC.element_to_be_clickable((By.ID, "presentationButton"))
        )

        botao_apresentacao.click()        
        app.update_idletasks()

        while True:
            try:
                driver.title
                time.sleep(1)
            except Exception:
                logging.info("Browser has closed, finishing automation...")
                break

        return True

    except Exception as e:
        logging.error(f"Error during the automation: {e}")
        status_var.set("Error during the automation.")
        app.update_idletasks()
        if driver:
            driver.quit()
        return False

def launcher():
    try:
        if app:
            url = text_area.get("1.0", tk.END).strip()
        else:
            url = existing_url

        if not url:
            logging.error("No configured URL.")
            messagebox.showerror("Error", "No configured URL.")
            return

        msgError = check_url(url)
        if msgError:
            messagebox.showerror("Error", msgError)
            return

        status_var.set("Starting process...")
        app.update_idletasks()

        browsers = ["chrome", "edge"]
        for browser in browsers:
            try:
                print(f"Trying to automate with {browser}...")
                ok = automate(browser, url)
                if ok:
                    status_var.set("Ready.")
                    app.update_idletasks()
                    break
            except Exception as e:
                if browser == browsers[-1]:
                    raise e
                else:
                    continue

    except Exception as e:
        logging.error(f"Unexpected error in launcher: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        status_var.set("Unexpected error.")
        app.update_idletasks()

def check_url(urlIn=None):
    error = None
    if urlIn:
        url = urlIn
    else:
        url = text_area.get("1.0", tk.END).strip()

    pattern = re.compile(r'^(https?://)?(localhost|powerbislider\.com|www\.powerbislider\.com)(/.*)?$')
    if pattern.match(url):
        if launch_button:
            launch_button.config(state=tk.NORMAL, text="Test URL")
        if app:
            text_area.config(bg="white")
    else:
        if launch_button:
            launch_button.config(state=tk.DISABLED, text="Invalid URL")
        error = "Invalid URL!"
        if app:
            text_area.config(bg="misty rose")

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
    config_dir = _CONFIG_DIR
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)

    log_file = os.path.join(config_dir, 'error.log')
    config_file = os.path.join(config_dir, 'config.ini')

    logging.basicConfig(filename=log_file, level=logging.ERROR, format=_LOG_FORMAT)
    
    config = configparser.ConfigParser()
    if os.path.exists(config_file):
        config.read(config_file)
    existing_url = config['DEFAULT']['url'] if 'url' in config['DEFAULT'] else ''

    open_app = True
    launch_button = None
    
    # Cria a janela principal
    app = tk.Tk()
    app.title(f"Settings - PB Slider Launcher (v{_MANIFEST['version']})")

    # Dimensões iniciais
    width, height = 500, 280
    app.geometry(f"{width}x{height}")
    app.update_idletasks()

    # Centraliza
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    app.geometry(f"{width}x{height}+{x}+{y}")

    # **Configura a grid da janela principal**
    app.rowconfigure(0, weight=1)  # A linha 0 se expande
    app.rowconfigure(1, weight=0)  # A linha 1 não se expande (barra de status)
    app.columnconfigure(0, weight=1)

    app.resizable(True, True)

    icon_path = resource_path("logo.ico")
    if os.path.exists(icon_path):
        app.iconbitmap(icon_path)
    else:
        logging.warning(f"Icon path {icon_path} does not exist.")

    # Frame principal (conteúdo)
    main_frame = tk.Frame(app, padx=20, pady=20)
    main_frame.grid(row=0, column=0, sticky='nsew')  # Expande na linha 0

    main_frame.columnconfigure(0, weight=1)

    url_label = tk.Label(main_frame, text="Power BI Slider URL:")
    url_label.grid(row=0, column=0, sticky='w')

    text_area = tk.Text(main_frame, height=4)
    text_area.grid(row=1, column=0, sticky='we', pady=(5, 15))
    text_area.insert(tk.END, existing_url)
    text_area.bind("<KeyRelease>", lambda event: check_url())

    button_frame = tk.Frame(main_frame)
    button_frame.grid(row=2, column=0, sticky='we', pady=(0, 15))
    button_frame.columnconfigure(0, weight=1)

    launch_button = tk.Button(button_frame, text="Test URL", command=launcher, state=tk.DISABLED, width=20)
    launch_button.grid(row=0, column=0, sticky='w')

    startup_var = tk.BooleanVar()
    if is_in_startup():
        startup_var.set(True)

    startup_checkbox = tk.Checkbutton(button_frame, text="Start on System Startup", variable=startup_var)
    startup_checkbox.grid(row=0, column=1, sticky='e', padx=(10, 0))

    save_button = tk.Button(main_frame, text="Save", command=save_and_update_startup, height=2)
    save_button.grid(row=3, column=0, sticky='we')

    # Barra de status na linha 1 (fixa no rodapé)
    status_var = tk.StringVar(value="Ready.")
    status_bar = tk.Label(
        app,
        textvariable=status_var,
        bd=1,
        relief=tk.SUNKEN,
        anchor="e", # Alinhamento à direita
        padx=10, 
        font=("Arial", 8)
    )
    # Fica na linha 1, coluna 0, não se expande verticalmente.
    status_bar.grid(row=1, column=0, sticky="we")

    check_url()

    def on_close():
        app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_close)

    if len(sys.argv) > 1 and sys.argv[1] == "launch":
        launcher()

    app.mainloop()
