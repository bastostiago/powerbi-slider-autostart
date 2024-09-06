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

# Função para verificar se o atalho já existe no startup
def is_in_startup():
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    shortcut_path = os.path.join(startup_folder, f'{_MANIFEST["app_name"]}.lnk')
    return os.path.exists(shortcut_path)

# Função para adicionar o executável principal ao startup
def add_to_startup():
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    exe_path = os.path.join(os.getcwd(), f'{_MANIFEST["app_name"]}.exe')  # Altere para o nome do seu executável principal
    shortcut_path = os.path.join(startup_folder, f'{_MANIFEST["app_name"]}.lnk')

    # Usar pywin32 para criar o atalho com o argumento "noopenapp"
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = exe_path
    shortcut.Arguments = "noopenapp"  # Aqui adiciona o argumento ao atalho
    shortcut.WorkingDirectory = os.getcwd()
    shortcut.IconLocation = exe_path  # Opcional: Definir o ícone
    shortcut.save()

# Função para remover o executável principal do startup
def remove_from_startup():
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    shortcut_path = os.path.join(startup_folder, f'{_MANIFEST["app_name"]}.lnk')

    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

# Função para salvar o config.ini e atualizar o startup
def save_and_update_startup():
    # Salvar o arquivo config.ini no diretório apropriado
    config_value = text_area.get("1.0", tk.END).strip()    
    with open(config_file, 'w') as config_file_object:
        config_file_object.write(f'[DEFAULT]\nurl={config_value}')

    # Atualizar o executável no startup com base no estado do checkbox
    if startup_var.get():
        add_to_startup()
    else:
        remove_from_startup()

    messagebox.showinfo("Success", "Data has been saved!")

# Função para abrir o Power BI Slider no navegador
def launcher():
    # Abre a URL no navegador padrão
    if app:
        url = text_area.get("1.0", tk.END).strip()
    else:
        url = existing_url

    if not url:
        logging.error("No configured URL.")
        return
    webbrowser.open(f"{url.split('?')[0]}?autoStart=1")

    # Espera alguns segundos para o navegador carregar
    time.sleep(7)

    # Obtém a lista de janelas abertas que contêm o título do navegador (por exemplo, "Chrome" ou "Firefox")
    browser_window = None
    for window in gw.getAllTitles():
        if "Chrome" in window or "Firefox" in window or "Edge" in window:  # Ajuste para o nome do navegador
            browser_window = window
            break

    # Se a janela do navegador foi encontrada, ativa e envia o comando F11
    if browser_window:
        window = gw.getWindowsWithTitle(browser_window)[0]
        window.activate()  # Ativa a janela do navegador
        time.sleep(1)  # Espera para garantir que a janela foi ativada
        pyautogui.press('f11')  # Envia o comando para entrar em tela cheia
    else:
        logging.error("Browser window not found.")

# Função para verificar se a URL é válida e habilitar/desabilitar o botão de launcher
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

# Função para localizar o ícone corretamente, tanto ao rodar como script quanto ao rodar como executável
def resource_path(relative_path):
    """ Get absolute path to resource, works for PyInstaller """
    try:
        # Quando o aplicativo é empacotado com PyInstaller, os arquivos temporários ficam no _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Quando o script é executado normalmente, os arquivos estão no local original
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

    # Configurar o logging
    logging.basicConfig(filename=log_file, level="ERROR", format=_LOG_FORMAT)

    # Ler o arquivo config.ini
    config = configparser.ConfigParser()
    if os.path.exists(config_file):
        config.read(config_file)
    existing_url = config['DEFAULT']['url'] if 'url' in config['DEFAULT'] else ''

    open_app = True
    launch_button = None
    print(sys.argv)
    if len(sys.argv) > 1 and sys.argv[1] == "noopenapp":
        msgError = check_url(existing_url)
        if not msgError:
            launcher()
        else:
            logging.error(msgError)
    else:        

        # Cria a interface gráfica
        app = tk.Tk()
        app.title("PB Slider Launcher - Settings")

        # Define o tamanho da janela e a centraliza na tela
        app.geometry('400x265')  # Largura x Altura
        app.eval('tk::PlaceWindow . center')  # Centraliza a janela na tela
        app.resizable(False, False)

        icon_path = resource_path("logo.ico")
        app.iconbitmap(icon_path)

        # Label para instruções
        tk.Label(app, text="Power BI Slider URL:").pack(pady=5)

        # Text area para inserir a URL
        text_area = tk.Text(app, height=6, width=40)
        text_area.pack(pady=(0, 10))

        # Preencher o Text area com a URL existente, se houver
        text_area.insert(tk.END, existing_url)

        # Função de callback para verificar a URL sempre que o conteúdo do text area for alterado
        text_area.bind("<KeyRelease>", lambda event: check_url())

        # Botão para lançar o Power BI Slider, alinhado à direita
        launch_button = tk.Button(app, text="Test URL", command=launcher, state=tk.DISABLED, width=10)
        launch_button.pack(padx=37, anchor='e')

        # Checkbox para gerenciar o startup
        startup_var = tk.BooleanVar()
        
        # Verificar se o atalho já está no startup e ajustar a flag
        if is_in_startup():
            startup_var.set(True)
        
        startup_checkbox = tk.Checkbutton(app, text="Start on System Startup", variable=startup_var)
        startup_checkbox.pack(pady=5)

        # Botão para salvar a configuração e atualizar o startup
        tk.Button(app, text="Save", command=save_and_update_startup, width=35, height=2).pack(pady=5)

        # Verifica a URL na inicialização para habilitar ou desabilitar o botão
        check_url()

        app.mainloop()
