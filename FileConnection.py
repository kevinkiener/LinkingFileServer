import socketserver
from http.server import BaseHTTPRequestHandler
import subprocess
from http import HTTPStatus
import os
from urllib.parse import unquote

import win32gui
# import win32con

import win32com.client

import time
import sys

import pystray
from PIL import Image

import threading

import customtkinter as ctk

import json

import pyperclip as clip
import pyautogui

import win32con




# ctk.set_default_color_theme('dark-blue')

config_file_path = "config.json"

default_config = {"base_dir": "%OneDriveCommercial%", "server_port": 3030, "server_host": "http://localhost"}

relative_path = ''



if getattr(sys, 'frozen', False):
    path = os.path.dirname(sys.executable)
elif __file__:
    path = os.path.dirname(__file__)

log_file_path = os.path.join(path, "output.log")

# Open the log file in append mode
log_file = open(log_file_path, "a", buffering=1)

# Redirect stdout and stderr to the log file
sys.stdout = log_file
sys.stderr = log_file

config_file_path = os.path.join(path, config_file_path)

if not os.path.isfile(config_file_path):
    # If it doesn't exist, create it with the default configuration
    with open(config_file_path, "w") as config_file:
        json.dump(default_config, config_file, indent=4)
    print("Config file created with default settings.")

# Load the base_dir parameter from the config file
with open(config_file_path, "r") as config_file:
    config = json.load(config_file)
    relative_path = config.get("base_dir")
    default_config["base_dir"] = relative_path
    default_config["server_port"] = config.get("server_port")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, relative_path)

# dirname = os.path.dirname(__file__)
# TRAY_ICON = os.path.join(dirname, 'folder.png')

TRAY_ICON = resource_path('folder.png')
TRAY_ICON_ico = resource_path('folder.ico')



#for sending alt key for getting focus back
shell = win32com.client.Dispatch("WScript.Shell")

def bring_window_to_front(window_title):
    # Find the window handle by the window title (Explorer's title will be the folder name)

    hwnd = win32gui.FindWindow(None, window_title)
    print(hwnd)
    if hwnd:
        # If the window is minimized, restore it
        # win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # Bring it to the foreground
        # shell key needed to get focus back
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(hwnd)
    else:
        print("Window not found")

# def some_function():
#     #subprocess.Popen(r'explorer /select,"C:\path\of\folder\file"')
#     #os.startfile(r"C:\Users", show_cmd=1)
#     subprocess.Popen(['explorer', r'C:\Users'])

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def open_explorer(Path):
    # Path2 = Path.replace(r"%", " ")
    # Path2 = Path2[1:]
    # print(Path2)

    pyautogui.hotkey('ctrl', 'w')

    Path = unquote(Path)
    print(Path[1:])
    Path = relative_path + "/" + Path[1:]
    Path = Path.replace("/", "\\")
    print(Path)
    
    print(Path)
    expanded_path = os.path.expandvars(Path)
    print(expanded_path)
    shell.SendKeys('%')

    # if expanded_path.split('.')[-1] == 'pdf':
    #     subprocess.run(['open', expanded_path], check=True)
    # else:
    subprocess.Popen(['explorer', expanded_path])
    # subprocess.Popen(['explorer', Path2])

    time.sleep(1)

    # The window title will be the folder name, so extract the last part of the path
    window_title = os.path.basename(os.path.normpath(expanded_path))
    print(window_title)
    window_title = window_title + " - File Explorer"
    # Try to bring the window to the foreground
    # top_windows = []
    # win32gui.EnumWindows(windowEnumerationHandler, top_windows)
    # for i in top_windows:
    #     if "window_title" in i[1].lower():
    #         print(i)
    #         win32gui.ShowWindow(i[0],5)
    #         win32gui.SetForegroundWindow(i[0])
    #         break

    # bring_window_to_front(window_title)
    
class MyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        # if self.path == '/captureImage':
        #     # Insert your code here
        #     some_function()

        #self.send_response(200)
        open_explorer_flag = False
        print('Getting path : --------')
        print(self.path)
        # if self.path[:2] == r"/%":
        if not ('GET' in self.path) and not ('favicon' in self.path): 
            open_explorer_flag = True
        self.send_response(HTTPStatus.OK)
        self.end_headers()
        self.wfile.write(b'OpenExplorer')
        open_explorer(self.path)


httpd = socketserver.TCPServer(("", default_config['server_port']), MyHandler)
# os.startfile(r"C:\Users\kevin\OneDrive - TUM\Studium\PhD\Software\PythonServerFiles", show_cmd=1)


def create_image():
    return Image.open(TRAY_ICON)


def convert_clipboard_path():
    with open(config_file_path, "r") as config_file:
        config = json.load(config_file)
        port = config.get("server_port")
        server_host = config.get("server_host")
        global relative_path
        relative_path = config.get("base_dir")

    server_address = f"{server_host}:{port}"+'/' 
    # print(server_address)  
    expanded_path = os.path.expandvars(relative_path)
    # print(expanded_path)

    path_to_convert = clip.paste()
    print(path_to_convert)
    path_to_convert = path_to_convert.replace("\"", "")
    print(path_to_convert)
    # print(path_to_convert)
    try:
        path_to_convert = path_to_convert.split(expanded_path)[1][1:]
    except IndexError:
        print("relative path not in clipboard")
        return
    # print(path_to_convert)
    resulting_path = server_address + path_to_convert
    # print(resulting_path)
    resulting_path = resulting_path.replace("\\", "/")

    clip.copy(resulting_path)

def show_customtkinter_window_convert_path():
    def convert_path(event=None):
        with open(config_file_path, "r") as config_file:
            config = json.load(config_file)
            port = config.get("server_port")
            server_host = config.get("server_host")
            global relative_path
            relative_path = config.get("base_dir")

        server_address = f"{server_host}:{port}"+'/' 
        # print(server_address)  
        expanded_path = os.path.expandvars(relative_path)
        # print(expanded_path)

        path_to_convert = entry.get()
        path_to_convert = path_to_convert.replace("\"", "")
        # print(path_to_convert)
        path_to_convert = path_to_convert.split(expanded_path)[1][1:]
        # print(path_to_convert)
        resulting_path = server_address + path_to_convert
        # print(resulting_path)
        resulting_path = resulting_path.replace("\\", "/")
        # print(resulting_path)
        # output_label.configure(text=resulting_path)
        output_label.configure(state="normal")
        output_label.delete("0.0", "end")
        output_label.insert("0.0", resulting_path)
        output_label.configure(state="disabled")

        clip.copy(resulting_path)




    window = ctk.CTk()
    window.title("Convert Path")
    window.iconbitmap(TRAY_ICON_ico)
    # window.iconbitmap(create_image())


    label = ctk.CTkLabel(window, text="Path to convert:")
    label.pack(padx=20, pady=10)

    entry = ctk.CTkEntry(window, width=300)
    entry.insert(0, "")
    entry.bind("<Return>", convert_path)
    entry.pack(padx=20, pady=10)

    save_button = ctk.CTkButton(window, text="Convert", command=convert_path)
    save_button.pack(padx=20, pady=10)

    output_label = ctk.CTkTextbox(window, width=300, height =80)
    output_label.insert("0.0", "")
    output_label.configure(state="disabled")

    output_label.pack(padx=20, pady=10)

    window.mainloop()


def show_customtkinter_window_path():
    def save_path():
        global relative_path
        relative_path = entry.get()
        config["base_dir"] = relative_path
        with open(config_file_path, "w") as config_file:
            json.dump(config, config_file, indent=4)
        

        window.destroy()

    window = ctk.CTk()
    window.title("Edit Path")
    window.iconbitmap(TRAY_ICON_ico)
    # window.iconbitmap(create_image())


    label = ctk.CTkLabel(window, text="Relative Path:")
    label.pack(padx=20, pady=10)

    entry = ctk.CTkEntry(window, width=300)
    entry.insert(0, relative_path)
    entry.pack(padx=20, pady=10)

    save_button = ctk.CTkButton(window, text="Save", command=save_path)
    save_button.pack(padx=20, pady=10)

    window.mainloop()


def on_clicked(icon, item):
    print("Tray icon clicked")


def quit_program(icon):
    icon.stop()
    log_file.close()
    os._exit(0)

def setup_tray_icon():
    icon = pystray.Icon("File Connection")
    icon.icon = create_image()
    icon.title = "File Connection Server"
    icon.menu = pystray.Menu(
        # pystray.MenuItem("Open Explorer", lambda: open_explorer(r'C:\Users')),
        pystray.MenuItem("Convert Path", lambda: show_customtkinter_window_convert_path()),
        pystray.MenuItem("Convert Clipboard  Path", lambda: convert_clipboard_path()),
        pystray.MenuItem("Edit Path", lambda: show_customtkinter_window_path()),
        # pystray.MenuItem("Quit", lambda: icon.stop())
        pystray.MenuItem("Quit", lambda: quit_program(icon))
    )
    print("Setting up tray icon")
    icon.run()
    print("Tray icon setup complete")

def start_server():
    print("Starting server")
    httpd.serve_forever()

if __name__ == '__main__':
    tray_thread = threading.Thread(target=setup_tray_icon)
    server_thread = threading.Thread(target=start_server)

    tray_thread.start()
    server_thread.start()

    tray_thread.join()
    server_thread.join()
