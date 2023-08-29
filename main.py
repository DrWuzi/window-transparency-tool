import io
import os
import sys
import time
import json
import psutil
import base64
import pystray
import win32gui
import threading
import win32process

import tkinter as tk
from PIL import Image
from tkinter import ttk
from win32com.client import Dispatch

# Constants
APP_CONFIG_PATH = 'app_transparency_config.json'
DEFAULT_TRANSPARENCY = 255  # 0-255
APPLY_TRANSPARENCY_SETTINGS_INTERVAL = 1  # In seconds
PENGUIN_ICON_BASE64 = "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"


class WindowTransparencyApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Window Transparency Tool")

        self.active_icon = None
        self.load_transparency_config()

        self.create_gui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_window)
        self.on_close_window()

    def load_transparency_config(self):
        try:
            with open(APP_CONFIG_PATH, 'r') as file:
                self.app_transparency_config = json.load(file)
        except FileNotFoundError:
            self.app_transparency_config = {}
        except Exception as e:
            print("Error loading transparency config:", e)
            self.app_transparency_config = {}

    def save_transparency_config(self):
        with open(APP_CONFIG_PATH, 'w') as file:
            json.dump(self.app_transparency_config, file)

    def create_gui(self):
        def update_gui_list():
            app_listbox.delete(*app_listbox.get_children())

            sorted_entries = sorted(self.app_transparency_config.items(), key=lambda x: x[0].lower())

            for app_process_name, transparency in sorted_entries:
                app_listbox.insert("", "end", text=app_process_name, values=transparency)

        def on_add_button_click():
            process_name = new_process_entry.get()
            transparency = int(new_transparency_entry.get())
            if process_name:
                self.app_transparency_config[process_name] = transparency
                update_gui_list()
                self.save_transparency_config()

        def on_treeview_select(event):
            selected_item = app_listbox.selection()
            if selected_item:
                item = app_listbox.item(selected_item)
                process_name = item['text']
                transparency = item['values'][0]  # Values are stored as a list
                new_process_entry.delete(0, tk.END)
                new_process_entry.insert(0, process_name)
                new_transparency_entry.delete(0, tk.END)
                new_transparency_entry.insert(0, transparency)

        def get_visible_processes_without_tray():
            visible_processes = set()

            def enum_windows_callback(hwnd, _):
                try:
                    window_pid = win32process.GetWindowThreadProcessId(hwnd)[1]
                    process = psutil.Process(window_pid)
                    process_name = process.name().lower()
                    if process_name.endswith(".exe"):
                        process_name = process_name[:-4] if process_name.endswith(".exe") else process_name
                        if process_name not in self.app_transparency_config:
                            visible_processes.add(process_name)
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass

            win32gui.EnumWindows(enum_windows_callback, None)
            return visible_processes

        def update_processes_list():
            visible_processes_list.delete(0, tk.END)
            visible_processes = get_visible_processes_without_tray()
            visible_processes = sorted(visible_processes)
            for process_name in visible_processes:
                visible_processes_list.insert(tk.END, process_name)

        def refresh_list():
            update_processes_list()

        def on_right_list_select(event):
            selected_item = visible_processes_list.curselection()
            if selected_item:
                process_name = visible_processes_list.get(selected_item[0])
                new_process_entry.delete(0, tk.END)
                new_process_entry.insert(0, process_name)
                new_transparency_entry.delete(0, tk.END)
                new_transparency_entry.insert(0, "255")  # Default transparency value

        def on_app_listbox_key(event):
            selected_item = app_listbox.selection()
            new_process_entry.delete(0, tk.END)
            new_transparency_entry.delete(0, tk.END)
            if selected_item and event.keysym == "Delete":
                item = app_listbox.item(selected_item)
                process_name = item['text']

                # Delete the selected process from the configuration
                if process_name in self.app_transparency_config:
                    self.make_windows_transparent(f"{process_name}.exe", 255)
                    del self.app_transparency_config[process_name]
                    update_gui_list()
                    self.save_transparency_config()

        def create_startup_shortcut():
            app_executable_path = sys.executable
            startup_folder = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")

            shortcut_path = os.path.join(startup_folder, "WindowTransparencyTool.lnk")

            if not os.path.exists(shortcut_path):
                shell = Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = app_executable_path
                shortcut.WorkingDirectory = os.path.dirname(app_executable_path)
                shortcut.IconLocation = app_executable_path
                shortcut.save()

        def remove_startup_shortcut():
            startup_folder = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
            shortcut_path = os.path.join(startup_folder, "WindowTransparencyTool.lnk")

            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)

        def toggle_startup_shortcut():
            if not autostart_var.get():
                create_startup_shortcut()
                autostart_var.set(True)
            else:
                remove_startup_shortcut()
                autostart_var.set(False)

        frame = ttk.Frame(self.root, padding=10)
        frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        app_listbox = ttk.Treeview(frame, columns=("Transparency",))
        app_listbox.heading("#0", text="Process Name")  # Hier ändern wir die Spaltennummer
        app_listbox.heading("#1", text="Transparency")
        app_listbox.column("#0", width=150)  # Breite der Spalte "Process Name"
        app_listbox.column("#1", width=100)  # Breite der Spalte "Transparency"
        app_listbox.grid(row=0, column=0, columnspan=2, pady=10)

        app_listbox.heading("#0", text="Process Name")
        app_listbox.heading("#1", text="Transparency")

        new_process_label = ttk.Label(frame, text="New Process Name:")
        new_process_label.grid(row=1, column=0, sticky=tk.W)

        new_process_entry = ttk.Entry(frame)
        new_process_entry.grid(row=1, column=1)

        new_transparency_label = ttk.Label(frame, text="Transparency (0-255):")
        new_transparency_label.grid(row=2, column=0, sticky=tk.W)

        new_transparency_entry = ttk.Entry(frame)
        new_transparency_entry.grid(row=2, column=1)

        add_button = ttk.Button(frame, text="Add Process", command=on_add_button_click)
        add_button.grid(row=3, column=0, columnspan=2, pady=10)

        autostart_var = tk.BooleanVar()
        autostart_var.set(False)

        autostart_label = ttk.Label(frame, text="Autostart")
        autostart_label.grid(row=4, column=0, columnspan=2, pady=0)

        add_button = ttk.Button(frame, text="Add", command=toggle_startup_shortcut)
        add_button.grid(row=5, column=0)

        remove_button = ttk.Button(frame, text="Remove", command=toggle_startup_shortcut)
        remove_button.grid(row=5, column=1)

        app_listbox.bind("<<TreeviewSelect>>", on_treeview_select)
        app_listbox.bind("<Delete>", on_app_listbox_key)

        visible_processes_list = tk.Listbox(frame)
        visible_processes_list.grid(row=0, column=2, rowspan=5, padx=10, pady=10, sticky=(tk.N, tk.S))
        visible_processes_list.bind("<<ListboxSelect>>", on_right_list_select)

        refresh_button = ttk.Button(frame, text="Refresh", command=refresh_list)
        refresh_button.grid(row=5, column=2, padx=10, sticky=(tk.W, tk.E))

        # Berechne die Bildschirmgröße
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Berechne die Fenstergröße
        window_width = 405
        window_height = 400

        # Berechne die Position, um das Fenster in der Mitte zu platzieren
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2

        # Setze die Fensterposition
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.root.resizable(False, False)

        update_gui_list()
        refresh_list()

    def set_window_transparency(self, hwnd, transparency):
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x00080000
        LWA_ALPHA = 0x00000002

        current_exstyle = win32gui.GetWindowLong(hwnd, GWL_EXSTYLE)
        new_exstyle = current_exstyle | WS_EX_LAYERED
        win32gui.SetWindowLong(hwnd, GWL_EXSTYLE, new_exstyle)

        win32gui.SetLayeredWindowAttributes(hwnd, 0, transparency, LWA_ALPHA)

    def make_windows_transparent(self, process_name, transparency):
        def enum_windows_callback(hwnd, _):
            try:
                window_pid = win32process.GetWindowThreadProcessId(hwnd)[1]
                process = psutil.Process(window_pid)
                if process.name().lower() == process_name.lower():
                    self.set_window_transparency(hwnd, transparency)
            except psutil.NoSuchProcess:
                pass

        win32gui.EnumWindows(enum_windows_callback, None)

    def apply_transparency_settings(self):
        while True:
            try:
                for app_process_name, transparency in self.app_transparency_config.items():
                    formatted_process_name = f"{app_process_name}.exe"
                    self.make_windows_transparent(formatted_process_name, transparency)
                time.sleep(APPLY_TRANSPARENCY_SETTINGS_INTERVAL)
            except Exception as e:
                print("An error occurred:", e)

    def on_close_window(self):
        if self.active_icon:
            self.active_icon.stop()
        self.root.withdraw()  # Hide the main window

        self.root.withdraw()  # Verstecke das Hauptfenster
        menu = (
            pystray.MenuItem("Open", lambda: self.root.deiconify()),
            pystray.MenuItem("Exit", lambda: self.root.destroy())
        )

        icon = pystray.Icon(
            'Transparency Tool',
            icon=Image.open(io.BytesIO(base64.b64decode(PENGUIN_ICON_BASE64))),
            menu=menu)

        self.active_icon = icon

        def icon_runner():
            icon.run()

        threading.Thread(target=icon_runner).start()

    def start(self):
        transparency_thread = threading.Thread(target=self.apply_transparency_settings)
        transparency_thread.start()
        self.root.mainloop()


if __name__ == "__main__":
    app = WindowTransparencyApp()
    app.start()
