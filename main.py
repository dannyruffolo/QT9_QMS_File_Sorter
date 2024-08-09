import os
import sys
import shutil
import time
import logging
import signal
import subprocess
import threading
import tkinter as tk
import pystray
import tempfile
import requests
import json
import queue
from pystray import MenuItem as item
from PIL import Image
from tkinter import messagebox, filedialog
from datetime import datetime
from logging.handlers import RotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from plyer import notification

CURRENT_VERSION = "3.0.3"
GITHUB_REPO = "dannyruffolo/QT9_QMS_File_Sorter"
TARGET_PATH_FILE = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "QT9Software", "QT9 QMS File Sorter", "target_path.json")
PREFERENCES_FILE = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "QT9Software", "QT9 QMS File Sorter", "preferences_file.json")

global user_preferences, target_path
default_target_path = os.path.expanduser(r"~\OneDrive - QT9 Software\Recordings")
target_path = default_target_pathuser_preferences = {}
user_preferences = {}
selected_folder = ""  # To store the last selected folder
gui_queue = queue.Queue()
root = tk.Tk()
root.withdraw()

log_format = '%(asctime)s - %(name)s - %(levelname)s - [Line #%(lineno)d] - %(message)s'
log_file_path = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "QT9Software", "QT9 QMS File Sorter")
if not os.path.exists(log_file_path):
    os.makedirs(log_file_path)
log_file = os.path.join(log_file_path, 'app.log')

# Set up RotatingFileHandler
log_handler = RotatingFileHandler(log_file, mode='a', maxBytes=5*1024*1024, backupCount=2, encoding=None, delay=0)
log_handler.setLevel(logging.INFO)
formatter = logging.Formatter(log_format, datefmt='%m-%d-%Y %H:%M:%S')
log_handler.setFormatter(formatter)

# Get the root logger and set the handler
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

keep_running = threading.Event()
tray_icon = None
tray_icon_lock = threading.Lock()

# Predefined preferences
default_preferences = {
    "QT9 QMS Audit Management": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Audit"],
    "QT9 QMS Calibrations": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Calibrations"],
    "QT9 QMS CAPA_NCP": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/CAPA"],
    "QT9 QMS Change Control": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Change Control"],
    "QT9 QMS Customer Module": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Customer Feedback - Surveys"],
    "QT9 QMS Deviations": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Deviations"],
    "QT9 QMS Doc Control": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Document Control"],
    "QT9 QMS ECR_ECN": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/ECR-ECN"],
    "QT9 QMS Inspections": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Inspections"],
    "QT9 QMS Preventive Maintenance": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Preventative Maintenance"],
    "QT9 QMS Supplier Surveys_Evaluations": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Supplier Surveys - Evaluations"],
    "QT9 QMS Training Module": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Training Module"],
    "QT9 QMS Product Design": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Product Design"],
    "QT9 QMS Quality Events": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/Quality Events"],
    "QT9 QMS Test Module": ["C:/Users/druffolo/Box/QT9 University/Training Recordings/TEST - DO NOT USE"],
}

def create_default_files():
    # Ensure the directory exists
    os.makedirs(os.path.dirname(TARGET_PATH_FILE), exist_ok=True)

    # Create target path file with default value if it doesn't exist
    if not os.path.exists(TARGET_PATH_FILE):
        with open(TARGET_PATH_FILE, 'w') as file:
            json.dump({'target_path': default_target_path}, file)

    # Create preferences file with default value if it doesn't exist
    if not os.path.exists(PREFERENCES_FILE):
        with open(PREFERENCES_FILE, 'w') as file:
            json.dump({}, file)

def populate_preferences():
    with open(PREFERENCES_FILE, 'w') as file:
        json.dump(default_preferences, file, indent=4)
    logging.info("Preferences populated successfully.")

    load_user_preferences()
    update_preferences_display()

def setup_system_tray():
    global tray_icon
    with tray_icon_lock:
        if tray_icon is not None:
            return  # Tray icon already created
        try:
            icon_image_path = r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico'
            icon_image = Image.open(icon_image_path)
        except FileNotFoundError:
            try:
                icon_image_path = r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico'
                icon_image = Image.open(icon_image_path)
            except FileNotFoundError as e:
                raise Exception("Both logo files could not be found.") from e

        menu = (item('Open Menu', lambda: gui_queue.put(show_main_gui)),
                item('Check for Updates', lambda: gui_queue.put(system_tray_check_for_updates)),
                item('Quit', quit_application))

        tray_icon = pystray.Icon("QT9 QMS File Sorter", icon_image, "QT9 QMS File Sorter", menu)
        tray_icon.run()

def create_main_gui():
    gui_queue.put(show_main_gui)

def process_queue():
    try:
        while True:
            command = gui_queue.get_nowait()
            command()
    except queue.Empty:
        pass
    root.after(100, process_queue)

def show_main_gui():

    main_window = tk.Toplevel(root)
    main_window.title("Menu")

    try:
        main_window.iconbitmap(r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico')
    except tk.TclError:
        try:
            main_window.iconbitmap(r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico')
        except tk.TclError as e:
            raise Exception("Both logo files could not be found.") from e

    # Window size
    window_width = 400
    window_height = 350

    # Get screen width and height
    screen_width = main_window.winfo_screenwidth()
    screen_height = main_window.winfo_screenheight()

    # Calculate x and y coordinates
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))

    # Set the window's position to the center of the screen
    main_window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    main_window.configure(bg='grey')
    label = tk.Label(main_window, text="QT9 QMS File Sorter", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 18))
    label.pack(pady=(20, 10))
    button_frame = tk.Frame(main_window, bg='grey')
    button_frame.place(relx=0.5, rely=0.55, anchor='center')
    open_config_btn = tk.Button(button_frame, text="Configuration Menu", command=config_gui, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    open_config_btn.grid(row=0, column=0, pady=13)  # Adjust row index as needed
    check_for_update_btn = tk.Button(button_frame, text="Check For Updates", command=lambda: check_for_updates(True), fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    check_for_update_btn.grid(row=2, column=0, pady=13)
    open_qt9_folder_btn = tk.Button(button_frame, text="Application Files", command=open_qt9_folder, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    open_qt9_folder_btn.grid(row=1, column=0, pady=13)
    move_to_startup_btn = tk.Button(button_frame, text="Run App on Startup", command=run_move_to_startup, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    move_to_startup_btn.grid(row=3, column=0, pady=13)

    process_queue()

def show_up_to_date_window():
    def close_window():
        window.destroy()
        window.quit()

    window = tk.Toplevel(root)
    window.title("Update Check")
    window_width = 232
    window_height = 114
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))
    window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
    window.resizable(False, False)

    # Disable minimize and maximize buttons
    window.attributes('-toolwindow', True)

    window.configure(bg='grey')

    # Use a Frame as a container for better control over the layout
    frame = tk.Frame(window, bg='grey')
    frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    label = tk.Label(frame, text="Your program is up to date.", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 10, 'bold'))
    label.place(relx=0.5, rely=0.3, anchor='center')
    
    # Pack the OK button at the bottom of the frame, centered
    ok_button = tk.Button(frame, text="OK", command=close_window, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 10, 'bold'), width=8)
    ok_button.pack(side=tk.BOTTOM, pady=(5, 0))

    ok_button.place(anchor='se', relx=0.98, rely=1.0)
    window.mainloop()

def check_for_updates(show_up_to_date=False):
    try:
        response = requests.get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest")
        response.raise_for_status()  # Ensure the request was successful
        latest_version = response.json()['tag_name']
        latest_version_tuple = tuple(map(int, latest_version.strip('v').split('.')))
        current_version_tuple = tuple(map(int, CURRENT_VERSION.strip('v').split('.')))
        if latest_version_tuple > current_version_tuple:
            show_update_gui(latest_version)
        elif show_up_to_date:
            show_up_to_date_window()
    except requests.RequestException as e:
        logging.error(f"Error checking for updates: {e}")

def system_tray_check_for_updates():
    check_for_updates(show_up_to_date=True)

def show_update_gui(latest_version):
    global update_window
    def on_install():
        download_and_install_update(latest_version)
        update_window.destroy()

    def on_decline():
        update_window.destroy()

    update_window = tk.Toplevel(root)
    update_window.attributes('-topmost', True)
    update_window.resizable(False, False)

    try:
        update_window.iconbitmap(r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico')
    except tk.TclError:
        try:
            update_window.iconbitmap(r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico')
        except tk.TclError as e:
            logging.error("Both logo files could not be found.")

    # Window size and position
    window_width = 300
    window_height = 150
    screen_width = update_window.winfo_screenwidth()
    screen_height = update_window.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))
    update_window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    update_window.configure(bg='grey')
    update_window.title("Update Available")
    
    label = tk.Label(update_window, text=f"Version {latest_version} is available.\nDo you want to install the update?", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 10, 'bold'))
    label.place(relx=0.5, rely=0.3, anchor='center')
    
    button_frame = tk.Frame(update_window, bg='grey')
    button_frame.place(relx=0.5, rely=0.6, anchor='center')
    
    install_btn = tk.Button(button_frame, text="Install", command=on_install, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 10, 'bold'), width=8)
    install_btn.grid(row=0, column=0, padx=10)
    
    later_btn = tk.Button(button_frame, text="Later", command=on_decline, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 10, 'bold'), width=8)
    later_btn.grid(row=0, column=1, padx=10)
    
    update_window.mainloop()

def uninstall_old_version():
    try:
        uninstaller_path = r'C:\Program Files\QT9 QMS File Sorter\unins000.exe'
        if os.path.exists(uninstaller_path):
            subprocess.run([uninstaller_path, '/SILENT'], check=True)  # Run with silent flag
            logging.info("Old version uninstalled successfully.")
        else:
            logging.info("Uninstaller not found. Proceeding with the installation of the new version.")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error during uninstallation: {e}")
        sys.exit(1)  # Exit if uninstallation fails to prevent potential conflicts

def download_and_install_update(latest_version):
    try:
        download_url = f"https://github.com/{GITHUB_REPO}/releases/download/{latest_version}/FileSorter_{latest_version}_Installer.exe"
        response = requests.get(download_url)
        temp_dir = tempfile.mkdtemp()
        installer_path = os.path.join(temp_dir, "FileSorter_Installer.exe")
        with open(installer_path, 'wb') as file:
            file.write(response.content)
        messagebox.showinfo("Update", "Download completed. Starting the installer.", parent=update_window)
        os.startfile(installer_path)
        uninstall_old_version()
        os._exit(0)  # Exit the current process to allow the installer to run
    except Exception as e:
        messagebox.showerror("Update Error", f"Failed to download and install update: {e}", parent=update_window)

def periodic_check():
    check_for_updates()
    threading.Timer(3600, periodic_check).start()  # Check for updates every hour

def select_target_path():
    global target_path
    selected_path = filedialog.askdirectory()
    if selected_path:
        target_path = selected_path
        save_target_path(target_path)
        restart_observer(target_path)
        target_path_label.config(text=target_path)  # Update the label with the new path
        logging.info(f"Target path updated to: {target_path}")
    else:
        logging.info("No path selected.")

def save_target_path(path):
    with open(TARGET_PATH_FILE, 'w') as file:
        json.dump({"target_path": path}, file)

def load_target_path():
    global target_path
    if os.path.exists(TARGET_PATH_FILE):
        with open(TARGET_PATH_FILE, 'r') as file:
            data = json.load(file)
            target_path = data.get("target_path", default_target_path)
    else:
        target_path = default_target_path

def select_destination_folder():
    global selected_folder  # Declare as global to ensure it can be accessed
    selected_folder = filedialog.askdirectory()
    if selected_folder:  # Check if a folder was selected
        destination_folder_label.config(text=selected_folder)  # Update the label with the selected folder path
    else:
        destination_folder_label.config(text="No folder selected")

def add_to_preferences():
    global selected_folder
    folder_path = destination_folder_label.cget("text")
    file_name = file_name_entry.get()
    if folder_path == "No folder selected" or not file_name:
        messagebox.showwarning("Warning", "Please select a folder and enter a file name first.")
        return

    file_name = file_name_entry.get()
    if file_name and selected_folder:  # Ensure there's a file name entered and a folder selected
        # Initialize the list for this file name if it doesn't exist
        if file_name not in user_preferences:
            user_preferences[file_name] = []
        # Append the selected folder to the list for this file name
        if selected_folder not in user_preferences[file_name]:
            user_preferences[file_name].append(selected_folder)
            update_preferences_display()  # Update the preferences display
        file_name_entry.delete(0, tk.END)  # Clear the file name entry field
        selected_folder = ""  # Reset the selected folder variable
        destination_folder_label.config(text="No folder selected")  # Reset the label
    else:
        logging.info("Please select a folder and enter a file name first.")

def save_user_preferences(user_preferences):
    # Update user_preferences with the current entries
    updated_preferences = {}
    for frame in preferences_display_frame.winfo_children():
        file_name = frame.children['!entry'].get()
        folder_path = frame.children['!label'].cget("text")
        if file_name and folder_path:
            updated_preferences[file_name] = [folder_path]

    # Replace the old preferences with the updated ones
    user_preferences.clear()
    user_preferences.update(updated_preferences)

    with open(PREFERENCES_FILE, 'w') as file:
        json.dump(user_preferences, file, indent=4)
    logging.info("Preferences saved successfully.")
    update_preferences_display()

def update_preferences_display():
    global preferences_display_frame, user_preferences
    for widget in preferences_display_frame.winfo_children():
        widget.destroy()

    for file_name, folders in user_preferences.items():
        frame = tk.Frame(preferences_display_frame, bg='grey')
        frame.pack(fill=tk.X, pady=2)

        file_name_entry = tk.Entry(frame, width=35, font=('Segoe UI Variable', 13))
        file_name_entry.insert(0, file_name)
        file_name_entry.pack(side=tk.LEFT, padx=15)

        folders_label = tk.Label(frame, text=", ".join(folders), bg='white', fg='black', font=('Segoe UI Variable', 13), width=80, anchor='w')
        folders_label.pack(side=tk.LEFT, padx=5, fill=tk.X)

        delete_button = tk.Button(frame, text="Delete", command=lambda fn=file_name: delete_preference(fn), fg='#ffffff', bg='#ff0000', font=('Segoe UI Variable', 13))
        delete_button.pack(side=tk.RIGHT, padx=5)

def delete_preference(file_name):
    if file_name in user_preferences:
        del user_preferences[file_name]
        update_preferences_display()

def load_user_preferences():
    global user_preferences
    if os.path.exists(PREFERENCES_FILE):
        with open(PREFERENCES_FILE, 'r') as file:
            user_preferences = json.load(file)
    else:
        user_preferences = default_preferences

def clear_all_preferences():
    global user_preferences
    user_preferences.clear()  # Clear the dictionary
    with open(PREFERENCES_FILE, 'w') as file:
        json.dump(user_preferences, file)
    update_preferences_display()
    messagebox.showinfo("Clear All", "All preferences have been cleared.")

def config_gui():
    global file_name_entry, destination_folder_label, preferences_display_frame, target_path_label
    config = tk.Toplevel(root)
    config.title("File Sorter Configuration")

    # Attempt to set the window icon with fallback paths
    try:
        config.iconbitmap(r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico')
    except tk.TclError:
        try:
            config.iconbitmap(r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico')
        except tk.TclError as e:
            raise Exception("Both logo files could not be found.") from e

    # Window size and position
    window_width = 1185
    window_height = 775
    screen_width = config.winfo_screenwidth()
    screen_height = config.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))
    config.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    config.configure(bg='grey')
    label = tk.Label(config, text="QT9 QMS File Sorter Configuration", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 18))
    label.pack(pady=(20, 10))

    # Frame for target path selection
    target_path_frame = tk.Frame(config, bg='grey')
    target_path_frame.pack(pady=(0, 20))

    tk.Label(target_path_frame, text="Target Path:", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 13)).pack(side=tk.LEFT, padx=(0, 5))
    target_path_label = tk.Label(target_path_frame, text=target_path, bg='white', fg='black', font=('Segoe UI Variable', 13))
    target_path_label.pack(side=tk.LEFT, padx=(0, 5))
    select_target_path_button = tk.Button(target_path_frame, text="Select Target Path", command=select_target_path, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 13))
    select_target_path_button.pack(side=tk.LEFT, padx=(5, 0))

    # Frame for buttons and inputs
    input_frame = tk.Frame(config, bg='grey')
    input_frame.pack(pady=(0, 20))

    tk.Label(input_frame, text="File Name Contains:", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 13)).pack(side=tk.LEFT, padx=(0, 5))
    file_name_entry = tk.Entry(input_frame, width=35, font=('Segoe UI Variable', 13))
    file_name_entry.pack(side=tk.LEFT)

    # Modify the button frame to include the destination folder label
    button_frame = tk.Frame(config, bg='grey')
    button_frame.pack(pady=(0, 10))
    
    select_folder_button = tk.Button(button_frame, text="Select Folder", command=select_destination_folder, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 13))
    select_folder_button.pack(side=tk.LEFT, padx=10)
    
    destination_folder_label = tk.Label(button_frame, text="No folder selected", bg='white', fg='black', font=('Segoe UI Variable', 13))
    destination_folder_label.pack(side=tk.LEFT, padx=10)
    
    # Frame to hold the add and populate buttons
    add_populate_frame = tk.Frame(config, bg='grey')
    add_populate_frame.pack(pady=(10, 10))

    # Initialize and pack the add_button and populate_preferences_btn in the same frame
    add_button = tk.Button(add_populate_frame, text="Add to Preferences", command=add_to_preferences, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 13))
    add_button.pack(side=tk.LEFT, padx=5)
    
    populate_preferences_btn = tk.Button(add_populate_frame, text="Add QMS Defaults", command=populate_preferences, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 13))
    populate_preferences_btn.pack(side=tk.LEFT, padx=5)
    
    # Create a frame to hold the canvas and scrollbar
    display_frame = tk.Frame(config, bg='grey')
    display_frame.pack(fill=tk.BOTH, expand=True)
    
    # Create a canvas and a scrollbar for the preferences_display_frame
    canvas = tk.Canvas(display_frame, bg='grey', bd=0, highlightthickness=0)
    scrollbar = tk.Scrollbar(display_frame, orient="vertical", command=canvas.yview)
    preferences_display_frame = tk.Frame(canvas, bg='grey')
    
    # Configure the canvas and scrollbar
    preferences_display_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    canvas.create_window((0, 0), window=preferences_display_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Pack the canvas and scrollbar
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Bind the mouse wheel event to the canvas
    def on_mouse_wheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<MouseWheel>", on_mouse_wheel)
    
    update_preferences_display()  # Update the preferences display with the current preferences
    
    # Pack the save button at the bottom of the window, centered
    save_button = tk.Button(config, text="Save Preferences", command=lambda: save_user_preferences(user_preferences), fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 13))
    save_button.pack(side=tk.BOTTOM, pady=(10, 10))
    
    # Add the clear all button at the bottom of the window, centered
    clear_all_button = tk.Button(config, text="Clear All", command=clear_all_preferences, fg='#ffffff', bg='#ff0000', font=('Segoe UI Variable', 13))
    clear_all_button.pack(side=tk.BOTTOM, pady=(10, 10))
    
    # Load and display preferences when the config window is created
    load_user_preferences()
    update_preferences_display()
    
    config.mainloop()

def move_files():
    global selected_folder, user_preferences
    logging.info('Starting move_files function')

    global user_preferences  # Assuming user_preferences is defined globally

    # Check if user_preferences is a string and convert it to a dictionary if so
    if isinstance(user_preferences, str):
        try:
            user_preferences_dict = json.loads(user_preferences)
        except json.JSONDecodeError:
            logging.error("Error decoding JSON from user_preferences")
            return
    else:
        user_preferences_dict = user_preferences

    for filename in os.listdir(target_path):
        logging.info(f'Processing file: {filename}')
        new_filename = None

        # Iterate over the dictionary of user preferences
        for preference, folder in user_preferences_dict.items():
            if preference in filename:
                logging.info(f'File {filename} matches core file name {preference} and will be processed')
                _, file_extension = os.path.splitext(filename)
                new_filename = f"{preference} {datetime.now().strftime('%m-%d-%Y')}{file_extension}"
                destination_folder = folder
                break

        if new_filename and destination_folder:
            source_file_path = os.path.join(target_path, filename)
            destination_folder_path = destination_folder[0] if isinstance(destination_folder, list) else destination_folder
            destination_file_path = os.path.join(destination_folder_path, new_filename)  # Corrected to use destination_folder_path
            if os.path.exists(destination_file_path):
                logging.info(f'The file {os.path.basename(source_file_path)} already exists in the destination folder. Skipping this file.')
            else:
                try:
                    time.sleep(1)  # Wait for 1 second
                    shutil.move(source_file_path, destination_file_path)
                    logging.info(f'The file {os.path.basename(source_file_path)} has been moved successfully.')
                    # Assuming send_notification is defined elsewhere
                    send_notification(os.path.basename(source_file_path), new_filename, os.path.basename(destination_folder_path))
                except PermissionError as e:
                    logging.error(f'Permission denied for {source_file_path} ({filename}). Error: {e}')
                except IOError as e:
                    logging.error(f'IOError encountered for {source_file_path} ({filename}). Error: {e}')
                except Exception as e:
                    logging.error(f'Unexpected error moving {source_file_path} ({filename}). Error: {e}')
    logging.info('Completed move_files function')

def send_notification(original_file_name, new_file_name, destination_folder):
    try:
        logging.info(f'Attempting to send notification for {new_file_name}')
        notification.notify(
            title='QT9 U Recording Transfer',
            message=f'The file "{original_file_name}" has been renamed to "{new_file_name}" and moved to \\Training Recordings\\{destination_folder}.',
            timeout=1  # Timeout in seconds
        )
    except Exception as e:
        logging.error(f'An error occurred while trying to send a notification: {e}')

class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        logging.info(f'The file {os.path.basename(event.src_path)} has been created!')
        move_files()

def run_move_to_startup():
    try:
        source_path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
        username = os.getlogin()
        destination_path = fr"C:\Users\{username}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        shortcut_name = "QT9 QMS File Sorter.lnk"
        source_shortcut_path = os.path.join(source_path, shortcut_name)
        os.makedirs(destination_path, exist_ok=True)
        if os.path.exists(source_shortcut_path):
            shutil.copy2(source_shortcut_path, destination_path)
            messagebox.showinfo("Success", f"Application setup to run on startup.")
        else:
            messagebox.showerror("Error", f"Shortcut '{shortcut_name}' not found in the source directory.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to move shortcut: {str(e)}")

def open_qt9_folder():
    try:
        os.startfile(log_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open folder: {str(e)}")

def quit_application(icon, item):
    icon.stop()
    global keep_running
    keep_running = False
    os._exit(0)

def signal_handler(signum, frame):
    logging.info(f'Signal {signum} received, stopping observer.')
    keep_running.clear()

observer = Observer()
event_handler = MyHandler()

def start_observer():
    global observer
    
    # Check if the path exists
    if not os.path.exists(target_path):
        logging.error(f"Error: The path {target_path} does not exist.")
        return

    observer.schedule(event_handler, path=target_path, recursive=False)
    observer.start()

    try:
        printed_message = False
        while keep_running.is_set():
            if not printed_message:
                logging.info('Observer is running')
                printed_message = True
            time.sleep(1)
    except Exception as e:
        logging.error(f'Unexpected error in main loop. Error: {e}')
    finally:
        logging.info('Stopping the observer')
        observer.stop()
        observer.join()
        logging.info('Observer has been successfully stopped')

def restart_observer(new_path):
    global observer
    observer.stop()
    observer.join()
    observer = Observer()
    observer.schedule(event_handler, path=new_path, recursive=False)
    observer.start()
    logging.info(f"Observer restarted with new path: {new_path}")

def main():
    create_default_files()
    load_target_path()
    load_user_preferences()
    
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    keep_running.set()
    logging.info('Set up signal handlers')

    root.after(100, process_queue)
    logging.info('Started main GUI')
    
    # Start the observer in a separate thread
    observer_thread = threading.Thread(target=start_observer)
    observer_thread.start()
    logging.info('Started observer thread')

    # Set up the system tray icon
    tray_thread = threading.Thread(target=setup_system_tray)
    tray_thread.start()
    logging.info('Set up the system tray icon')

    # Start periodic update checks
    periodic_check()
    logging.info('Started periodic update checks')

    # Show the update window first
    logging.info('Showing update window')
    check_for_updates(show_up_to_date=True)

    # Show the main menu GUI
    create_main_gui()
    logging.info('Created main menu GUI')

    root.mainloop()
    logging.info('Entered main loop')

    # Signal the observer thread to stop
    logging.info('Signaling observer thread to stop')
    keep_running.clear()
    observer_thread.join()
    logging.info('Observer thread has been successfully stopped')

if __name__ == "__main__":
    main()