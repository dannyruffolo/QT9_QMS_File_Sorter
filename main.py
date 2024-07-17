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
from pystray import MenuItem as item
from PIL import Image, ImageTk
from tkinter import messagebox
from tkinter import ttk
from datetime import datetime
from logging.handlers import RotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from plyer import notification

# Define CURRENT_VERSION and GITHUB_REPO constants
CURRENT_VERSION = "2.2.4"
GITHUB_REPO = "dannyruffolo/QT9_QMS_File_Sorter"

# Configuration
recordings_path = os.path.expanduser(r"~\OneDrive - QT9 Software\Recordings")
log_format = '%(asctime)s - %(name)s - %(levelname)s - [Line #%(lineno)d] - %(message)s'

# Updated path for log_file to store it in the specified folder
log_file_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "QT9 QMS File Sorter")
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

def setup_system_tray():
    try:
        icon_image_path = r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico'
        icon_image = Image.open(icon_image_path)
    except FileNotFoundError:
        try:
            icon_image_path = r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico'
            icon_image = Image.open(icon_image_path)
        except FileNotFoundError as e:
            raise Exception("Both logo files could not be found.") from e

    # Menu items
    menu = (item('Open Setup Wizard', show_main_gui),
            item('Check for Updates', check_for_updates), 
            item('Quit', quit_application),
            )

    # Create the system tray icon without running it immediately
    icon = pystray.Icon("QT9 QMS File Sorter", icon_image, "QT9 QMS File Sorter", menu)

    # Define a function to run the icon's event loop in a separate thread
    def run_icon():
        icon.run()

    # Create and start the thread
    icon_thread = threading.Thread(target=run_icon)
    icon_thread.daemon = True  # Daemonize thread to close it when the main program exits
    icon_thread.start()

def main_gui():
    root = tk.Tk()
    root.title("Setup Wizard")
    
    try:
        root.iconbitmap(r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico')
    except tk.TclError:
        try:
            root.iconbitmap(r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico')
        except tk.TclError as e:
            raise Exception("Both logo files could not be found.") from e

    # Window size
    window_width = 400
    window_height = 280

    # Get screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate x and y coordinates
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))

    # Set the window's position to the center of the screen
    root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    root.configure(bg='grey')
    label = tk.Label(root, text="QT9 QMS File Sorter Setup", bg='grey', fg='#ffffff', font=('Segoe UI Variable', 18))
    label.pack(pady=(20, 10))
    button_frame = tk.Frame(root, bg='grey')
    button_frame.place(relx=0.5, rely=0.55, anchor='center')
    check_for_update_btn = tk.Button(button_frame, text="Check For Updates", command=check_for_updates, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    check_for_update_btn.grid(row=0, column=0, pady=13)
    move_to_startup_btn = tk.Button(button_frame, text="Run App on Startup", command=run_move_to_startup, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    move_to_startup_btn.grid(row=1, column=0, pady=13)
    open_qt9_folder_btn = tk.Button(button_frame, text="Open Application Logs", command=open_qt9_folder, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    open_qt9_folder_btn.grid(row=2, column=0, pady=13)
    root.mainloop()

def show_main_gui():
    def gui_thread():
        main_gui()
    t = threading.Thread(target=gui_thread)
    t.daemon = True
    t.start()

def show_up_to_date_window():
    window = tk.Tk()
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

    # Use a Frame as a container for better control over the layout
    frame = ttk.Frame(window)
    frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    label = ttk.Label(frame, text="Your program is up to date.")
    label.place(relx=0.5, rely=0.3, anchor='center')
    
    # Pack the OK button at the bottom of the frame, centered
    ok_button = ttk.Button(frame, text="OK", command=window.destroy)
    ok_button.pack(side=tk.BOTTOM, pady=(5, 0))

    ok_button.place(anchor='se', relx=0.98, rely=1.0)
    window.mainloop()

def check_for_updates():
    try:
        response = requests.get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest")
        latest_version = response.json()['tag_name']
        print(f"Latest version: {latest_version}")
        # Convert version strings to tuples of integers for comparison
        latest_version_tuple = tuple(map(int, latest_version.strip('v').split('.')))
        current_version_tuple = tuple(map(int, CURRENT_VERSION.strip('v').split('.')))
        if latest_version_tuple > current_version_tuple:
            show_update_gui(latest_version)
        elif latest_version_tuple <= current_version_tuple:
            show_up_to_date_window()    
    except Exception as e:
        print(f"Error checking for updates: {e}")

def show_update_gui(latest_version):
    def on_install():
        download_and_install_update(latest_version)
        update_window.destroy()

    def on_decline():
        update_window.destroy()

    update_window = tk.Tk()
    update_window.attributes('-topmost', True)

    try:
        update_window.iconbitmap(r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico')
    except tk.TclError:
        try:
            update_window.iconbitmap(r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico')
        except tk.TclError as e:
            raise Exception("Both logo files could not be found.") from e
    
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
            subprocess.run(uninstaller_path + ' /SILENT', check=True)  # '/SILENT' is an example flag for silent uninstallation
            print("Old version uninstalled successfully.")
        else:
            print("Uninstaller not found. Proceeding with the installation of the new version.")
    except subprocess.CalledProcessError as e:
        print(f"Error during uninstallation: {e}")
        sys.exit(1)  # Exit if uninstallation fails to prevent potential conflicts

def download_and_install_update(latest_version):
    try:
        download_url = f"https://github.com/{GITHUB_REPO}/releases/download/{latest_version}/FileSorter_{latest_version}_Installer.exe"
        response = requests.get(download_url)
        temp_dir = tempfile.mkdtemp()
        installer_path = os.path.join(temp_dir, "FileSorter_Installer.exe")
        with open(installer_path, 'wb') as file:
            file.write(response.content)
        messagebox.showinfo("Update", "Download completed. Starting the installer.")
        os.startfile(installer_path)
        uninstall_old_version()
        sys.exit()  # Exit the current process to allow the installer to run
    except Exception as e:
        messagebox.showerror("Update Error", f"Failed to download and install update: {e}")

def periodic_check():
    check_for_updates()
    threading.Timer(3600, periodic_check).start()  # Check for updates every hour

def move_files():
    logging.info('Starting move_files function')
    core_file_names = {
        "QT9 QMS Change Control": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Change Control"),
        "QT9 QMS Doc Control": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Document Control"),
        "QT9 QMS Deviations": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Deviations"),
        "QT9 QMS Inspections": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Inspections"),
        "QT9 QMS CAPA_NCP": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\CAPA"),
        "QT9 QMS Audit Management": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Audit"),
        "QT9 QMS Supplier Surveys_Evaluations": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Supplier Surveys - Evaluations"),
        "QT9 QMS Preventive Maintenance": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Preventative Maintenance"),
        "QT9 QMS ECR_ECN": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\ECR-ECN"),
        "QT9 QMS Customer Module": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Customer Feedback - Surveys"),
        "QT9 QMS Training Module": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Training Module"),
        "QT9 QMS Calibrations": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\Calibrations"),
        "QT9 QMS Test Module": os.path.expanduser(r"~\Box\QT9 University\Training Recordings\TEST - DO NOT USE"),
    }
    for filename in os.listdir(recordings_path):
        logging.info(f'Processing file: {filename}')
        new_filename = None
        destination_folder = None
        for core_file_name, folder in core_file_names.items():
            if core_file_name in filename:
                logging.info(f'File {filename} matches core file name {core_file_name} and will be processed')
                _, file_extension = os.path.splitext(filename)
                new_filename = f"{core_file_name} {datetime.now().strftime('%m-%d-%Y')}{file_extension}"
                destination_folder = folder
                break

        if new_filename and destination_folder:
            source_file_path = os.path.join(recordings_path, filename)
            destination_file_path = f"{destination_folder}/{new_filename}"
            if os.path.exists(destination_file_path):
                logging.info(f'The file {os.path.basename(source_file_path)} already exists in the destination folder. Skipping this file.')
            else:
                try:
                    time.sleep(1)  # Wait for 1 second
                    shutil.move(source_file_path, destination_file_path)
                    logging.info(f'The file {os.path.basename(source_file_path)} has been moved successfully.')
                    send_notification(os.path.basename(source_file_path), new_filename, os.path.basename(destination_folder))
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
            timeout=5000
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
        os.startfile(os.path.join(os.path.expanduser("~"), "AppData", "Local", "QT9 QMS File Sorter"))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open folder: {str(e)}")

def quit_application(icon, item):
    icon.stop()
    global keep_running
    keep_running = False
    os._exit(0)

def signal_handler(signum, frame):
    global keep_running
    logging.info('Signal received, stopping observer.')
    keep_running = False

def main():
    periodic_check()
    
    tray_thread = threading.Thread(target=setup_system_tray)
    tray_thread.start()

    setup_thread = threading.Thread(target=main_gui)
    setup_thread.start()

    global keep_running
    keep_running = True
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path=recordings_path, recursive=False)
    logging.info('Starting the observer')
    observer.start()
    logging.info(f'Observer started and is monitoring: {recordings_path}')
    
    try:
        while keep_running:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info('KeyboardInterrupt received, stopping observer.')
    except Exception as e:
        logging.error(f'Unexpected error in main loop. Error: {e}')
    finally:
        observer.stop()
        observer.join()
        logging.info('Observer has been successfully stopped')

if __name__ == "__main__":
    main()