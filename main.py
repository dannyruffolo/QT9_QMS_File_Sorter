import os
import shutil
import time
import logging
import signal
import threading
import tkinter as tk
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
from tkinter import messagebox
from datetime import datetime
from logging.handlers import RotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from plyer import notification

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

def show_gui(icon, item):
    icon.stop()
    create_gui()

def quit_application(icon, item):
    icon.stop()
    global keep_running
    keep_running = False

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
    menu = (item('Open Setup Wizard', show_gui), item('Quit', quit_application))
    # Create and run the system tray icon
    icon = pystray.Icon("QT9 QMS File Sorter", icon_image, "QT9 QMS File Sorter", menu)
    icon.run()

def show_splash_screen(duration=3):
    splash_root = tk.Tk()
    splash_root.overrideredirect(True)
    splash_root.attributes("-alpha", 0.9)
    bg_color = "#333333"
    text_color = "#FFFFFF"
    font = ("Segoe UI Variable", 20)
    splash_root.configure(bg=bg_color)
    window_width = 400
    window_height = 250
    screen_width = splash_root.winfo_screenwidth()
    screen_height = splash_root.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))
    splash_root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
    frame = tk.Frame(splash_root, bg=bg_color, bd=5)
    frame.place(relx=0.5, rely=0.5, anchor="center")

    # Attempt to find and load the logo image from multiple paths
    logo_paths = [
        r'C:\Program Files\QT9 QMS File Sorter\QT9Logo.png',
        r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\QT9Logo.png',
        r'C:/Users/druffolo/Downloads/QT9Logo.png'  # Add or update paths as necessary
    ]
    logo = None
    for path in logo_paths:
        try:
            logo = tk.PhotoImage(file=path)
            break  # Exit the loop if the file is found and loaded successfully
        except tk.TclError:
            continue  # Try the next path if the current one fails

    if not logo:
        raise Exception("Logo file could not be found in any of the specified paths.")

    logo_label = tk.Label(frame, image=logo, bg=bg_color)
    logo_label.pack()
    splash_label = tk.Label(frame, text="Starting Application...", font=font, fg=text_color, bg=bg_color)
    splash_label.pack(expand=True, fill=tk.BOTH, pady=(15, 0))
    splash_root.after(duration * 1000, splash_root.destroy)
    splash_root.mainloop()

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

def create_gui():
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
    window_height = 250

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
    button_frame.place(relx=0.5, rely=0.5, anchor='center')
    move_to_startup_btn = tk.Button(button_frame, text="Run App on Startup", command=run_move_to_startup, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    move_to_startup_btn.grid(row=0, column=0, pady=15)
    open_qt9_folder_btn = tk.Button(button_frame, text="Open Application Logs", command=open_qt9_folder, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
    open_qt9_folder_btn.grid(row=1, column=0, pady=20)
    root.mainloop()

class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        logging.info(f'The file {os.path.basename(event.src_path)} has been created!')
        move_files()

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

def signal_handler(signum, frame):
    global keep_running
    logging.info('Signal received, stopping observer.')
    keep_running = False

def main():
    tray_thread = threading.Thread(target=setup_system_tray)
    tray_thread.start()

    show_splash_screen(2)

    setup_thread = threading.Thread(target=create_gui)
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