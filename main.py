import os
import shutil
import time
import threading
import signal
from pathlib import Path
from datetime import datetime

import logging
from logging.handlers import RotatingFileHandler

import tkinter as tk
from PIL import ImageTk
from tkinter import messagebox

from PIL import Image
from plyer import notification
import pystray
from pystray import MenuItem as item

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

recordings_path = os.path.expanduser(r"~\OneDrive - QT9 Software\Recordings")


class Application:
    def __init__(self):
        """
        Initializes the Application object, setting up logging, defining essential paths, and initializing core file names.
        """
        # Initialize icon_image_paths before calling setup_system_tray
        self.icon_image_paths = [
            os.path.join('C:', os.sep, 'Program Files', 'QT9 QMS File Sorter', 'app_icon.ico'),
            os.path.join('C:', os.sep, 'Users', 'druffolo', 'Desktop', 'File Sorter Installer & EXE Files', 'app_icon.ico')
        ]
        self.logo_paths = [
            os.path.join('C:', os.sep, 'Program Files', 'QT9 QMS File Sorter', 'QT9Logo.png'),
            os.path.join('C:', os.sep, 'Users', 'druffolo', 'Desktop', 'File Sorter Installer & EXE Files', 'QT9Logo.png'),
            Path('C:/Users/druffolo/Downloads/QT9Logo.png')
        ]
        self.show_splash_screen()
        self.show_gui()
        self.setup_logging()
        self.setup_system_tray()  # Now this is called after icon_image_paths is defined
        self.keep_running = True
        # Register signal handlers
        signal.signal(signal.SIGTERM, self.signal_handler)

        self.core_file_names = {
            'QT9 QMS Change Control': Path.home() / 'Box/QT9 University/Training Recordings/Change Control',
            'QT9 QMS Doc Control': Path.home() / 'Box/QT9 University/Training Recordings/Document Control',
            'QT9 QMS Deviations': Path.home() / 'Box/QT9 University/Training Recordings/Deviations',
            'QT9 QMS Inspections': Path.home() / 'Box/QT9 University/Training Recordings/Inspections',
            'QT9 QMS CAPA_NCP': Path.home() / 'Box/QT9 University/Training Recordings/CAPA',
            'QT9 QMS Audit Management': Path.home() / 'Box/QT9 University/Training Recordings/Audit',
            'QT9 QMS Supplier Surveys_Evaluations': Path.home() / 'Box/QT9 University/Training Recordings/Supplier Surveys - Evaluations',
            'QT9 QMS Preventive Maintenance': Path.home() / 'Box/QT9 University/Training Recordings/Preventative Maintenance',
            'QT9 QMS ECR_ECN': Path.home() / 'Box/QT9 University/Training Recordings/ECR-ECN',
            'QT9 QMS Customer Module': Path.home() / 'Box/QT9 University/Training Recordings/Customer Feedback - Surveys',
            'QT9 QMS Training Module': Path.home() / 'Box/QT9 University/Training Recordings/Training Module',
            'QT9 QMS Calibrations': Path.home() / 'Box/QT9 University/Training Recordings/Calibrations',
            'QT9 QMS Test Module': Path.home() / 'Box/QT9 University/Training Recordings/TEST - DO NOT USE',
        }

    def setup_logging(self):
        """
        Sets up logging for the application, including file rotation and formatting.
        """
        log_format = '%(asctime)s - %(name)s - %(levelname)s - [Line #%(lineno)d] - %(message)s'
        log_file_path = Path.home() / 'AppData/Local/QT9 QMS File Sorter'
        log_file_path.mkdir(parents=True, exist_ok=True)
        log_file = log_file_path / 'app.log'

        # Set up RotatingFileHandler
        log_handler = RotatingFileHandler(log_file, mode='a', maxBytes=5*1024*1024, backupCount=2, encoding=None, delay=0)
        log_handler.setLevel(logging.INFO)
        formatter = logging.Formatter(log_format, datefmt='%m-%d-%Y %H:%M:%S')
        log_handler.setFormatter(formatter)

        # Get the root logger and set the handler
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logger.addHandler(log_handler)

    def load_image(self, paths, image_type='icon'):
        """
        Attempts to load an image from a list of paths.

        :param paths: A list of file paths to try loading the image from.
        :param image_type: The type of image to load ('icon' or 'logo').
        :return: The loaded image, or None if all paths fail.
        """
        for path in paths:
            try:
                if image_type == 'icon':
                    return Image.open(path)
                elif image_type == 'logo':
                    return tk.PhotoImage(file=path)
            except FileNotFoundError:
                logging.warning(f'File not found: {path}. Trying next.')
            except IOError as e:
                logging.error(f'IO error when opening {path}: {e}')
                continue
        logging.error(f'None of the specified paths contain the {image_type} image.')
        return None

    def show_gui(self, icon=None, item=None):
        """
        Shows the main GUI window.

        :param icon: Optional. The system tray icon object.
        :param item: Optional. The menu item selected.
        """
        if icon:
            icon.stop()
        self.create_gui()

    def quit_application(self, icon=None, item=None):
        """
        Quits the application, optionally stopping the system tray icon.
    
        :param icon: Optional. The system tray icon object.
        :param item: Optional. The menu item selected.
        """
        if icon:
            icon.stop()
        self.keep_running = False  # Ensure this attribute is used to control the main loop of the application

    def setup_system_tray(self):
        """
        Sets up the system tray icon and menu for the application.
        """
        icon_image = self.load_image(self.icon_image_paths, 'icon')
        if icon_image is None:
            logging.error('System tray icon setup failed due to missing icon image.')
            return

        menu = (item('Open Setup Wizard', self.show_gui), item('Quit', self.quit_application))
        icon = pystray.Icon('QT9 QMS File Sorter', icon_image, 'QT9 QMS File Sorter', menu)
        icon.run()

    def setup_window(self, root, width, height, title):
        """
        Configures the main window's size, position, and title.

        :param root: The Tk root window object.
        :param width: The desired width of the window.
        :param height: The desired height of the window.
        :param title: The title of the window.
        """
        root.title(title)
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x_coordinate = int((screen_width / 2) - (width / 2))
        y_coordinate = int((screen_height / 2) - (height / 2))
        root.geometry(f'{width}x{height}+{x_coordinate}+{y_coordinate}')
        root.configure(bg='grey')

    def show_splash_screen(self, duration=3):
        """
        Displays a splash screen for a specified duration.

        :param duration: The duration in seconds to display the splash screen.
        """
        try:
            splash_root = tk.Tk()
            self.setup_window(splash_root, 400, 250, 'Starting Application...')

            splash_root.overrideredirect(True)
            splash_root.attributes('-alpha', 0.9)
            bg_color = '#333333'
            text_color = '#FFFFFF'
            font = ('Segoe UI Variable', 20)
            splash_root.configure(bg=bg_color)  # Set the background color of the splash window

            logo = self.load_image(self.logo_paths, 'logo')
            if logo is None:
                logging.error('Splash screen setup failed due to missing logo image.')
                return

            frame = tk.Frame(splash_root, bg=bg_color, bd=5)
            frame.place(relx=0.5, rely=0.5, anchor='center')

            logo_label = tk.Label(frame, image=logo, bg=bg_color)
            logo_label.pack()
            splash_label = tk.Label(frame, text='Starting Application...', font=font, fg=text_color, bg=bg_color)
            splash_label.pack()

            splash_root.after(duration * 1000, splash_root.destroy)
            splash_root.mainloop()
        except Exception as e:
            logging.error(f'Failed to show splash screen: {str(e)}')

    def run_move_to_startup(self):
        """
        Moves the application shortcut to the startup folder to run at system startup.
        """
        source_path = Path('C:/ProgramData/Microsoft/Windows/Start Menu/Programs')
        username = os.getlogin()
        destination_path = Path(f'C:/Users/{username}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup')
        shortcut_name = 'QT9 QMS File Sorter.lnk'
        source_shortcut_path = source_path / shortcut_name
        destination_path.mkdir(parents=True, exist_ok=True)
        if source_shortcut_path.exists():
            shutil.copy2(source_shortcut_path, destination_path)
            logging.info('Application setup to run on startup.')
            messagebox.showinfo('Success', 'The application has been set to run on startup.')
        else:
            logging.error(f'Shortcut \'{shortcut_name}\' not found in the source directory.')

    def open_qt9_folder(self):
        """
        Opens the application's log folder in the file explorer.
        """
        try:
            os.startfile(str(Path.home() / 'AppData/Local/QT9 QMS File Sorter'))
        except Exception as e:
            messagebox.showerror('Error', f'Failed to open folder: {str(e)}')

    def run_move_to_startup_threaded(self):
        """
        Wraps the run_move_to_startup method in a thread to prevent UI blocking.
        """
        operation_thread = threading.Thread(target=self.run_move_to_startup)
        operation_thread.start()

    def create_gui(self):
        """
        Creates and displays the main GUI for the application setup wizard.
        """
        root = tk.Tk()
        self.setup_window(root, 400, 250, 'Setup Wizard')
    
        try:
            icon_image = self.load_image(self.icon_image_paths, 'icon')
            # Open the icon image using PIL
            pil_image = Image.open(icon_image.filename)
            # Convert the PIL image object to a Tkinter-compatible photo image
            tk_image = ImageTk.PhotoImage(pil_image)
            root.iconphoto(True, tk_image)
        except Exception as e:
            messagebox.showerror('Error', str(e))
    
        label = tk.Label(root, text='QT9 QMS File Sorter Setup', bg='grey', fg='#ffffff', font=('Segoe UI Variable', 18))
        label.pack(pady=(20, 10))
    
        button_frame = tk.Frame(root, bg='grey')
        button_frame.place(relx=0.5, rely=0.5, anchor='center')
    
        move_to_startup_btn = tk.Button(button_frame, text='Run App on Startup', command=self.run_move_to_startup_threaded, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
        move_to_startup_btn.grid(row=0, column=0, pady=15)
    
        open_qt9_folder_btn = tk.Button(button_frame, text='Open Application Logs', command=self.open_qt9_folder, fg='#ffffff', bg='#0056b8', font=('Segoe UI Variable', 14), width=19)
        open_qt9_folder_btn.grid(row=1, column=0, pady=20)
    
        root.mainloop()

    def signal_handler(self, signum, frame):
        """
        Handles interrupt signals to gracefully shut down the application.

        :param signum: The signal number.
        :param frame: The current stack frame.
        """
        logging.info('Signal received, stopping observer.')
        self.keep_running = False

    def send_notification(self, original_file_name, new_file_name, destination_folder):
        """
        Sends a desktop notification about a file operation.

        :param original_file_name: The original name of the file.
        :param new_file_name: The new name of the file after operation.
        :param destination_folder: The folder to which the file was moved.
        """
        try:
            logging.info(f'Attempting to send notification for {new_file_name}')
            notification.notify(
                title='QT9 U Recording Transfer',
                message=f'The file "{original_file_name}" has been renamed to "{new_file_name}" and moved to \\Training Recordings\\{destination_folder}.',
                timeout=5000
            )
        except Exception as e:
            logging.error(f'An error occurred while trying to send a notification: {e}')

    def move_files(self):
        """
        Moves files from the recordings path to their respective destination folders based on core file names.

        This function iterates over all files in the recordings path, checks if the file name contains any of the core file names,
        and moves the file to the corresponding destination folder. It also handles file name conflicts, permission errors,
        and other exceptions during the file move operation. Notifications are sent for successful moves.
        """
        logging.info('Starting move_files function')
        for filename in os.listdir(recordings_path):
            logging.info(f'Processing file: {filename}')
            new_filename = None
            destination_folder = None
            for core_file_name, folder in self.core_file_names.items():
                if core_file_name in filename:
                    logging.info(f'File {filename} matches core file name {core_file_name} and will be processed')
                    _, file_extension = os.path.splitext(filename)
                    new_filename = f'{core_file_name} {datetime.now().strftime("%m-%d-%Y")}{file_extension}'
                    destination_folder = folder
                    break
            if new_filename and destination_folder:
                source_file_path = os.path.join(recordings_path, filename)
                destination_file_path = f'{destination_folder}/{new_filename}'
                if os.path.exists(destination_file_path):
                    logging.info(f'The file {os.path.basename(source_file_path)} already exists in the destination folder. Skipping this file.')
                else:
                    try:
                        time.sleep(1)  # Wait for 1 second
                        shutil.move(source_file_path, destination_file_path)
                        logging.info(f'The file {os.path.basename(source_file_path)} has been moved successfully.')
                        self.send_notification(os.path.basename(source_file_path), new_filename, os.path.basename(destination_folder))
                    except PermissionError as e:
                        logging.error(f'Permission denied for {source_file_path} ({filename}). Error: {e}')
                    except IOError as e:
                        logging.error(f'IO error during file move from {source_file_path} to {destination_file_path}. Error: {e}')
                    except Exception as e:
                        logging.error(f'An unexpected error occurred while moving {source_file_path} to {destination_file_path}. Error: {e}')
                    except Exception as e:
                        logging.error(f'An unexpected error occurred while moving {source_file_path} to {destination_file_path}. Error: {e}')
            else:
                logging.info(f'No matching core file name found for {filename}. The file will not be moved.')

    def start_observer(self):
        """
        Starts a file system observer that watches for file creation in a specified path.

        This function sets up an observer to monitor a directory for new files. When a new file is created,
        it triggers the move_files function to process and potentially move the file to a designated folder.
        The observer runs in a loop until the application is manually stopped.
        """
        event_handler = MyHandler(self)  # Pass the current instance of Application
        observer = Observer()
        observer.schedule(event_handler, recordings_path, recursive=False)
        observer.start()
        try:
            while self.keep_running:
                time.sleep(1)
        except KeyboardInterrupt:
            observer.stop()
        observer.join()

class MyHandler(FileSystemEventHandler):
    def __init__(self, app_instance):
        """
        Initializes the MyHandler class with an application instance.

        :param app_instance: The instance of the Application class to use for moving files when an event is triggered.
        """
        self.app_instance = app_instance
    
    def on_created(self, event):
        """
        Handles the on_created event by moving files when a new file is detected in the watched directory.

        :param event: The event object containing information about the created file.
        """
        logging.info(f'The file {os.path.basename(event.src_path)} has been created!')
        self.app_instance.move_files()

if __name__ == '__main__':
    app = Application()
    app.show_splash_screen()
    app.show_gui()
    app.setup_system_tray()
    app.start_observer()
    app.run()