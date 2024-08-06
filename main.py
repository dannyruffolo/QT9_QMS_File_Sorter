import os
import sys
import shutil
import time
import logging
import signal
import subprocess
import threading
import pystray
import tempfile
import requests
import json
from pystray import MenuItem as item
from PIL import Image
from datetime import datetime
from logging.handlers import RotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from plyer import notification
from PyQt5.QtCore import QMetaObject, Qt, pyqtSignal, QThread, QObject, Q_ARG, QTimer
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QHBoxLayout, QMessageBox, QFileDialog
from PyQt5.QtGui import QIcon, QPixmap

def setup_logging():
    log_format = '%(asctime)s - %(name)s - %(levelname)s - [Line #%(lineno)d] - %(message)s'
    log_file_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "QT9 QMS File Sorter")
    if not os.path.exists(log_file_path):
        os.makedirs(log_file_path)
    log_file = os.path.join(log_file_path, 'app.log')

    log_handler = RotatingFileHandler(log_file, mode='a', maxBytes=5*1024*1024, backupCount=2, encoding=None, delay=0)
    log_handler.setLevel(logging.INFO)
    formatter = logging.Formatter(log_format, datefmt='%m-%d-%Y %H:%M:%S')
    log_handler.setFormatter(formatter)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.addHandler(log_handler)

class GlobalState:
    def __init__(self):
        self.CURRENT_VERSION = "3.0.0"
        self.GITHUB_REPO = "dannyruffolo/QT9_QMS_File_Sorter"
        self.lock = threading.Lock()
        self.recordings_path = os.path.expanduser(r"~\OneDrive - QT9 Software\Recordings")
        self.user_preferences = {}
        self.selected_folder = None

global_state = GlobalState()

class AppState:
    def __init__(self):
        self.user_preferences = {}
        self.selected_folder = ""

    def get_user_preferences(self):
        return self.user_preferences

    def set_user_preferences(self, preferences):
        self.user_preferences = preferences

    def get_selected_folder(self):
        return self.selected_folder

    def set_selected_folder(self, folder):
        self.selected_folder = folder

def show_message(message):
    app = create_app_instance()
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Information)
    msg_box.setText(message)
    msg_box.setWindowTitle("Information")
    msg_box.exec_()

def handle_update_signal(message):
    logging.info(f"Handling update signal with message: {message}")
    if message == "up_to_date":
        show_message("Your application is up to date.")
    else:
        show_update_gui("A new update is available.")

class Worker(QObject):
    update_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    def check_for_updates(self, show_up_to_date=False):
        logging.info("Worker: Checking for updates")
        try:
            response = requests.get(f"https://api.github.com/repos/{global_state.GITHUB_REPO}/releases/latest")
            response.raise_for_status()
            latest_version = response.json()['tag_name']
            latest_version_tuple = tuple(map(int, latest_version.strip('v').split('.')))
            current_version_tuple = tuple(map(int, global_state.CURRENT_VERSION.strip('v').split('.')))
            if latest_version_tuple > current_version_tuple:
                self.update_signal.emit(latest_version)
            elif latest_version_tuple <= current_version_tuple and show_up_to_date:
                self.update_signal.emit("up_to_date")
        except requests.exceptions.RequestException as req_err:
            logging.error(f"Network error while checking for updates: {req_err}")
        except json.JSONDecodeError as json_err:
            logging.error(f"JSON parsing error while checking for updates: {json_err}")
        except Exception as e:
            logging.error(f"Unexpected error while checking for updates: {e}")
        
        time.sleep(2)  # Simulate delay
        self.update_signal.emit("up_to_date" if show_up_to_date else "new_update")

def check_for_updates(show_up_to_date=False):
    logging.info("Main: Initiating check for updates")
    worker = Worker()
    worker.update_signal.connect(handle_update_signal)
    QTimer.singleShot(0, lambda: worker.check_for_updates(show_up_to_date))

def setup_system_tray():
    try:
        icon_image = Image.open("icon.png")
    except FileNotFoundError:
        icon_image = Image.new('RGB', (64, 64), color = 'red')

    # Menu items
    menu = (item('Open Menu', show_main_gui),
            item('Check for Updates', lambda: check_for_updates(True)), 
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

def create_app_instance():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app

def center_window(window):
    app = create_app_instance()
    screen = app.primaryScreen().availableGeometry()
    x_coordinate = (screen.width() - window.width()) // 2
    y_coordinate = (screen.height() - window.height()) // 2
    window.move(x_coordinate, y_coordinate)

def main_gui():
    main_window = QMainWindow()
    main_window.setWindowTitle("Menu")
    try:
        main_window.setWindowIcon(QIcon(r'C:\Program Files\QT9 QMS File Sorter\app_icon.ico'))
    except Exception:
        try:
            main_window.setWindowIcon(QIcon(r'C:\Users\druffolo\Desktop\File Sorter Installer & EXE Files\app_icon.ico'))
        except Exception as e:
            raise Exception("Both logo files could not be found.") from e
    main_widget = QWidget()
    main_window.setCentralWidget(main_widget)
    layout = QVBoxLayout(main_widget)
    label = QLabel("QT9 QMS File Sorter")
    label.setStyleSheet("color: #ffffff; background-color: grey; font: 18pt 'Segoe UI Variable';")
    layout.addWidget(label)
    check_for_update_btn = QPushButton("Check For Updates")
    check_for_update_btn.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 14pt 'Segoe UI Variable';")
    check_for_update_btn.clicked.connect(lambda: Worker().check_for_updates(True))
    layout.addWidget(check_for_update_btn)
    move_to_startup_btn = QPushButton("Run App on Startup")
    move_to_startup_btn.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 14pt 'Segoe UI Variable';")
    move_to_startup_btn.clicked.connect(run_move_to_startup)
    layout.addWidget(move_to_startup_btn)
    open_qt9_folder_btn = QPushButton("Open Application Logs")
    open_qt9_folder_btn.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 14pt 'Segoe UI Variable';")
    open_qt9_folder_btn.clicked.connect(open_qt9_folder)
    layout.addWidget(open_qt9_folder_btn)
    open_config_btn = QPushButton("Open Config")
    open_config_btn.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 14pt 'Segoe UI Variable';")
    open_config_btn.clicked.connect(config_gui)
    layout.addWidget(open_config_btn)
    main_window.resize(400, 350)
    main_window.show()

def show_main_gui():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    main_gui()
    app.exec_()

def show_update_gui(message):
    app = create_app_instance()
    window = QMainWindow()
    window.setWindowTitle("Update Check")
    window.setFixedSize(232, 114)
    center_window(window)
    window.setWindowFlags(window.windowFlags() & ~Qt.WindowMinMaxButtonsHint)

    main_widget = QWidget()
    main_widget.setStyleSheet("background-color: grey;")
    window.setCentralWidget(main_widget)
    layout = QVBoxLayout(main_widget)
    layout.setContentsMargins(10, 10, 10, 10)

    label = QLabel(f"New version available: {message}")
    label.setStyleSheet("color: #ffffff; font: bold 10pt 'Segoe UI Variable';")
    label.setAlignment(Qt.AlignCenter)
    layout.addWidget(label)

    ok_button = QPushButton("OK")
    ok_button.setStyleSheet("color: #ffffff; background-color: #0056b8; font: bold 10pt 'Segoe UI Variable';")
    ok_button.clicked.connect(window.close)
    layout.addWidget(ok_button, alignment=Qt.AlignRight)

    window.show()

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
        download_url = f"https://github.com/{global_state.GITHUB_REPO}/releases/download/{latest_version}/FileSorter_{latest_version}_Installer.exe"
        response = requests.get(download_url)
        response.raise_for_status()  # Ensure we catch HTTP errors
        temp_dir = tempfile.mkdtemp()
        installer_path = os.path.join(temp_dir, "FileSorter_Installer.exe")
        with open(installer_path, 'wb') as file:
            file.write(response.content)
        
        app = create_app_instance()
        QMessageBox.information(None, "Update", "Download completed. Starting the installer.")
        
        os.startfile(installer_path)
        uninstall_old_version()
        sys.exit()  # Exit the current process to allow the installer to run

    except requests.exceptions.HTTPError as http_err:
        logging.error(f"HTTP error occurred while downloading the update: {http_err}")
    except requests.exceptions.ConnectionError as conn_err:
        logging.error(f"Connection error occurred while downloading the update: {conn_err}")
    except requests.exceptions.Timeout as timeout_err:
        logging.error(f"Timeout error occurred while downloading the update: {timeout_err}")
    except requests.exceptions.RequestException as req_err:
        logging.error(f"An error occurred while downloading the update: {req_err}")
    except subprocess.CalledProcessError as sub_err:
        logging.error(f"An error occurred while installing the update: {sub_err}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def show_error_message(message):
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    error_dialog = QMessageBox()
    error_dialog.setIcon(QMessageBox.Critical)
    error_dialog.setText(message)
    error_dialog.setWindowTitle("Error")
    error_dialog.exec_()

def periodic_check():
    check_for_updates()
    threading.Timer(3600, periodic_check).start()

def select_destination_folder(app_state):
    folder = QFileDialog.getExistingDirectory()
    if folder:
        QMetaObject.invokeMethod(lambda: app_state.set_selected_folder(folder), Qt.QueuedConnection)
    else:
        QMetaObject.invokeMethod(lambda: show_message("No folder selected"), Qt.QueuedConnection)

def add_to_preferences(app_state, destination_folder_label, file_name_entry):
    folder_path = destination_folder_label.text()
    file_name = file_name_entry.text()
    if folder_path == "No folder selected" or not file_name:
        show_message("Please select a folder and enter a file name.")
        return

    with global_state.lock:
        app_state.preferences[file_name] = folder_path

    QMetaObject.invokeMethod(preferences_display_label, "append", Qt.QueuedConnection, Q_ARG(str, f"Added {file_name}: {folder_path} to preferences"))

def save_user_preferences(app_state):
    with global_state.lock:
        app_state.save_user_preferences()
    QMetaObject.invokeMethod(preferences_display_label, "append", Qt.QueuedConnection, Q_ARG(str, "Preferences saved successfully."))
    update_preferences_display(app_state, preferences_display_label)

def update_preferences_display(app_state, preferences_display_label):
    user_preferences = app_state.get_user_preferences()
    display_text = ""
    for file_name, folders in user_preferences.items():
        display_text += f"{file_name}: {folders}\n"
    QMetaObject.invokeMethod(preferences_display_label, "setPlainText", Qt.QueuedConnection, Q_ARG(str, display_text))

def load_user_preferences(app_state):
    preferences_file_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "QT9 QMS File Sorter", "preferences_file.json")
    try:
        with open(preferences_file_path, 'r') as file:
            preferences = json.load(file)
            app_state.set_user_preferences(preferences)
    except FileNotFoundError:
        show_message("Preferences file not found.")
    except json.JSONDecodeError:
        show_message("Error decoding preferences file.")

def config_gui(app_state):
    global file_name_entry, destination_folder_label, preferences_display_label
    app = create_app_instance()
    config = QMainWindow()
    config.setWindowTitle("File Sorter Configuration")
    config.setFixedSize(810, 500)
    center_window(config)

    main_widget = QWidget()
    main_widget.setStyleSheet("background-color: grey;")
    config.setCentralWidget(main_widget)
    layout = QVBoxLayout(main_widget)
    layout.setContentsMargins(10, 10, 10, 10)

    label = QLabel("QT9 QMS File Sorter Configuration")
    label.setStyleSheet("color: #ffffff; font: 18pt 'Segoe UI Variable';")
    label.setAlignment(Qt.AlignCenter)
    layout.addWidget(label)

    input_frame = QHBoxLayout()
    layout.addLayout(input_frame)

    file_name_label = QLabel("File Name Contains:")
    file_name_label.setStyleSheet("color: #ffffff; font: 13pt 'Segoe UI Variable';")
    input_frame.addWidget(file_name_label)

    file_name_entry = QLineEdit()
    file_name_entry.setStyleSheet("font: 13pt 'Segoe UI Variable';")
    input_frame.addWidget(file_name_entry)

    button_frame = QHBoxLayout()
    layout.addLayout(button_frame)

    select_folder_button = QPushButton("Select Folder")
    select_folder_button.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 13pt 'Segoe UI Variable';")
    select_folder_button.clicked.connect(lambda: select_destination_folder(app_state))
    button_frame.addWidget(select_folder_button)

    destination_folder_label = QLabel("No folder selected")
    destination_folder_label.setStyleSheet("background-color: white; color: black; font: 13pt 'Segoe UI Variable';")
    button_frame.addWidget(destination_folder_label)

    add_button = QPushButton("Add to Preferences")
    add_button.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 13pt 'Segoe UI Variable';")
    add_button.clicked.connect(lambda: add_to_preferences(app_state, destination_folder_label, file_name_entry))
    layout.addWidget(add_button)

    preferences_display_label = QTextEdit()
    preferences_display_label.setStyleSheet("font: 13pt 'Segoe UI Variable';")
    preferences_display_label.setReadOnly(True)
    layout.addWidget(preferences_display_label)

    save_button = QPushButton("Save Preferences")
    save_button.setStyleSheet("color: #ffffff; background-color: #0056b8; font: 13pt 'Segoe UI Variable';")
    save_button.clicked.connect(lambda: save_user_preferences(app_state))
    layout.addWidget(save_button)

    config.show()

def move_files():
    global selected_folder
    logging.info('Starting move_files function')

    global user_preferences  # Assuming user_preferences is defined globally

    # Check if user_preferences is a string and convert it to a dictionary if so
    with global_state.lock:
        if isinstance(global_state.user_preferences, str):
            user_preferences = json.loads(global_state.user_preferences)
        else:
            user_preferences_dict = global_state.user_preferences

    for filename in os.listdir(global_state.recordings_path):
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
            source_file_path = os.path.join(global_state.recordings_path, filename)
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
            app = QApplication(sys.argv)
            QMessageBox.information(None, "Success", "Application setup to run on startup.")
        else:
            app = QApplication(sys.argv)
            QMessageBox.critical(None, "Error", f"Shortcut '{shortcut_name}' not found in the source directory.")
    except Exception as e:
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Error", f"Failed to move shortcut: {str(e)}")

def open_qt9_folder():
    try:
        os.startfile(os.path.join(os.path.expanduser("~"), "AppData", "Local", "QT9 QMS File Sorter"))
    except Exception as e:
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Error", f"Failed to open folder: {str(e)}")

def quit_application(icon, item):
    icon.stop()
    QApplication.quit()
    os._exit(0)

def signal_handler(signum, frame):
    QApplication.quit()

def main():
    setup_logging()
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    # Create the worker and the thread
    worker = Worker()
    worker_thread = QThread()
    
    # Move the worker to the thread
    worker.moveToThread(worker_thread)
    
    # Start the thread
    worker_thread.start()
    
    # Check for updates
    check_for_updates()

    # Setup the system tray
    setup_system_tray()
    
    # Show the main GUI
    show_main_gui()
    
    global keep_running
    keep_running = True
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path=global_state.recordings_path, recursive=False)
    logging.info('Starting the observer')
    observer.start()
    logging.info(f'Observer started and is monitoring: {global_state.recordings_path}')
    
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
        worker_thread.quit()
        worker_thread.wait()

    app.exec_()

if __name__ == "__main__":
    main()