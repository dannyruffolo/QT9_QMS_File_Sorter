import tkinter as tk
from tkinter import messagebox
import requests
import subprocess
import os
import sys
import time
import threading

def check_for_updates():
    current_version = '2.1.1'
    repo = 'dannyruffolo/QT9_QMS_File_Sorter'
    api_url = f'https://api.github.com/repos/{repo}/releases/latest'

    try:
        requests.get(api_url)
        # Simulate a newer version for testing purposes
        latest_version = '2.2.0'  # This should be higher than current_version
        if latest_version > current_version:
            assets = [{'browser_download_url': 'http://example.com/download'}]  # Mock asset for testing
            return True, latest_version, assets[0]['browser_download_url']
        else:
            print("No downloadable assets available for the latest version.")
            return False, None, None
    except Exception as e:
        print(f"Error checking for updates: {e}")
    return False, None, None

def download_and_update(download_url, window):
    local_filename = download_url.split('/')[-1]
    with requests.get(download_url, stream=True) as r:
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    subprocess.run([local_filename, '/SILENT'])
    messagebox.showinfo("Update Installed", "The update has been successfully installed. Please restart the application.")
    window.destroy()
    os.execl(sys.executable, sys.executable, *sys.argv)

def skip_update(window):
    messagebox.showinfo("Update Skipped", "You have chosen to skip the update. The application will continue running with the current version.")
    window.destroy()

def show_update_gui(latest_version, download_url):
    update_window = tk.Toplevel()
    update_window.title("Update Available")
    
    tk.Label(update_window, text=f"Version {latest_version} is available. Do you want to download and install the update?").pack()
    
    tk.Button(update_window, text="Yes", command=lambda: download_and_update(download_url, update_window)).pack(side=tk.LEFT)
    tk.Button(update_window, text="No", command=lambda: skip_update(update_window)).pack(side=tk.RIGHT)

def check_for_updates_and_notify():
    update_available, latest_version, download_url = check_for_updates()
    if update_available:
        show_update_gui(latest_version, download_url)

def check_for_updates_periodically():
    initial_delay = 0
    interval = 20

    def run_check():
        while True:
            # Schedule check_for_updates_and_notify to run in the main thread
            root.after(0, check_for_updates_and_notify)
            time.sleep(interval)

    threading.Timer(initial_delay, run_check).start()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Application Update")
    root.withdraw()  # This hides the root window
    check_for_updates_periodically()
    root.mainloop()