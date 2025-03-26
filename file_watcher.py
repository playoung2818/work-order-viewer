import os
import time
import logging
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Logging Configuration
logging.basicConfig(level=logging.INFO)

# Paths to monitor
WATCHED_FOLDERS = [
    r"\\NEOUSYSSERVER\Drive D\QuickBooks\2- Year 2024\Work Order- WO",
    r"\\NEOUSYSSERVER\Drive D\QuickBooks\3- Year 2025\Work Order- WO",
    r"C:\Users\Admin\OneDrive - neousys-tech\Share NTA Warehouse\02 Work Order- Word file\Work Order 2024",
    r"C:\Users\Admin\OneDrive - neousys-tech\Share NTA Warehouse\02 Work Order- Word file\Work Order 2025"
]

# Path to your main processing script
FILE_SCRIPT_DEBUG = r"c:\Users\Admin\OneDrive - neousys-tech\Desktop\File Watcher\File Script Debug.py"

class FileEventHandler(FileSystemEventHandler):
    """Handles file creation events and triggers another script."""

    def on_created(self, event):
        if event.is_directory:
            return

        file_path = event.src_path
        file_name = os.path.basename(file_path)

        # Ignore temporary files
        if file_name.startswith("~RF") and file_name.endswith(".TMP"):
            logging.info(f"🛑 Ignoring temporary file: {file_path}")
            return

        logging.info(f"📁 New file detected: {file_path}")
        logging.info("🚀 Running 'File Script Debug.py'...")

        try:
            # Run the script in a new terminal
            subprocess.run(["python", FILE_SCRIPT_DEBUG], shell=True, check=True)
            logging.info(f"✅ Successfully ran '{FILE_SCRIPT_DEBUG}'")
        except subprocess.CalledProcessError as e:
            logging.error(f"❌ Failed to run '{FILE_SCRIPT_DEBUG}': {e}")

def start_monitoring():
    """Starts monitoring the folders."""
    event_handler = FileEventHandler()
    observer = Observer()

    for folder in WATCHED_FOLDERS:
        if os.path.exists(folder):
            observer.schedule(event_handler, folder, recursive=True)
            logging.info(f"🚀 Monitoring started on {folder}")
        else:
            logging.warning(f"⚠️ Folder does not exist: {folder}")

    observer.start()
    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    logging.info("🔧 Starting file watcher...")
    start_monitoring()
