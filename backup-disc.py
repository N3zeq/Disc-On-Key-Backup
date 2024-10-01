import os
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading  # Import the threading module
from loguru import logger

# Define the file extensions for Word documents
WORD_EXTENSIONS = {'.doc', '.docx'}


logger.add(r"C:\Logs\DOK_Backup_{time}.log", rotation="500 MB",
           retention="10 Days", compression="zip", enqueue=True)


class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, src_dir, dest_dir):
        self.src_dir = src_dir
        self.dest_dir = dest_dir

    def on_modified(self, event):
        if not event.is_directory and self.is_word_file(event.src_path):
            src_file_path = event.src_path
            dest_file_path = os.path.join(self.dest_dir, os.path.relpath(src_file_path, self.src_dir))

            # Ensure the destination directory exists
            os.makedirs(os.path.dirname(dest_file_path), exist_ok=True)
            shutil.copy2(src_file_path, dest_file_path)
            print(f"Copied '{src_file_path}' to '{dest_file_path}'")
            logger.info(f"Copied '{src_file_path}' to '{dest_file_path}'")

    def is_word_file(self, file_path):
        return os.path.splitext(file_path)[1].lower() in WORD_EXTENSIONS


def initial_copy(src_dir, dest_dir):
    for root, dirs, files in os.walk(src_dir):
        for file in files:
            if is_word_file(file):
                src_file_path = os.path.join(root, file)
                dest_file_path = os.path.join(dest_dir, os.path.relpath(src_file_path, src_dir))

                # Ensure the destination directory exists
                os.makedirs(os.path.dirname(dest_file_path), exist_ok=True)
                shutil.copy2(src_file_path, dest_file_path)
                print(f"Copied initial file '{src_file_path}' to '{dest_file_path}'")
                logger.info(f"Copied initial file '{src_file_path}' to '{dest_file_path}'")


def is_word_file(file_name):
    return os.path.splitext(file_name)[1].lower() in WORD_EXTENSIONS


def monitor_directory(src_dir, dest_dir):
    event_handler = FileChangeHandler(src_dir, dest_dir)
    observer = Observer()
    observer.schedule(event_handler, src_dir, recursive=True)  # Monitor recursively
    observer.start()
    print(f"Monitoring changes in '{src_dir}' for Word documents...")
    logger.info(f"Monitoring changes in '{src_dir}' for Word documents...")

    try:
        while monitoring_active:
            time.sleep(1)
    except Exception as e:
        print(f"Error: {e}")
        logger.error(f"Error: {e}")
    finally:
        observer.stop()
        observer.join()


def select_source_directory():
    source_dir = filedialog.askdirectory(title="Select Source Directory")
    source_entry.delete(0, tk.END)
    source_entry.insert(0, source_dir)


def select_destination_directory():
    destination_dir = filedialog.askdirectory(title="Select Destination Directory")
    destination_entry.delete(0, tk.END)
    destination_entry.insert(0, destination_dir)


def start_monitoring():
    global monitoring_active, monitoring_thread
    src_dir = source_entry.get()
    dest_dir = destination_entry.get()

    if not src_dir or not dest_dir:
        messagebox.showerror("Error", "Please select both source and destination directories.")
        return

    # Ask for initial copy
    if messagebox.askyesno("Initial Copy", "Do you want to perform the initial copy of Word documents?"):
        initial_copy(src_dir, dest_dir)

    # Start monitoring
    monitoring_active = True
    monitoring_thread = threading.Thread(target=monitor_directory, args=(src_dir, dest_dir))
    monitoring_thread.start()


def on_closing():
    global monitoring_active
    monitoring_active = False  # Signal to stop monitoring
    root.destroy()  # Close the GUI


# Set up the GUI
root = tk.Tk()
root.title("Word Document Monitor")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

source_label = tk.Label(frame, text="Source Directory:")
source_label.grid(row=0, column=0, sticky="w")

source_entry = tk.Entry(frame, width=50)
source_entry.grid(row=0, column=1)

source_button = tk.Button(frame, text="Browse", command=select_source_directory)
source_button.grid(row=0, column=2)

destination_label = tk.Label(frame, text="Destination Directory:")
destination_label.grid(row=1, column=0, sticky="w")

destination_entry = tk.Entry(frame, width=50)
destination_entry.grid(row=1, column=1)

destination_button = tk.Button(frame, text="Browse", command=select_destination_directory)
destination_button.grid(row=1, column=2)

start_button = tk.Button(frame, text="Start Monitoring", command=start_monitoring)
start_button.grid(row=2, columnspan=3, pady=(10, 0))

# Global variables for monitoring state
monitoring_active = False
monitoring_thread = None

# Handle window closing event
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
