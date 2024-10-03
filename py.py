import os
import time
import win32com.client as win32
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import tkinter as tk
from tkinter import messagebox

# Basic playlist style ui
# Able to change source folder
# Set holding slide [bg-image]
# Deploy a portable .exe file
# Create a server so can send/receive data from another device

class PowerPointHandler:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.ppt_app = win32.Dispatch("PowerPoint.Application")
        self.ppt_app.Visible = True
        self.presentations = []
        self.current_presentation = None
        self.current_index = 0  # Start at the first presentation
        self.load_presentations()

    def load_presentations(self):
        # Load all non-temporary .pptx files in the folder
        self.presentations = sorted(
            [
                os.path.join(self.folder_path, f)
                for f in os.listdir(self.folder_path)
                if f.endswith('.pptx') and not f.startswith('~')
            ]
        )
        print(f"Loaded {len(self.presentations)} presentations.")

    def open_presentation(self, index):
        if 0 <= index < len(self.presentations):
            file_path = self.presentations[index]
            self.current_presentation = self.ppt_app.Presentations.Open(file_path)
            time.sleep(1)  # Allow time to fully open
            print(f"Opened presentation: {file_path}")

    def run_slideshow(self):
        if self.current_presentation:
            self.current_presentation.SlideShowSettings.Run()
            print("Started slideshow")

    def close_presentation(self):
        if self.current_presentation:
            self.current_presentation.Close()
            self.current_presentation = None 
            print("Presentation closed.")

    def show_next_popup(self, next_presentation_title):
        # Create a simple Tkinter window to wait for permission
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        root.attributes("-topmost", True)  # Ensure the window is always on top
        result = messagebox.askyesno("Next Presentation", f"Do you want to proceed to '{next_presentation_title}'?")
        root.destroy()
        return result

    def advance_to_next_presentation(self):
        if self.current_presentation:
            try:
                if self.ppt_app.SlideShowWindows.Count > 0:
                    slide_show_window = self.ppt_app.SlideShowWindows(1)
                    if slide_show_window.View.CurrentShowPosition == self.current_presentation.Slides.Count + 1:
                        self.close_presentation()  # Close the current presentation
                        time.sleep(1)
                        if self.current_index < len(self.presentations) - 1:
                            next_presentation_title = os.path.basename(self.presentations[self.current_index + 1])
                            self.current_index += 1  # Increment index only after confirmation
                            self.open_presentation(self.current_index)
                            if self.show_next_popup(next_presentation_title):
                                self.run_slideshow()
                            else:
                                print("Presentation halted by user.")
                                self.close_program()
                        else:
                            print("No more presentations left to play.")
                            self.close_program()
                else:
                    print("No active slideshow window found.")
                    self.close_program()
            except Exception as e:
                print(f"An error occurred: {e}")

    def close_program(self):
        self.ppt_app.Quit()
        print("Program closed.")
        exit()
        
    def run(self):
        if self.presentations:
            self.open_presentation(self.current_index)
            self.run_slideshow()
            while True:
                time.sleep(1)
                self.advance_to_next_presentation()
                if self.current_index >= len(self.presentations):
                    self.close_program()
                    break

class IgnoreTempFilesHandler(FileSystemEventHandler):
    def dispatch(self, event):
        if event.src_path.startswith('~$'):
            return
        super().dispatch(event)

class FolderSyncHandler(FileSystemEventHandler):
    def __init__(self, power_point_handler):
        self.power_point_handler = power_point_handler

    def on_modified(self, event):
        filename = os.path.basename(event.src_path)
        if not filename.endswith('.pptx') or filename.startswith('~$'):
            return
        print(f"Detected modification in {event.src_path}")
        time.sleep(0.5)  # Allow time for the file to be fully saved
        self.power_point_handler.load_presentations()

def main():
    folder_path = os.path.join(os.path.dirname(__file__), "test")
    ppt_handler = PowerPointHandler(folder_path)
    
    event_handler = FolderSyncHandler(ppt_handler)
    observer = Observer()
    observer.schedule(event_handler, folder_path, recursive=False)
    observer.start()

    try:
        ppt_handler.run()
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
