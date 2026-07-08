import os
import tkinter as tk
from tkinter import filedialog, scrolledtext

def select_directory():
    directory_path = filedialog.askdirectory()
    if directory_path:
        display_filenames(directory_path)

def display_filenames(directory_path):
    items = os.listdir(directory_path)
    files = [f for f in items if os.path.isfile(os.path.join(directory_path, f))]
    folders = [f for f in items if os.path.isdir(os.path.join(directory_path, f))]

    text_box.insert(tk.END, f"---------- {os.path.basename(directory_path)} ----------\n")
    text_box.insert(tk.END, f"Path: {directory_path}\n\n")
    
    if folders:
        text_box.insert(tk.END, "Folders:\n")
        for folder in folders:
            text_box.insert(tk.END, folder + '\n')
    
    if files:
        text_box.insert(tk.END, "\nFiles:\n")
        for file in files:
            text_box.insert(tk.END, file + '\n')
    
    total_folders = len(folders)
    total_files = len(files)
    total_items = total_folders + total_files
    
    text_box.insert(tk.END, '\n')
    text_box.insert(tk.END, f"Total Folders: {total_folders}\n")
    text_box.insert(tk.END, f"Total Files: {total_files}\n")
    text_box.insert(tk.END, f"Total Items: {total_items}\n")
    text_box.insert(tk.END, '\n')

def clear_text():
    text_box.delete(1.0, tk.END)

# Set up the main application window
root = tk.Tk()
root.title("Directory File Lister")

# Create a frame for the buttons
frame = tk.Frame(root)
frame.pack(pady=10)

# Add a button to select directory
select_button = tk.Button(frame, text="Select Directory", command=select_directory)
select_button.pack(side=tk.LEFT, padx=10)

# Add a button to clear the text box
clear_button = tk.Button(frame, text="Clear", command=clear_text)
clear_button.pack(side=tk.LEFT, padx=10)

# Add a text box to display file names
text_box = scrolledtext.ScrolledText(root, wrap=tk.WORD)
text_box.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

# Run the application
root.mainloop()
