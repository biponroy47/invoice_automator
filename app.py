import os
import tkinter as tk
from tkinter import filedialog

# Global variables
selected_folder_path = ""
files_in_folder = []

def select_folder():
    global selected_folder_path, files_in_folder
    selected_folder_path = filedialog.askdirectory()
    if selected_folder_path:
        folder_label.config(text=f"Selected: {selected_folder_path}")
        # Store all files in the selected folder
        files_in_folder = [
            os.path.join(selected_folder_path, f)
            for f in os.listdir(selected_folder_path)
            if os.path.isfile(os.path.join(selected_folder_path, f))
        ]
        update_file_listbox()

def update_file_listbox():
    file_listbox.delete(0, tk.END)  # Clear existing entries
    for file in files_in_folder:
        file_listbox.insert(tk.END, os.path.basename(file))
    print(files_in_folder)

root = tk.Tk()
root.title("Monthly Invoice Automator")
root.geometry("400x400")

label = tk.Label(root, text="Choose the folder containing all invoices:", font=("Arial", 10))
label.pack(pady=10)

button = tk.Button(root, text="Select Folder", command=select_folder, font=("Arial", 10))
button.pack(pady=5)

folder_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
folder_label.pack(pady=5)

file_listbox = tk.Listbox(root, width=40, height=10, font=("Arial", 9))
file_listbox.pack(pady=5)

root.mainloop()
