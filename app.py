import io
import os
import tkinter as tk
from tkinter import filedialog
import fitz  # PyMuPDF
import csv


# Global variables
selected_folder_path = ""
files_in_folder = []
pdf_text_content = []
invoice_data = []
desired_field_single = ["Invoice number :", "Invoice date :", "Term :", "Currency :", "TOTAL :", ]
desired_field_double = ["Customer contract :", "PO number :"]
csv_content = ""

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
        read_all_pdfs()
        format_text_content()


def update_file_listbox():
    file_listbox.delete(0, tk.END)  # Clear existing entries
    for file in files_in_folder:
        file_listbox.insert(tk.END, os.path.basename(file))


def print_invoices():
    global invoice_data
    for i, invoice in enumerate(invoice_data):
        print(f"\n--- Invoice {i} ---")
        for j, line in enumerate(invoice):
            print(f"[{j}] {line}")


def read_all_pdfs():
    global pdf_text_content
    pdf_text_content = []  # Clear previous content

    pdf_files = [f for f in files_in_folder if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("No PDF files found.")
        return

    for pdf_file in pdf_files:
        try:
            with fitz.open(pdf_file) as doc:
                file_text = ""
                for page in doc:
                    file_text += page.get_text()
                pdf_text_content.append(file_text)
        except Exception as e:
            print(f"Error reading {pdf_file}: {e}")


def format_text_content():
    global invoice_data
    invoice_data = []  # Clear previous content

    for pdf in pdf_text_content:
        invoice_data.append(pdf.split('\n'))

    for i, invoice in enumerate(invoice_data):
        extracted_data = []
        for j, line in enumerate(invoice):
            if line in desired_field_single :
                extracted_data.append({line: invoice[j + 1]})
            elif line in desired_field_double :
                extracted_data.append({line: invoice[j + 1] + invoice[j + 2]})
        invoice_data[i] = extracted_data
    
    # print(print_invoices())
    generate_summary_button.pack(pady=5)


# def generate_summary():
#     field_names = [desired_field_single[0], desired_field_single[1], desired_field_single[2], desired_field_single[3], desired_field_double[0], desired_field_double[1], desired_field_single[4]]
#     buffer = io.StringIO()
#     writer = csv.DictWriter(buffer, fieldnames=field_names)
#     writer.writeheader()
#     for invoice in invoice_data:
#         status_label.config(text=f"Parsing Invoice_{invoice[0].values()}")
#         status_label.pack(pady=5)
#         row = {}
#         for pair in invoice:
#             if isinstance(pair, dict):
#                 row.update(pair)
#             else:
#                 k, v = pair
#                 row[k] = v
#         writer.writerow(row)

#     csv_content = buffer.getvalue()
#     buffer.close()
#     print(csv_content)

def generate_summary():
    # 1) Prepare header and buffer
    field_names = [
        desired_field_single[0],
        desired_field_single[1],
        desired_field_single[2],
        desired_field_single[3],
        desired_field_double[0],
        desired_field_double[1],
        desired_field_single[4]
    ]
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=field_names)
    writer.writeheader()

    # 2) Make a shallow copy of invoice_data so we can pop from front
    invoices = list(invoice_data)

    # 3) Define the per‐invoice processor
    def process_next():
        if not invoices:
            # All done → finalize
            global csv_content
            csv_content = buffer.getvalue()
            buffer.close()
            status_label.config(text="Parsing complete!")
            # e.g. now call your save function or print:
            print(csv_content)
            save_button.pack(pady=20)
            return

        invoice = invoices.pop(0)

        # Update status immediately
        status_label.config(text=f"Parsing Invoice_{next(iter(invoice[0].values()))!r}.pdf ...")

        # Merge your invoice list of pairs/dicts into one dict row
        row = {}
        for pair in invoice:
            if isinstance(pair, dict):
                row.update(pair)
            else:
                k, v = pair
                row[k] = v

        writer.writerow(row)

        # Schedule the next invoice after 1000 ms (1 s)
        root.after(500, process_next)

    # 4) Kick off the process
    process_next()

def save_file():
    # 2) Show Save As dialog
    file_path = filedialog.asksaveasfilename(
        defaultextension = ".csv",
        filetypes        = [("CSV files","*.csv")],
        initialfile      = "transactions",
        title            = "Save CSV as..."
    )
    
    # 3) If user didn't cancel, write out the content
    if file_path:
        with open(file_path, "w", newline="") as f:
            f.write(csv_content)
            save_label.config(text=f"File saved!\nCSV saved to: {file_path}")
        # print(f"CSV saved to: {file_path}")
    # else:
    #     print("Save cancelled")


root = tk.Tk()
root.title("Monthly Invoice Automator")
root.geometry("400x800")

label = tk.Label(root, text="Choose the folder containing all invoices:", font=("Arial", 10))
label.pack(pady=10)

select_folder_button = tk.Button(root, text="Select Folder", command=select_folder, font=("Arial", 10))
select_folder_button.pack(pady=5)

folder_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
folder_label.pack(pady=5)

file_listbox_label = tk.Label(root, text="Selected Invoices:", font=("Arial", 9), wraplength=300, justify="left")
file_listbox_label.pack(pady=5)

file_listbox = tk.Listbox(root, width=40, height=15, font=("Arial", 9))
file_listbox.pack(pady=5)

generate_summary_button = tk.Button(root, text="Generate Summary", command=generate_summary, font=("Arial", 10))

status_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
status_label.pack(pady=5)

save_button = tk.Button(root, text="Save File", command=save_file, font=("Arial", 12))

save_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
save_label.pack(pady=5)

root.mainloop()

