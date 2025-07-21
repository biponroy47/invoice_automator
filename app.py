import io
import os
import tkinter as tk
from tkinter import filedialog
import fitz
import csv


selected_folder_path = ""
files_in_folder = []
pdf_text_content = []
invoice_data = []
extracted_data = []
csv_content = ""

def select_folder():
    global selected_folder_path, files_in_folder
    selected_folder_path = filedialog.askdirectory()
    if selected_folder_path:
        folder_label.config(text=f"Selected: {selected_folder_path}")
        files_in_folder = [
            os.path.join(selected_folder_path, f)
            for f in os.listdir(selected_folder_path)
            if os.path.isfile(os.path.join(selected_folder_path, f))
        ]
        update_file_listbox()
        read_all_pdfs()
        format_text_content()


def update_file_listbox():
    file_listbox.delete(0, tk.END)
    for file in files_in_folder:
        if os.path.basename(file).endswith(".pdf"):
            file_listbox.insert(tk.END, os.path.basename(file))


def print_invoices():
    global invoice_data
    for i, invoice in enumerate(invoice_data):
        print(f"\n--- Invoice {i} ---")
        for j, line in enumerate(invoice):
            print(f"[{j}] {line}")


def read_all_pdfs():
    global pdf_text_content
    pdf_text_content = []

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
    for pdf in pdf_text_content:
        invoice_data.append(pdf.split('\n'))

    print_invoices()

    for i, invoice in enumerate(invoice_data):
        data = ""
        for j in range (len(invoice)):
            if invoice[j] == "Invoice number :":
                j += 1
                while invoice[j] != "Invoice date :":
                    data += invoice[j] + " "
                    j += 1
                data += ","
            if invoice[j] == "Invoice date :":
                j += 1
                while invoice[j] != "Customer contract :":
                    data += invoice[j]  + " "
                    j += 1
                data += ","
            if invoice[j] == "Customer contract :":
                j += 1
                while invoice[j] != "PO number :":
                    data += invoice[j]  + " "
                    j += 1
                data += ","
            if invoice[j] == "PO number :":
                j += 1
                while invoice[j] != "Term :":
                    data += invoice[j]  + " "
                    j += 1
                data += ","
            if invoice[j] == "Term :":
                j += 1
                while invoice[j] != "Currency :":
                    data += invoice[j]  + " "
                    j += 1
                data += ","
            if invoice[j] == "Currency :":
                j += 1
                while invoice[j] != "TOTAL :":
                    data += invoice[j]  + " "
                    j += 1
                data += ","
            if invoice[j] == "TOTAL :":
                temp = invoice[j + 1].replace('$','').replace(',','')
                data += temp + ","
            while invoice[j] != "Project :":
                j += 1

            if invoice[j] == "Project :":
                data += invoice[j + 1]
                break

        global extracted_data
        extracted_data.append(data)

    generate_summary_button.pack(pady=5)


def calculate_total():
    total = 0
    # for invoice in invoice_data:
    #     for line in invoice:
    #         if "TOTAL :" in line:
    #             temp = float(line["TOTAL :"].replace('$','').replace(',',''))
    #             total += temp
    return total

def generate_summary():
    field_names = [
        "Invoice number :",
        "Invoice date :",
        "Customer contract :",
        "PO number :",
        "Term :",
        "Currency :",
        "TOTAL :",
        "Project :",
    ]

    headers = ""
    for field in field_names:
        headers += field + ","
    headers += "PL,Verified\n"

    global csv_content
    csv_content += headers

    for row in extracted_data:
        csv_content += row + "\n"
        
    # print(csv_content)

    status_label.config(text="Parsing complete!")

    save_button.pack(pady=20)
    return

def save_file():
    file_path = filedialog.asksaveasfilename(
        defaultextension = ".csv",
        filetypes        = [("CSV files","*.csv")],
        initialfile      = "invoices_summary",
        title            = "Save CSV as..."
    )
    
    if file_path:
        with open(file_path, "w", newline="") as f:
            f.write(csv_content)
            save_label.config(text=f"File saved!\nCSV saved to: {file_path}")


root = tk.Tk()
root.title("Monthly Invoice Automator")
root.geometry("350x650")

label = tk.Label(root, text="Choose the folder containing all invoices:", font=("Arial", 10))
label.pack(pady=10)

select_folder_button = tk.Button(root, text="Select Folder", command=select_folder, font=("Arial", 10))
select_folder_button.pack(pady=5)

folder_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
folder_label.pack(pady=5)

file_listbox_label = tk.Label(root, text="Selected Invoices:", font=("Arial", 9), wraplength=300, justify="left")
file_listbox_label.pack(pady=10)

file_listbox = tk.Listbox(root, width=40, height=15, font=("Arial", 9))
file_listbox.pack(pady=5)

generate_summary_button = tk.Button(root, text="Generate Summary", command=generate_summary, font=("Arial", 10))

status_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
status_label.pack(pady=5)

save_button = tk.Button(root, text="Save File", command=save_file, font=("Arial", 12))

save_label = tk.Label(root, text="", font=("Arial", 9), wraplength=300, justify="left")
save_label.pack(pady=5)

root.mainloop()

