import tkinter as tk
from tkinter import ttk

def create_certificates_form(sheet):
    cert_window = tk.Toplevel()
    cert_window.title("Certificates Form")
    cert_window.geometry("600x400")

    heading_label = tk.Label(cert_window, text="Certificate Details", font=("Arial", 16, "bold"))
    heading_label.pack(pady=20)

    form_frame = ttk.Frame(cert_window, padding=20)
    form_frame.pack(fill='both', expand=True)

    fields = ["Certificate Type", "Issue Date", "Expiry Date", "Certificate Number", "Issuing Authority"]
    entries = []

    for i, field in enumerate(fields):
        label = ttk.Label(form_frame, text=field)
        label.grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = ttk.Entry(form_frame)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries.append(entry)

    submit_button = ttk.Button(
        form_frame,
        text="Submit",
        command=lambda: handle_certificates_submit(sheet, entries)
    )
    submit_button.grid(row=len(fields), column=0, columnspan=2, pady=20)

def handle_certificates_submit(sheet, entries):
    # Add your submit logic here
    pass 

def show_details(self):
    selected_items = self.tree.selection()
    if not selected_items:
        print("No item selected.")
        return  # Exit the function if no item is selected

    selected_item = selected_items[0]
    # Proceed with your logic using 'selected_item'
    # For example, retrieve details from the treeview
    item_details = self.tree.item(selected_item)
    print("Selected item details:", item_details) 

def edit_record(self):
    selected_items = self.tree.selection()
    if not selected_items:
        print("No item selected.")
        return  # Exit the function if no item is selected
    item = selected_items[0]
    # Proceed with your logic using 'item'
    # For example, retrieve details from the treeview
    item_details = self.tree.item(item)
    print("Selected item details:", item_details) 