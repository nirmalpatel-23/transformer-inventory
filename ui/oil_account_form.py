import tkinter as tk
from tkinter import ttk

def create_oil_account_form(sheet):
    oil_window = tk.Toplevel()
    oil_window.title("Oil Account Form")
    oil_window.geometry("600x400")

    heading_label = tk.Label(oil_window, text="Oil Account Details", font=("Arial", 16, "bold"))
    heading_label.pack(pady=20)

    form_frame = ttk.Frame(oil_window, padding=20)
    form_frame.pack(fill='both', expand=True)

    fields = ["Date", "Oil Type", "Quantity (L)", "Purpose", "Location"]
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
        command=lambda: handle_oil_account_submit(sheet, entries)
    )
    submit_button.grid(row=len(fields), column=0, columnspan=2, pady=20)

def handle_oil_account_submit(sheet, entries):
    # Add your submit logic here
    pass 