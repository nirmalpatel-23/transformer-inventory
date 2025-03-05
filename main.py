import tkinter as tk
from tkinter import ttk
from config.sheets_setup import setup_google_sheets, SHEET_CONFIGS
from ui.add_form import create_add_form
from ui.search_form import create_search_form
from ui.physical_form import create_physical_form, PhysicalVerificationForm
from ui.create_physical_verification_form import create_physical_verification_form
from ui.internal_form import create_internal_form, InternalVerificationForm
from ui.testing_form import create_testing_form
from ui.final_bill import create_final_bill
from ui.oil_account_form import create_oil_account_form
from ui.certificates_form import create_certificates_form

def main():
    # Get the master worksheet specifically for the lot number functionality
    master_sheet = setup_google_sheets().worksheet("MASTER SHEET")
    if not master_sheet:
        return

    root = tk.Tk()
    root.title("Data Management")
    root.geometry("800x600")

    heading_frame = ttk.Frame(root)
    heading_frame.pack(pady=40)
    
    heading_label = tk.Label(
        heading_frame, 
        text="OMEGA ELECTRICALS", 
        font=("Arial", 24, "bold"), 
        fg="blue"
    )
    heading_label.pack()

    main_frame = ttk.Frame(root)
    main_frame.pack(expand=True)

    top_frame = ttk.Frame(main_frame)
    top_frame.pack(pady=40)

    add_button = ttk.Button(
        top_frame, 
        text="ADD DATA",
        command=lambda: create_add_form(master_sheet),
        width=20
    )
    add_button.pack(side=tk.LEFT, padx=20)

    search_button = ttk.Button(
        top_frame,
        text="SEARCH DATA",
        command=lambda: create_search_form(master_sheet),
        width=20
    )
    search_button.pack(side=tk.LEFT, padx=20)

    physical_button = ttk.Button(
        top_frame,
        text="PHYSICAL 2",
        command=lambda: create_physical_verification_form(master_sheet),
        width=20
    )
    physical_button.pack(side=tk.LEFT, padx=20)

    middle_frame = ttk.Frame(main_frame)
    middle_frame.pack(pady=40)

    middle_buttons = ["TESTING", "OIL ACCOUNT", "CERTIFICATES"]
    for text in middle_buttons:
        button = ttk.Button(
            middle_frame,
            text=text,
            command=lambda t=text, s=master_sheet: handle_button_click(t, s),
            width=20
        )
        button.pack(side=tk.LEFT, padx=20)

    bottom_frame = ttk.Frame(main_frame)
    bottom_frame.pack(pady=40)

    bottom_buttons = ["PHYSICAL", "INTERNAL", "FINAL BILL"]
    for text in bottom_buttons:
        button = ttk.Button(
            bottom_frame,
            text=text,
            command=lambda t=text, s=master_sheet: handle_button_click(t, s),
            width=20
        )
        button.pack(side=tk.LEFT, padx=20)

    root.mainloop()

def handle_button_click(button_type, sheet):
    if button_type == "PHYSICAL":
        PhysicalVerificationForm(sheet)
    elif button_type == "INTERNAL":
        InternalVerificationForm(sheet)
    elif button_type == "TESTING":
        create_testing_form(sheet)
    elif button_type == "OIL ACCOUNT":
        create_oil_account_form(sheet)
    elif button_type == "CERTIFICATES":
        create_certificates_form(sheet)
    elif button_type == "FINAL BILL":
        create_final_bill(sheet)

if __name__ == "__main__":
    main() 