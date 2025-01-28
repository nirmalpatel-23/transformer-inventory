import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from utils.helpers import on_division_select
from config.sheets_setup import save_to_sheets, setup_google_sheets

def create_add_form(sheet):
    add_window = tk.Toplevel()
    add_window.title("Add Data")
    add_window.geometry("500x520")  # Adjusted height to fit all elements
    
    # Create main frame with padding
    main_frame = ttk.Frame(add_window, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title Label
    title_label = ttk.Label(
        main_frame, 
        text="Add New Transformer", 
        font=("Arial", 14, "bold")
    )
    title_label.pack(pady=(0, 20))

    # Division Frame
    division_frame = ttk.LabelFrame(main_frame, text="Division", padding="10")
    division_frame.pack(fill=tk.X, pady=(0, 15))

    selected_division = tk.StringVar(value="")
    divisions = ["MODASA", "TALOD", "BHILODA"]

    # Create division radio buttons with better spacing
    radio_frame = ttk.Frame(division_frame)
    radio_frame.pack(pady=5)
    
    for col_index, division in enumerate(divisions):
        rb = ttk.Radiobutton(
            radio_frame,
            text=division,
            variable=selected_division,
            value=division,
            padding="5"
        )
        rb.pack(side=tk.LEFT, padx=20)

    # Details Frame
    details_frame = ttk.LabelFrame(main_frame, text="Transformer Details", padding="10")
    details_frame.pack(fill=tk.X, pady=(0, 15))  # Reduced bottom padding

    fields = ["Truck No", "MR No", "Lot No", "Date", "Total TC"]
    entries = []

    # Create fields with better layout
    for i, field in enumerate(fields):
        field_frame = ttk.Frame(details_frame)
        field_frame.pack(fill=tk.X, pady=5)
        
        label = ttk.Label(field_frame, text=field, width=15)
        label.pack(side=tk.LEFT, padx=(10, 5))
        
        if field == "Date":
            entry = DateEntry(
                field_frame,
                width=30,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                showweeknumbers=False,
                firstweekday='sunday',
                showothermonthdays=False,
                selectmode='day',
                font=("Arial", 10)
            )
        else:
            entry = ttk.Entry(field_frame, width=32)
            
            if field == "Lot No":
                entry.configure(state='readonly')
        
        entry.pack(side=tk.LEFT, padx=(0, 10))
        entries.append(entry)
        
        if field == "Lot No":
            selected_division.trace('w', lambda *args, e=entry: 
                                  on_division_select(None, selected_division, e, sheet))

    # Button Frame
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=10)
    
    # Submit Button
    submit_button = ttk.Button(
        button_frame,
        text="Submit",
        command=lambda: on_add_submit(sheet, entries, selected_division),
        width=15
    )
    submit_button.pack(side=tk.LEFT, padx=10)
    
    # Close Form Button
    close_button = ttk.Button(
        button_frame,
        text="Close Form",
        command=add_window.destroy,
        width=15
    )
    close_button.pack(side=tk.LEFT, padx=10)

    # Center the window on the screen
    add_window.update_idletasks()
    width = add_window.winfo_width()
    height = add_window.winfo_height()
    x = (add_window.winfo_screenwidth() // 2) - (width // 2)
    y = (add_window.winfo_screenheight() // 2) - (height // 2)
    add_window.geometry(f'{width}x{height}+{x}+{y}')

def on_add_submit(sheet, entries, selected_division):
    division = selected_division.get()
    if not division:
        messagebox.showerror("Error", "Please select a division!")
        return

    data = [entry.get() for entry in entries]
    if not all(data):
        messagebox.showerror("Error", "All fields are required!")
        return

    truck_no, mr_no, lot_no, date, total_tc = data

    try:
        total_tc = int(total_tc)
    except ValueError:
        messagebox.showerror("Error", "Total TC must be a number!")
        return

    first_page_data = [division, truck_no, mr_no, lot_no, date]
    
    # Save to both Master Sheet and Testing Sheet
    try:
        # Save to Master Sheet
        rows_to_add = [first_page_data for _ in range(total_tc)]
        if save_to_sheets(rows_to_add, "add_form"):
            # # Save to Testing Sheet
            # testing_sheet = setup_google_sheets().worksheet("TESTING")
            # for _ in range(total_tc):
            #     testing_sheet.append_row(first_page_data)
            
            # Create second form
            from .second_form import create_second_form
            create_second_form(sheet, total_tc, first_page_data)
        else:
            messagebox.showerror("Error", "Failed to save data to Master Sheet")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save data: {str(e)}") 