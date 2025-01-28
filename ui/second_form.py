import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import save_to_sheets, setup_google_sheets

def create_second_form(sheet, total_tc, first_page_data):
    # Get the current values from both sheets
    all_values = sheet.get_all_values()
    start_row = len(all_values) - total_tc + 1
    
    # Get Testing sheet reference
    testing_sheet = setup_google_sheets().worksheet("TESTING")
    testing_values = testing_sheet.get_all_values()
    testing_start_row = len(testing_values) - total_tc + 1

    second_window = tk.Toplevel()
    second_window.title("Additional TC Details")
    second_window.geometry("1200x800")
    
     # Create main container frame
    container = ttk.Frame(second_window)
    container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    # Add title label
    title_label = ttk.Label(
        container,
        text="Enter details for each TC:",
        font=("Arial", 11, "bold")
    )
    title_label.pack(pady=(0, 20))

    # Create canvas and scrollbar for scrolling
    canvas = tk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)

    # Create frame inside canvas to hold TC forms
    scrollable_frame = ttk.Frame(canvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    # Configure canvas
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_x.set)

    # Pack scrollbar and canvas
    scrollbar.pack(side="right", fill="y")
    scrollbar_x.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)

    # Configure grid for TC forms
    max_columns = 5  # Maximum number of columns before wrapping
    row_num = 0
    col_num = 0
    sets = []

    # Create frames for each TC
    for i in range(total_tc):
        # Calculate grid position
        col_num = i % max_columns
        row_num = i // max_columns

        # Create frame for individual TC
        tc_frame = ttk.LabelFrame(
            scrollable_frame,
            text=f"DETAILS OF TC-{i + 1}",
            padding="10"
        )
        tc_frame.grid(row=row_num, column=col_num, padx=10, pady=10, sticky="nsew")

        # Create form fields
        fields = [
            ("TC NO :-", "tc_no"),
            ("MAKE :-", "make"),
            ("TC CAPACTIY :-", "tc_capacity"),
            ("JOB NO :-", "job_no")
        ]

        tc_entries = {}
        for idx, (label_text, field_name) in enumerate(fields):
            # Label
            label = ttk.Label(tc_frame, text=label_text)
            label.grid(row=idx, column=0, sticky="w", pady=5)
            
            # Entry
            entry = ttk.Entry(tc_frame, width=25)
            entry.grid(row=idx, column=1, sticky="ew", padx=(10, 0), pady=5)
            tc_entries[field_name] = entry

        # Add this set of entries to our sets list
        sets.append((
            tc_entries["tc_no"],
            tc_entries["make"],
            tc_entries["tc_capacity"],
            tc_entries["job_no"]
        ))

    # Define submit function before creating buttons
    def on_second_submit():
        second_page_data = []
        testing_data = []
        
        for set_entries in sets:
            print(set_entries)
            set_data = [entry.get() for entry in set_entries]
            if not all(set_data):
                messagebox.showerror("Error", "All fields are required!")
                return
            
            # Prepare data for both sheets
            second_page_data.append(set_data)
            testing_row = first_page_data + set_data  # Combine first page and second page data
            testing_data.append(testing_row)

        try:
            # Update Master Sheet
            additional_info = {
                "start_row": start_row
            }
            if save_to_sheets(second_page_data, "second_form", additional_info):
                # Update Testing Sheet
                # for idx, row_data in enumerate(testing_data):
                #     row_num = testing_start_row + idx
                #     range_name = f'A{row_num}:W{row_num}'  # Adjust column range as needed
                #     testing_sheet.update(range_name, [row_data])
                
                messagebox.showinfo("Success", "Data added successfully!")
                second_window.destroy()
            else:
                messagebox.showerror("Error", "Failed to save data to Master Sheet")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {str(e)}")
            print(f"Debug - Save error details: {str(e)}")

    # Button frame
    button_frame = ttk.Frame(container)
    button_frame.pack(pady=20)
    
    # Submit button
    submit_button = ttk.Button(
        button_frame,
        text="Submit",
        command=on_second_submit,  # Now this function exists when referenced
        width=15
    )
    submit_button.pack(side=tk.LEFT, padx=10)
    
    # Close Form button
    close_button = ttk.Button(
        button_frame,
        text="Close Form",
        command=second_window.destroy,
        width=15
    )
    close_button.pack(side=tk.LEFT, padx=10)

    # Configure mouse wheel scrolling
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Configure horizontal scrolling with Shift + Mouse wheel
    def _on_shift_mousewheel(event):
        canvas.xview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)

    # Make the window resizable
    second_window.resizable(True, True)

# Function to handle Add Data button click on the first page
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
        total_tc = int(total_tc)  # Convert Total TC to an integer
    except ValueError:
        messagebox.showerror("Error", "Total TC must be a number!")
        return

    # Prepare data for the first page
    first_page_data = [division, truck_no, mr_no, lot_no, date]

    # Add the first page data to Google Sheets
    rows_to_add = [first_page_data for _ in range(total_tc)]
    try:
        for row in rows_to_add:
            sheet.append_row(row)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to add data: {e}")
        return

    # Move to the second page
    create_second_form(sheet, total_tc, first_page_data)

    # Rest of the second form code...
    # (I've truncated this for brevity, but you should include the full second form implementation) 