import tkinter as tk
from tkinter import ttk, messagebox

def create_search_form(sheet):
    search_window = tk.Toplevel()
    search_window.title("Search Data")
    search_window.geometry("1200x600")

    main_container = ttk.Frame(search_window, padding="20")
    main_container.pack(fill=tk.BOTH, expand=True)

    header_label = ttk.Label(
        main_container,
        text="SEARCH BAR",
        font=("Arial", 12, "bold"),
        anchor="center"
    )
    header_label.pack(pady=(0, 10))

    instruction_label = ttk.Label(
        main_container,
        text="*SEARCH BY MR NO. OR DATE",
        anchor="w"
    )
    instruction_label.pack(fill=tk.X, pady=(0, 10))
    # Search fields frame
    search_fields_frame = ttk.Frame(main_container)
    search_fields_frame.pack(fill=tk.X, pady=(0, 20))

    # MR NO and DATE entry fields
    mr_no_label = ttk.Label(search_fields_frame, text="MR NO.:-")
    mr_no_label.pack(side=tk.LEFT, padx=(0, 10))
    mr_no_entry = ttk.Entry(search_fields_frame, width=30)
    mr_no_entry.pack(side=tk.LEFT, padx=(0, 20))

    date_label = ttk.Label(search_fields_frame, text="DATE:-")
    date_label.pack(side=tk.LEFT, padx=(0, 10))
    date_entry = ttk.Entry(search_fields_frame, width=30)
    date_entry.pack(side=tk.LEFT)

    # Create the table
    tree_frame = ttk.Frame(main_container)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    # Create Treeview with scrollbar
    tree_scroll = ttk.Scrollbar(tree_frame)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    tree = ttk.Treeview(
        tree_frame,
        columns=("SRNO", "DIVISION", "TRUCK NO", "MR NO", "LOT NO", "DATE", 
                "TC NO", "MAKE", "TC CAPACITY", "JOB NO", "EDIT"),
        show="headings",
        yscrollcommand=tree_scroll.set
    )
    tree_scroll.config(command=tree.yview)

    # Define column headings and widths
    columns = [
        ("SRNO", 50),
        ("DIVISION", 100),
        ("TRUCK NO", 100),
        ("MR NO", 100),
        ("LOT NO", 100),
        ("DATE", 100),
        ("TC NO", 100),
        ("MAKE", 100),
        ("TC CAPACITY", 100),
        ("JOB NO", 100),
        ("EDIT", 50)
    ]

    for col, width in columns:
        tree.heading(col, text=col)
        tree.column(col, width=width, anchor=tk.CENTER)

    tree.pack(fill=tk.BOTH, expand=True)

    def search_data():
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)

        mr_no = mr_no_entry.get()
        date = date_entry.get()

        try:
            # Get all data from sheet
            data = sheet.get_all_values()
            if len(data) > 0:
                headers = data[0]  # Skip header row
                data = data[1:]

            # Filter based on search criteria
            filtered_data = []
            for row in data:
                if (mr_no and mr_no in row[2]) or (date and date in row[4]):  # Assuming MR NO is column 3 and DATE is column 5
                    filtered_data.append(row)

            # Insert filtered data into treeview
            for i, row in enumerate(filtered_data, 1):
                tree.insert("", tk.END, values=(i,) + tuple(row) + ("",))

        except Exception as e:
            messagebox.showerror("Error", f"Error searching data: {e}")

    def edit_record(event):
        item = tree.selection()[0]
        # Get the values of the selected item
        values = tree.item(item)['values']
        # Implement edit functionality here
        messagebox.showinfo("Edit", f"Editing record: {values}")

    # Bind double-click event to edit_record function
    tree.bind('<Double-1>', edit_record)

    # Search button
    search_button = ttk.Button(
        main_container,
        text="Search",
        command=search_data
    )
    search_button.pack(pady=10)

    # Configure tree tag for edit column
    tree.tag_configure('edit', foreground='blue')
# Function to create the first page form

    # Rest of the search form code...
    # (I've truncated this for brevity, but you should include the full search form implementation) 