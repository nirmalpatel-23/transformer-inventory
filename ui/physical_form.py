import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import save_to_sheets, get_worksheet, setup_google_sheets
from tkcalendar import DateEntry
import pandas as pd

def setup_styles():
    style = ttk.Style()
    style.configure(
        'Enrolled.TButton',
        background='grey',
        foreground='white'
    )

class PhysicalVerificationForm:
    def __init__(self, sheet):
        self.window = tk.Toplevel()
        self.window.title("PHYSICAL FORM")
        self.window.geometry("1200x800")
        
        # Initialize buttons list
        self.buttons = []
        
        # Store the master sheet reference
        self.master_sheet = sheet
        
        # Create main container with scrollbar
        self.main_container = ttk.Frame(self.window)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Add title
        title_label = ttk.Label(
            self.main_container, 
            text="PHYSICAL FORM", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Add search section between title and table
        self.create_top_search_section()
        
        # Create the table frame
        self.create_data_table()
        
        # Create separator
        ttk.Separator(self.main_container, orient='horizontal').pack(
            fill='x', pady=20
        )
        
        # Rest of the existing form components...
        # Create canvas and scrollbar for the form section
        self.create_form_section()
        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def create_search_section(self):
        search_frame = ttk.LabelFrame(self.scrollable_frame, text="Search Transformer", padding=10)
        search_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Search fields in one row
        ttk.Label(search_frame, text="TC No.:").grid(row=0, column=0, padx=5, pady=5)
        self.tc_search = ttk.Entry(search_frame, width=30)
        self.tc_search.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(search_frame, text="Job No.:").grid(row=0, column=2, padx=5, pady=5)
        self.job_search = ttk.Entry(search_frame, width=30)
        self.job_search.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Button(search_frame, text="Search", command=self.search_tc).grid(row=0, column=4, padx=20, pady=5)
        
        # Add a Close button
        ttk.Button(
            search_frame, 
            text="Close Form", 
            command=self.window.destroy
        ).grid(row=0, column=5, padx=20, pady=5)

    def create_tc_details_section(self):
        self.details_frame = ttk.LabelFrame(self.scrollable_frame, text="Transformer Details", padding=10)
        self.details_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Create fields for TC details in a grid layout
        self.detail_fields = [
            "DIVISION", "TRUCK NO", "MR NO",
            "LOT NO", "DATE", "TC NO",
            "MAKE", "TC CAPACITY", "JOB NO"
        ]
        self.detail_entries = {}
        
        # Create a 3x3 grid layout
        for i, field in enumerate(self.detail_fields):
            row = i // 3
            col = i % 3
            
            # Create a frame for each field
            field_frame = ttk.Frame(self.details_frame)
            field_frame.grid(row=row, column=col, padx=10, pady=5, sticky="nsew")
            
            ttk.Label(field_frame, text=field).pack(anchor="w")
            entry = ttk.Entry(field_frame, state='readonly', width=30)
            entry.pack(fill="x", pady=(2, 0))
            self.detail_entries[field] = entry

    def create_inspection_form(self):
        inspection_frame = ttk.LabelFrame(self.scrollable_frame, text="Physical Inspection Details", padding=10)
        inspection_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Add Date Frame at the top
        date_frame = ttk.Frame(inspection_frame)
        date_frame.pack(fill=tk.X, pady=(0, 15), padx=10)
        
        ttk.Label(date_frame, text="Physical Date:").pack(side=tk.LEFT, padx=(0, 10))
        self.physical_date = DateEntry(
            date_frame,
            width=20,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy',
            font=("Arial", 10)
        )
        self.physical_date.pack(side=tk.LEFT)
        
        left_frame = ttk.Frame(inspection_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=10)
        
        right_frame = ttk.Frame(inspection_frame)
        right_frame.pack(side="left", fill="both", expand=True, padx=10)
        
        # Updated sections with unique keys
        sections = {
            "HT SIDE": [
                ("HT_BUSHING", "BUSHING"), 
                ("HT_METAL_PART", "METAL PART"), 
                ("HT_CONNECTOR", "HT CONNECTOR")
            ],
            "LT SIDE": [
                ("LT_BUSHING", "BUSHING"), 
                ("LT_METAL_PART", "METAL PART"), 
                ("LT_CONNECTOR", "LT CONNECTOR")
            ],
            "OIL SIDE": [
                ("GAUGE_GLASS", "GAUGE GLASS"),
                ("OIL_AS_PER_NP", "OIL AS PER NP"),
                ("OIL_POSITION", "OIL POSITION"),
                ("OUTSIDE_PAINT", "OUTSIDE PAINT")
            ],
            "OTHER DETAILS": [
                ("BOLT_NUTS", "BOLT NUTS"),
                ("ROD_GASKET", "ROD GASKET"),
                ("TOP_GASKET", "TOP GASKET"),
                ("NAME_PLATE", "NAME PLATE"),
                ("BREATHER", "BREATHER"),
                ("LABOUR_CHARGE", "LABOUR CHARGE"),
                ("BS", "B/S"),
                ("CONSERVATOR_TANK", "CONSERVATOR TANK (in Kg)"),
                ("RADIATORS", "RADIATORS"),
                ("REMARKS", "REMARKS")
            ]
        }
        
        self.inspection_entries = {}
        left_sections = ["HT SIDE", "LT SIDE"]
        
        for section in left_sections:
            self.create_section(left_frame, section, sections[section])
            
        right_sections = ["OIL SIDE", "OTHER DETAILS"]
        for section in right_sections:
            self.create_section(right_frame, section, sections[section])
        
        # Submit button at the bottom
        submit_frame = ttk.Frame(self.scrollable_frame)
        submit_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(
            submit_frame, 
            text="Submit Inspection", 
            command=self.submit_inspection,
            width=30
        ).pack(anchor="center")

    def create_section(self, parent, section_title, fields):
        section_frame = ttk.LabelFrame(parent, text=section_title, padding=10)
        section_frame.pack(fill="x", pady=5)
        
        for key, display_text in fields:  # Modified to handle tuples
            field_frame = ttk.Frame(section_frame)
            field_frame.pack(fill="x", pady=2)
            
            ttk.Label(field_frame, text=display_text).pack(side="left", padx=(0, 10))
            entry = ttk.Entry(field_frame, width=30)
            entry.pack(side="left", fill="x", expand=True)
            self.inspection_entries[key] = entry  # Use the unique key

    def search_tc(self):
        job_no = self.job_search.get().strip()

        if not job_no:
            messagebox.showerror("Error", "Please enter a Job No.")
            return

        try:
            # Search in Master Sheet
            master_sheet = setup_google_sheets().worksheet("MASTER")
            master_data = master_sheet.get_all_values()

            found = False
            for row in master_data:
                if len(row) > 8 and row[8] == job_no:  # Assuming Job No is in column I (index 8)
                    found = True
                    # Update detail entries with the found data
                    for idx, field in enumerate(self.detail_fields):
                        entry = self.detail_entries[field]
                        entry.configure(state='normal')
                        entry.delete(0, tk.END)
                        entry.insert(0, row[idx])
                        entry.configure(state='readonly')
                    break

            if not found:
                messagebox.showwarning("Not Found", "No transformer found with matching Job No.")

        except Exception as e:
            messagebox.showerror("Error", f"Error searching transformer: {str(e)}")
            print(f"Debug - Error details: {str(e)}")

    def submit_inspection(self):
        # Validate physical date
        physical_date = self.physical_date.get()
        if not physical_date:
            messagebox.showerror("Error", "Please select Physical Date")
            return
        
        # Validate all inspection entries
        inspection_data = {}
        for field, entry in self.inspection_entries.items():
            value = entry.get().strip()
            if not value:
                messagebox.showerror("Error", f"Please fill in {field}")
                return
            inspection_data[field] = value
        
        # Get TC NO for row identification
        tc_no = self.detail_entries["TC NO"].get()
        if not tc_no:
            messagebox.showerror("Error", "Please search for a transformer first")
            return
        
        try:
            # Get both sheets
            spreadsheet = setup_google_sheets()
            master_sheet = self.master_sheet
            physical_sheet = spreadsheet.worksheet("PHYSICAL")
            
            # Update Master Sheet
            master_row = None
            master_data = master_sheet.get_all_values()
            
            for idx, row in enumerate(master_data):
                if len(row) > 5 and row[5] == tc_no:  # Column F (index 5)
                    master_row = idx + 1  # 1-based row index for Google Sheets
                    break
            
            if master_row is None:
                messagebox.showerror("Error", "Could not find the transformer row to update in Master Sheet")
                return
            
            # Check if TC NO exists in Physical Sheet
            physical_data = physical_sheet.get_all_values()
            physical_row = None
            
            for idx, row in enumerate(physical_data):
                if len(row) > 5 and row[5] == tc_no:  # Column F (index 5) for TC NO
                    physical_row = idx + 1  # 1-based row index
                    break
            
            # Updated inspection fields order with physical date
            inspection_fields = [
                "HT_BUSHING", "HT_METAL_PART", "HT_CONNECTOR",
                "LT_BUSHING", "LT_METAL_PART", "LT_CONNECTOR",
                "GAUGE_GLASS", "OIL_AS_PER_NP", "OIL_POSITION", "OUTSIDE_PAINT",
                "BOLT_NUTS", "ROD_GASKET", "TOP_GASKET", "NAME_PLATE",
                "BREATHER", "LABOUR_CHARGE", "BS", "CONSERVATOR_TANK", "RADIATORS", "REMARKS", "PHYSICAL_DATE"
            ]
            
            update_data = [inspection_data[field] for field in inspection_fields[:-1]] + [physical_date]
            
            # Update Master Sheet (columns J to AD)
            master_range = f'J{master_row}:AD{master_row}'
            master_sheet.update(master_range, [update_data])
            
            # Prepare complete data row for Physical sheet
            physical_data = [
                self.detail_entries["DIVISION"].get(),
                self.detail_entries["TRUCK NO"].get(),
                self.detail_entries["MR NO"].get(),
                self.detail_entries["LOT NO"].get(),
                self.detail_entries["DATE"].get(),
                self.detail_entries["TC NO"].get(),
                self.detail_entries["MAKE"].get(),
                self.detail_entries["TC CAPACITY"].get(),
                self.detail_entries["JOB NO"].get(),
                *update_data  # Add all inspection data including date
            ]
            
            if physical_row:
                # Calculate column range for Physical sheet
                num_columns = len(physical_data)
                if num_columns <= 26:
                    last_col = chr(ord('A') + num_columns - 1)
                else:
                    # For columns beyond Z
                    q, r = divmod(num_columns - 1, 26)
                    last_col = chr(ord('A') + q - 1) + chr(ord('A') + r)
                
                # Update existing row in Physical Sheet
                physical_range = f'A{physical_row}:{last_col}{physical_row}'
                try:
                    physical_sheet.update(physical_range, [physical_data])
                    message = "Physical inspection data updated successfully in both sheets!"
                except Exception as e:
                    print(f"Debug - Physical sheet update error: {str(e)}")
                    print(f"Debug - Range: {physical_range}")
                    print(f"Debug - Data length: {len(physical_data)}")
                    raise
            else:
                # Add new row if TC NO not found in Physical Sheet
                physical_sheet.append_row(physical_data)
                message = "Physical inspection data added successfully to both sheets!"
            
            messagebox.showinfo("Success", message)
            
            # Clear all entries
            for entry in self.inspection_entries.values():
                entry.delete(0, tk.END)
            
            self.tc_search.delete(0, tk.END)
            self.job_search.delete(0, tk.END)
            
            for entry in self.detail_entries.values():
                entry.configure(state='normal')
                entry.delete(0, tk.END)
                entry.configure(state='readonly')
            
            self.tc_search.focus()
            
            # After successful submission, reload the table
            self.load_table_data()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save inspection data: {str(e)}")
            print(f"Debug - Save error details: {str(e)}")

    def create_data_table(self):
        # Create frame for table
        table_frame = ttk.Frame(self.main_container)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Create Treeview with style
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)  # Adjusted row height
        
        # Create Treeview
        columns = (
            "DIVISION", "MR NO", "LOT NO", "DATE", "TC NO", 
            "MAKE", "TC CAPACITY", "JOB NO", "Physical Completed", "EXECUTE"
        )
        
        self.tree = ttk.Treeview(
            table_frame, 
            columns=columns, 
            show='headings', 
            height=10,
            style="Treeview"
        )
        
        # Configure column headings and widths
        column_widths = {
            "DIVISION": 100,
            "MR NO": 100,
            "LOT NO": 100,
            "DATE": 100,
            "TC NO": 100,
            "MAKE": 100,
            "TC CAPACITY": 100,
            "JOB NO": 100,
            "Physical Completed": 120,
            "EXECUTE": 100  # Width for buttons
        }
        
        # Configure columns
        for col in columns:
            self.tree.heading(col, text=col)
            width = column_widths.get(col, 100)
            self.tree.column(col, width=width, anchor='center', stretch=False)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack the table and scrollbar
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind events for button management
        self.tree.bind('<Map>', lambda e: self.load_table_data())
        self.tree.bind('<Destroy>', lambda e: self.cleanup_buttons())
        self.tree.bind('<<TreeviewSelect>>', self.update_button_positions)
        self.tree.bind('<Configure>', self.update_button_positions)
        
        # Bind mouse wheel and scrollbar events
        self.tree.bind('<MouseWheel>', self.update_button_positions)
        self.tree.bind('<Button-4>', self.update_button_positions)  # Linux scroll up
        self.tree.bind('<Button-5>', self.update_button_positions)  # Linux scroll down
        self.tree.bind('<<TreeviewOpen>>', self.update_button_positions)
        self.tree.bind('<<TreeviewClose>>', self.update_button_positions)

    def cleanup_buttons(self):
        # Clean up buttons when table is destroyed
        for button in self.buttons:
            button.destroy()
        self.buttons = []

    def load_table_data(self):
        try:
            # Clear existing items and buttons
            for button in self.buttons:
                button.destroy()
            self.buttons = []
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Get data from master sheet
            data = self.master_sheet.get_all_values()
            headers = data[0]  # Keep the header reference if needed
            
            # Find required column indices
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
                "TC NO": 5, "MAKE": 6, "TC CAPACITY": 7, "JOB NO": 8
            }
            
            # Process each row, starting from row 3 (index 2) to skip first two rows
            for row in data[2:]:  # Changed from data[1:] to data[2:]
                row_data = [row[required_cols[col]] for col in required_cols]
                
                # Check if physical inspection is completed
                tc_no = row[required_cols["TC NO"]]
                has_physical_data = any(row[9:28])  # Columns J to AB
                physical_status = "YES" if has_physical_data else "NO"
                row_data.append(physical_status)
                row_data.append("")  # Empty string for EXECUTE column
                
                # Add row to tree
                item = self.tree.insert('', tk.END, values=row_data)
                
                # Create button immediately
                self.create_button_for_row(item, tc_no, physical_status)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error loading table data: {str(e)}")
            print(f"Debug - Error details: {str(e)}")

    def create_button_for_row(self, item, tc_no, status):
        try:
            # Get button coordinates for the EXECUTE column
            bbox = self.tree.bbox(item, "EXECUTE")
            if not bbox:
                return
            
            # Debug print
            print(f"Creating button for TC No: {tc_no}, Status: {status}")
            
            # Create button with appropriate style and state
            if status == "NO":
                button = tk.Button(
                    self.tree,
                    text="Enrol",
                    bg='#FF0000',  # Red background
                    fg='white',
                    command=lambda tc=tc_no: self.start_enrollment(tc),  # Fixed command binding
                    width=8,
                    font=('Arial', 9, 'bold'),
                    relief="raised",
                    borderwidth=2
                )
            else:
                button = tk.Button(
                    self.tree,
                    text="Enrolled",
                    bg='#00FF00',  # Green background
                    fg='white',
                    state='disabled',
                    width=8,
                    font=('Arial', 9, 'bold'),
                    relief="raised",
                    borderwidth=2
                )
            
            # Calculate button position - center in the EXECUTE column
            x = bbox[0] + 10  # Slight offset from left edge of column
            y = bbox[1] + 2   # Slight offset from top of row
            
            # Place button
            button.place(x=x, y=y)
            
            # Store reference to prevent garbage collection
            self.buttons.append(button)
            
        except Exception as e:
            print(f"Error creating button: {str(e)}")

    def start_enrollment(self, tc_no):
        try:
            # Get row data from master sheet
            data = self.master_sheet.get_all_values()
            row_data = None
            
            # Skip header rows and find the matching TC NO
            for row in data[2:]:  # Skip first two rows
                if str(row[5]).strip() == str(tc_no).strip():  # TC NO is in column F (index 5)
                    row_data = row
                    break
            
            if row_data is None:
                messagebox.showerror("Error", f"TC No. {tc_no} not found in master sheet")
                return
            
            # Create new inspection window
            self.create_inspection_window(row_data)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error starting enrollment: {str(e)}")
            print(f"Debug - Start enrollment error: {str(e)}")
            print(f"Debug - TC No: {tc_no}")

    def create_inspection_window(self, row_data):
        # Create new window
        inspection_window = tk.Toplevel(self.window)
        inspection_window.title(f"Physical Inspection Form - TC No: {row_data[5]}")
        inspection_window.geometry("1000x800")
        
        # Configure window to be centered
        inspection_window.update_idletasks()
        width = inspection_window.winfo_width()
        height = inspection_window.winfo_height()
        x = (inspection_window.winfo_screenwidth() // 2) - (width // 2)
        y = (inspection_window.winfo_screenheight() // 2) - (height // 2)
        inspection_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Create main frame with scrollbar
        main_frame = ttk.Frame(inspection_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(main_frame, bg='white')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Main.TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Create styles
        style = ttk.Style()
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Section.TLabelframe', padding=15, relief='solid')
        style.configure('Main.TFrame', background='white')
        
        # TC Details Section with better styling
        details_frame = ttk.LabelFrame(
            scrollable_frame, 
            text="Transformer Details", 
            padding=15,
            style='Section.TLabelframe'
        )
        details_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        # Display TC details in grid
        tc_details = {
            "DIVISION": row_data[0],
            "MR NO": row_data[2],
            "DATE": row_data[4],
            "JOB NO": row_data[8],
            "TC NO": row_data[5]
        }
        
        for i, (label, value) in enumerate(tc_details.items()):
            ttk.Label(
                details_frame, 
                text=f"{label}:",
                font=("Arial", 10, 'bold')
            ).grid(row=i//3, column=(i%3)*2, padx=10, pady=8, sticky="e")
            
            entry = ttk.Entry(details_frame, width=20, font=("Arial", 10))
            entry.insert(0, value)
            entry.configure(state='readonly')
            entry.grid(row=i//3, column=(i%3)*2+1, padx=10, pady=8, sticky="w")
        
        # Add date field between TC Details and inspection sections
        date_frame = ttk.LabelFrame(
            scrollable_frame, 
            text="Inspection Date", 
            padding=15,
            style='Section.TLabelframe'
        )
        date_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        physical_date = DateEntry(
            date_frame,
            width=15,   
            background='darkblue',
            foreground='white',
            borderwidth=2,
            font=("Arial", 10),
            date_pattern='dd/mm/yyyy'
        )
        physical_date.pack(pady=5)
        
        # Create inspection sections frame
        inspection_frame = ttk.Frame(scrollable_frame)
        inspection_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        # Left Column (HT SIDE, LT SIDE, OIL SIDE)
        left_frame = ttk.Frame(inspection_frame)
        left_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 10))
        
        # Right Column (OTHER DETAILS)
        right_frame = ttk.Frame(inspection_frame)
        right_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(10, 0))
        
        # Store entries for later access
        entries = {}
        
        # Create sections with unique field names
        sections = {
            "left": {
                "HT SIDE": [
                    ("HT_BUSHING", "BUSHING"),
                    ("HT_METAL_PART", "METAL PART"),
                    ("HT_CONNECTOR", "HT CONNECTOR")
                ],
                "LT SIDE": [
                    ("LT_BUSHING", "BUSHING"),
                    ("LT_METAL_PART", "METAL PART"),
                    ("LT_CONNECTOR", "LT CONNECTOR")
                ],
                "OIL SIDE": [
                    ("GAUGE_GLASS", "GAUGE GLASS"),
                    ("OIL_AS_PER_NP", "OIL AS PER NP"),
                    ("OIL_POSITION", "OIL POSITION"),
                    ("OUTSIDE_PAINT", "OUTSIDE PAINT")
                ]
            },
            "right": {
                "OTHER DETAILS": [
                    ("BOLT_NUTS", "BOLT NUTS"),
                    ("ROD_GASKET", "ROD GASKET"),
                    ("TOP_GASKET", "TOP GASKET"),
                    ("NAME_PLATE", "NAME PLATE"),
                    ("BREATHER", "BREATHER"),
                    ("LABOUR_CHARGE", "LABOUR CHARGE"),
                    ("BS", "B/S"),
                    ("CONSERVATOR_TANK", "CONSERVATOR TANK (in Kg)"),
                    ("RADIATORS", "RADIATORS"),
                    ("REMARKS", "REMARKS")
                ]
            }
        }
        
        # Function to create section with unique field IDs
        def create_section(parent, title, fields):
            section_frame = ttk.LabelFrame(
                parent,
                text=title,
                padding=15,
                style='Section.TLabelframe'
            )
            section_frame.pack(fill="x", pady=(0, 20))
            
            for i, (field_id, field_label) in enumerate(fields):
                ttk.Label(
                    section_frame,
                    text=field_label,
                    font=("Arial", 10)
                ).grid(row=i, column=0, padx=10, pady=5, sticky="w")
                
                # Example for converting specific fields to radio buttons
                if field_label == "OUTSIDE PAINT":
                    oil_position_var = tk.StringVar(value="NR")  # Default value
                    full_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=oil_position_var, value="Reqd")
                    full_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    empty_radio = ttk.Radiobutton(section_frame, text="NR", variable=oil_position_var, value="NR")
                    empty_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = oil_position_var  # Store the variable for later access

                elif field_label == "GAUGE GLASS":
                    oil_position_var = tk.StringVar(value="NR")  # Default value
                    full_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=oil_position_var, value="Reqd")
                    full_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    empty_radio = ttk.Radiobutton(section_frame, text="NR", variable=oil_position_var, value="NR")
                    empty_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = oil_position_var  # Store the variable for later access

                elif field_label == "OIL POSITION":
                    oil_condition_var = tk.StringVar(value="Empty")  # Default value
                    good_radio = ttk.Radiobutton(section_frame, text="Full", variable=oil_condition_var, value="Full")
                    good_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    poor_radio = ttk.Radiobutton(section_frame, text="Empty", variable=oil_condition_var, value="Empty")
                    poor_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = oil_condition_var  # Store the variable for later access
                
                else:
                    entry = ttk.Entry(section_frame, width=30)
                    entry.grid(row=i, column=1, padx=5, pady=5)
                    entries[field_id] = entry
        
        # Create left sections
        for title, fields in sections["left"].items():
            create_section(left_frame, title, fields)
        
        # Create right sections
        for title, fields in sections["right"].items():
            create_section(right_frame, title, fields)
        
        # Button frame
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=20)
        
        # Submit and Cancel buttons with better styling
        submit_btn = ttk.Button(
            button_frame,
            text="Submit",
            command=lambda: self.submit_inspection(entries, physical_date, row_data),
            style='Submit.TButton',
            width=15
        )
        submit_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            command=inspection_window.destroy,
            width=15
        )
        cancel_btn.pack(side=tk.LEFT, padx=10)
        
        # Configure button styles
        style.configure(
            'Submit.TButton',
            font=('Arial', 10, 'bold'),
            background='#4CAF50'
        )

    def update_sheets(self, row_data, inspection_data):
        try:
            # Update Master Sheet
            master_row = None
            master_data = self.master_sheet.get_all_values()
            
            for idx, row in enumerate(master_data):
                if row[5] == row_data[5]:  # Match TC NO
                    master_row = idx + 1
                    break
            
            if master_row:
                # Update columns J to AD (10 to 30)
                range_name = f'J{master_row}:AD{master_row}'
                self.master_sheet.update(range_name, [inspection_data])
                
                # Update Physical Sheet
                # physical_sheet = setup_google_sheets().worksheet("PHYSICAL")
                
                # # Prepare complete row data
                # physical_row_data = row_data[:9] + inspection_data  # First 9 columns + inspection data
                
                # # Check if TC already exists in Physical sheet
                # physical_data = physical_sheet.get_all_values()
                # physical_row = None
                
                # for idx, row in enumerate(physical_data):
                #     if row[5] == row_data[5]:  # Match TC NO
                #         physical_row = idx + 1
                #         break
                
                # if physical_row:
                #     # Update existing row
                #     physical_sheet.update(f'A{physical_row}:AD{physical_row}', [physical_row_data])
                # else:
                #     # Add new row
                #     physical_sheet.append_row(physical_row_data)
                
                messagebox.showinfo("Success", "Inspection data saved successfully!")
            else:
                raise Exception("TC not found in master sheet")
            
        except Exception as e:
            raise Exception(f"Error updating sheets: {str(e)}")

    def create_top_search_section(self):
        # Create search frame
        search_frame = ttk.LabelFrame(self.main_container, text="Search Transformer", padding=10)
        search_frame.pack(fill=tk.X, pady=(0, 20), padx=20)
        
        # Create search fields container
        fields_frame = ttk.Frame(search_frame)
        fields_frame.pack(pady=5)
        
        # MR No. field
        ttk.Label(fields_frame, text="MR No.:").pack(side=tk.LEFT, padx=5)
        self.mr_search = ttk.Entry(fields_frame, width=30)
        self.mr_search.pack(side=tk.LEFT, padx=5)
        
        # Search button
        ttk.Button(
            fields_frame, 
            text="Search",
            command=self.search_and_update_table
        ).pack(side=tk.LEFT, padx=20)
        
        # Add Close Form button
        ttk.Button(
            fields_frame,
            text="Close Form",
            command=self.window.destroy,
            width=15
        ).pack(side=tk.LEFT, padx=20)

    def search_and_update_table(self):
        try:
            mr_no = self.mr_search.get().strip()
            
            if not mr_no:
                # If field is empty, show all data
                self.load_table_data()
                return
            
            # Clear existing items and buttons
            for button in self.buttons:
                button.destroy()
            self.buttons = []
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Get data from master sheet
            data = self.master_sheet.get_all_values()
            
            # Find required column indices
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
                "TC NO": 5, "MAKE": 6, "TC CAPACITY": 7, "JOB NO": 8
            }
            
            # Process each row, starting from row 3 (skip first two rows)
            for row in data[2:]:
                # Check if row matches search criteria
                row_mr_no = row[required_cols["MR NO"]]
                
                if mr_no.lower() in row_mr_no.lower():
                    row_data = [row[required_cols[col]] for col in required_cols]
                    
                    # Check if physical inspection is completed
                    has_physical_data = any(row[9:28])  # Columns J to AB
                    physical_status = "YES" if has_physical_data else "NO"
                    row_data.append(physical_status)
                    row_data.append("")  # Empty string for EXECUTE column
                    
                    # Add row to tree
                    item = self.tree.insert('', tk.END, values=row_data)
                    
                    # Create button immediately
                    self.create_button_for_row(item, row[required_cols["TC NO"]], physical_status)
            
            if not self.tree.get_children():
                messagebox.showinfo("No Results", "No matching records found.")
                self.load_table_data()  # Reload all data if no matches found
                
        except Exception as e:
            messagebox.showerror("Error", f"Error searching data: {str(e)}")
            print(f"Debug - Search error details: {str(e)}")

    def update_button_positions(self, event=None):
        """Update positions of all buttons when tree view changes"""
        try:
            # Hide all buttons first
            for button in self.buttons:
                button.place_forget()
            
            # Update position for each visible item
            for item in self.tree.get_children():
                bbox = self.tree.bbox(item, "EXECUTE")
                if bbox:  # Only place button if item is visible
                    values = self.tree.item(item)['values']
                    if values:
                        tc_no = values[4]  # TC NO is at index 4
                        status = values[-2]  # Physical Completed is second to last
                        self.create_button_for_row(item, tc_no, status)
        except Exception as e:
            print(f"Error updating button positions: {str(e)}")

    def setup_button_styles(self):
        """Setup custom styles for buttons"""
        style = ttk.Style()
        
        # Style for Enrol button (active/clickable)
        style.configure(
            'Enrol.TButton',
            background='#FF0000',  # Red color
            foreground='white',
            padding=5,
            font=('Arial', 9, 'bold')
        )
        
        # Style for Enrolled button (disabled)
        style.configure(
            'Enrolled.TButton',
            background='#00FF00',  # Green color
            foreground='white',
            padding=5,
            font=('Arial', 9, 'bold')
        )

    def submit_inspection(self, entries, physical_date, row_data):
        try:
            # Validate all entries
            for field_id, entry in entries.items():
                if not entry.get().strip():
                    messagebox.showerror("Error", f"Please fill in all fields")
                    return
            
            # Collect data in order for columns J to AD
            inspection_data = [
                physical_date.get(),  # Date of Physical
                # HT SIDE
                entries["HT_BUSHING"].get(),
                entries["HT_METAL_PART"].get(),
                entries["HT_CONNECTOR"].get(),
                # LT SIDE
                entries["LT_BUSHING"].get(),
                entries["LT_METAL_PART"].get(),
                entries["LT_CONNECTOR"].get(),
                # OIL SIDE
                entries["GAUGE_GLASS"].get(),
                entries["OIL_AS_PER_NP"].get(),
                entries["OIL_POSITION"].get(),
                entries["OUTSIDE_PAINT"].get(),
                # OTHER DETAILS
                entries["BOLT_NUTS"].get(),
                entries["ROD_GASKET"].get(),
                entries["TOP_GASKET"].get(),
                entries["NAME_PLATE"].get(),
                entries["BREATHER"].get(),
                entries["LABOUR_CHARGE"].get(),
                entries["BS"].get(),
                entries["CONSERVATOR_TANK"].get(),
                entries["RADIATORS"].get(),
                entries["REMARKS"].get()
            ]
            
            # Update sheets
            self.update_sheets(row_data, inspection_data)
            
            # Show success message and close window
            messagebox.showinfo("Success", "Inspection data saved successfully!")
            
            # Refresh main table
            self.load_table_data()
            
            # Close the inspection window
            for widget in self.window.winfo_children():
                if isinstance(widget, tk.Toplevel):
                    widget.destroy()
                    break
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error submitting inspection: {str(e)}")
            print(f"Debug - Submit error details: {str(e)}")

def create_physical_form(sheet):
    PhysicalVerificationForm(sheet) 