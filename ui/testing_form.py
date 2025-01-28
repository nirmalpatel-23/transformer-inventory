import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from config.sheets_setup import setup_google_sheets

class TestingVerificationForm:
    def __init__(self, sheet):
        self.window = tk.Toplevel()
        self.window.title("TESTING FORM")
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
            text="TESTING FORM", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Add search section
        self.create_top_search_section()
        
        # Create the table frame
        self.create_data_table()

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

    def create_data_table(self):
        table_frame = ttk.Frame(self.main_container)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)
        
        columns = (
            "DIVISION", "MR NO", "LOT NO", "DATE", "TC NO", 
            "MAKE", "TC CAPACITY", "JOB NO", "Testing Completed", "EXECUTE"
        )
        
        self.tree = ttk.Treeview(
            table_frame, 
            columns=columns, 
            show='headings', 
            height=10,
            style="Treeview"
        )
        
        # Configure columns
        column_widths = {col: 100 for col in columns}
        column_widths["Testing Completed"] = 120
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], anchor='center')
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind events
        self.tree.bind('<Map>', lambda e: self.load_table_data())
        self.tree.bind('<Destroy>', lambda e: self.cleanup_buttons())
        self.tree.bind('<<TreeviewSelect>>', self.update_button_positions)
        self.tree.bind('<Configure>', self.update_button_positions)

    def load_table_data(self):
        try:
            # Check if the treeview exists
            if not self.tree.winfo_exists():
                print("Treeview does not exist.")
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
            
            # Process each row, starting from row 3
            for row in data[2:]:
                row_data = [row[required_cols[col]] for col in required_cols]
                
                # Check if testing is completed (columns AT to BH)
                has_testing_data = any(row[47:62])  # Columns AT to BH
                testing_status = "YES" if has_testing_data else "NO"
                row_data.append(testing_status)
                row_data.append("")  # Empty string for EXECUTE column
                
                # Add row to tree
                item = self.tree.insert('', tk.END, values=row_data)
                
                # Create button
                self.create_button_for_row(item, row[required_cols["TC NO"]], testing_status)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error loading table data: {str(e)}")
            print(f"Debug - Error details: {str(e)}")

    def create_button_for_row(self, item, tc_no, status):
        try:
            bbox = self.tree.bbox(item, "EXECUTE")
            if not bbox:
                return
            
            if status == "NO":
                button = tk.Button(
                    self.tree,
                    text="Enrol",
                    bg='#FF0000',
                    fg='white',
                    command=lambda tc=tc_no: self.start_enrollment(tc),
                    width=8,
                    font=('Arial', 9, 'bold'),
                    relief="raised",
                    borderwidth=2
                )
            else:
                button = tk.Button(
                    self.tree,
                    text="Enrolled",
                    bg='#00FF00',
                    fg='white',
                    state='disabled',
                    width=8,
                    font=('Arial', 9, 'bold'),
                    relief="raised",
                    borderwidth=2
                )
            
            x = bbox[0] + 10
            y = bbox[1] + 2
            button.place(x=x, y=y)
            self.buttons.append(button)
            
        except Exception as e:
            print(f"Error creating button: {str(e)}")

    def cleanup_buttons(self):
        """Clean up buttons when table is destroyed"""
        for button in self.buttons:
            button.destroy()
        self.buttons = []

    def update_button_positions(self, event=None):
        """Update positions of all buttons when tree view changes"""
        try:
            for button in self.buttons:
                button.place_forget()
            
            for item in self.tree.get_children():
                bbox = self.tree.bbox(item, "EXECUTE")
                if bbox:
                    values = self.tree.item(item)['values']
                    if values:
                        tc_no = values[4]  # TC NO is at index 4
                        status = values[-2]  # Testing Completed is second to last
                        self.create_button_for_row(item, tc_no, status)
        except Exception as e:
            print(f"Error updating button positions: {str(e)}")

    def search_and_update_table(self):
        try:
            mr_no = self.mr_search.get().strip()
            
            if not mr_no:
                self.load_table_data()
                return
            
            # Clear existing items
            for button in self.buttons:
                button.destroy()
            self.buttons = []
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Get data and search
            data = self.master_sheet.get_all_values()
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
                "TC NO": 5, "MAKE": 6, "TC CAPACITY": 7, "JOB NO": 8
            }
            
            for row in data[2:]:
                row_mr_no = row[required_cols["MR NO"]]
                
                if mr_no.lower() in row_mr_no.lower():
                    row_data = [row[required_cols[col]] for col in required_cols]
                    
                    has_testing_data = any(row[45:60])  # Columns AT to BH
                    testing_status = "YES" if has_testing_data else "NO"
                    row_data.append(testing_status)
                    row_data.append("")
                    
                    item = self.tree.insert('', tk.END, values=row_data)
                    self.create_button_for_row(item, row[required_cols["TC NO"]], testing_status)
            
            if not self.tree.get_children():
                messagebox.showinfo("No Results", "No matching records found.")
                self.load_table_data()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error searching data: {str(e)}")
            print(f"Debug - Search error details: {str(e)}")

    def start_enrollment(self, tc_no):
        """Start the enrollment process for a TC"""
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

    def create_inspection_window(self, row_data):
        # Create new window
        inspection_window = tk.Toplevel(self.window)
        inspection_window.title(f"Testing Inspection Form - TC No: {row_data[5]}")
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
        
        # TC Details Section
        details_frame = ttk.LabelFrame(main_frame, text="Transformer Details", padding=15)
        details_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        # Display TC details in grid
        tc_details = {
            "DIVISION": row_data[0],
            "MR NO": row_data[2],
            "DATE": row_data[4],
            "TC NO": row_data[5],
            "MAKE": row_data[6],
            "CAPACITY": row_data[7],
            "JOB NO": row_data[8]
        }
        
        for i, (label, value) in enumerate(tc_details.items()):
            ttk.Label(details_frame, text=f"{label}:").grid(row=i//3, column=(i%3)*2, padx=5, pady=5)
            entry = ttk.Entry(details_frame, width=20)
            entry.insert(0, value)
            entry.configure(state='readonly')
            entry.grid(row=i//3, column=(i%3)*2+1, padx=5, pady=5)
        
        # Testing Date Section
        date_frame = ttk.LabelFrame(main_frame, text="Testing Date", padding=15)
        date_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        testing_date = DateEntry(
            date_frame,
            width=20,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        testing_date.pack(pady=5)
        
        # Create frames for test sections
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True, padx=10)
        
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 10))
        
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(10, 0))
        
        # Store entries for later access
        entries = {}
        
        # Define sections and their fields
        sections = {
            "left": {
                "Insulation Test (M Ohm)": [
                    ("HV_TO_E", "HV to E"),
                    ("LV_TO_E", "LV to E"),
                    ("HV_TO_LV", "HV to LV")
                ],
                "No Load Test (O C TEST)": [
                    ("NL_VOLTS", "No Load Volts"),
                    ("NL_AMP", "No Load Amp"),
                    ("NL_WATTS", "No Load Watts")
                ],
                "Full Load Test (S C Test)": [
                    ("FL_VOLTS", "Full Load Volts"),
                    ("FL_AMP", "Full Load Amp"),
                    ("FL_WATTS", "Full Load Watts")
                ]
            },
            "right": {
                "Other Tests": [
                    ("INDUCED_OV", "Induced O.V.Test (VOLTAGE AT 100 Hz/MIN)"),
                    ("HV_TEST", "HV TEST (22 KV/MIN)"),
                    ("OIL_TEST", "OIL DITEST TEST (40 KV/MIN)"),
                    ("NO_LOAD_RATIO", "NO LOAD (RATIO)"),
                    ("REMARKS", "REMARKS")
                ]
            }
        }
        
        # Function to create section
        def create_section(parent, title, fields):
            section_frame = ttk.LabelFrame(parent, text=title, padding=15)
            section_frame.pack(fill="x", pady=(0, 20))
            
            for i, (field_id, field_label) in enumerate(fields):
                ttk.Label(section_frame, text=field_label).grid(row=i, column=0, padx=5, pady=5, sticky="w")
                entry = ttk.Entry(section_frame, width=30)
                entry.grid(row=i, column=1, padx=5, pady=5)
                entries[field_id] = entry
        
        # Create sections
        for title, fields in sections["left"].items():
            create_section(left_frame, title, fields)
        
        for title, fields in sections["right"].items():
            create_section(right_frame, title, fields)
        
        # Add Submit button to right frame
        submit_btn = tk.Button(
            right_frame,
            text="SUBMIT",
            command=lambda: self.submit_inspection(entries, testing_date, row_data),
            width=20,
            height=2,
            bg='red',
            fg='white',
            font=('Arial', 12, 'bold')
        )
        submit_btn.pack(pady=(20, 0))  # Add padding at top

        # Make the window resizable
        inspection_window.resizable(True, True)

    def submit_inspection(self, entries, testing_date, row_data):
        try:
            # Validate entries
            for field_id, entry in entries.items():
                if not entry.get().strip():
                    messagebox.showerror("Error", "Please fill in all fields")
                    return
            
            # Collect data for sheets in proper order
            inspection_data = [
                testing_date.get(),
                # Insulation Test
                entries["HV_TO_E"].get(),
                entries["LV_TO_E"].get(),
                entries["HV_TO_LV"].get(),
                # No Load Test (O C TEST)
                entries["NL_VOLTS"].get(),
                entries["NL_AMP"].get(),
                entries["NL_WATTS"].get(),
                # Full Load Test (S C Test)
                entries["FL_VOLTS"].get(),
                entries["FL_AMP"].get(),
                entries["FL_WATTS"].get(),
                # Other Tests
                entries["INDUCED_OV"].get(),
                entries["HV_TEST"].get(),
                entries["OIL_TEST"].get(),
                entries["NO_LOAD_RATIO"].get(),
                entries["REMARKS"].get()
            ]
            
            # Update Master Sheet (columns AV to BJ)
            master_row = None
            master_data = self.master_sheet.get_all_values()
            
            for idx, row in enumerate(master_data):
                if row[5] == row_data[5]:  # Match TC NO
                    master_row = idx + 1
                    break
            
            if master_row:
                try:
                    # Update Master Sheet
                    range_name = f'AV{master_row}:BJ{master_row}'
                    self.master_sheet.update(range_name, [inspection_data])
                    
                    # Update Testing Sheet
                    # testing_sheet = setup_google_sheets().worksheet("TESTING")
                    # testing_data = inspection_data  # Combine TC details with testing data
                    
                    # # Check if TC already exists in Testing sheet
                    # testing_data_all = testing_sheet.get_all_values()
                    # testing_row = None
                    
                    # for idx, row in enumerate(testing_data_all):
                    #     if len(row) > 5 and row[5] == row_data[5]:  # Match TC NO
                    #         testing_row = idx + 1
                    #         break
                    
                    # if testing_row:
                    #     # Update existing row
                    #     range_name = f'J{testing_row}:X{testing_row}'
                    #     testing_sheet.update(range_name, [testing_data])
                    # else:
                    #     # Add new row
                    #     testing_sheet.append_row(testing_data)
                    
                    messagebox.showinfo("Success", "Testing data saved successfully!")
                    self.load_table_data()  # Refresh the main table
                    
                    # Close inspection window
                    for widget in self.window.winfo_children():
                        if isinstance(widget, tk.Toplevel):
                            widget.destroy()
                            break
                            
                except Exception as e:
                    messagebox.showerror("Error", f"Error saving data: {str(e)}")
                    print(f"Debug - Save error details: {str(e)}")
            else:
                raise Exception("TC not found in master sheet")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error submitting inspection: {str(e)}")
            print(f"Debug - Submit error details: {str(e)}")

def create_testing_form(sheet):
    TestingVerificationForm(sheet)