import sys
import os

# Add the parent directory to the system path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import save_to_sheets, get_worksheet, setup_google_sheets
from tkcalendar import DateEntry

class InternalVerificationForm:
    def __init__(self, sheet):
        self.window = tk.Toplevel()
        self.window.title("INTERNAL FORM")
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
            text="INTERNAL FORM", 
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
        # Similar to physical form but with internal inspection columns
        table_frame = ttk.Frame(self.main_container)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)
        
        columns = (
            "DIVISION", "MR NO", "LOT NO", "DATE", "TC NO", 
            "MAKE", "TC CAPACITY", "JOB NO", "Internal Completed", "EXECUTE"
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
        column_widths["Internal Completed"] = 120
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], anchor='center', stretch=False)
        
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
        self.tree.bind('<MouseWheel>', self.update_button_positions)

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
            
            # Find required column indices
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
                "TC NO": 5, "MAKE": 6, "TC CAPACITY": 7, "JOB NO": 8
            }
            
            # Process each row, starting from row 3
            for row in data[2:]:
                row_data = [row[required_cols[col]] for col in required_cols]
                
                # Check if internal inspection is completed (columns AC to AS)
                has_internal_data = any(row[30:47])  # Columns AC to AS
                internal_status = "YES" if has_internal_data else "NO"
                row_data.append(internal_status)
                row_data.append("")  # Empty string for EXECUTE column
                
                # Add row to tree
                item = self.tree.insert('', tk.END, values=row_data)
                
                # Create button
                self.create_button_for_row(item, row[required_cols["TC NO"]], internal_status)
                
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

    def create_inspection_window(self, row_data):
        # Create new window
        inspection_window = tk.Toplevel(self.window)
        inspection_window.title(f"Internal Inspection Form - TC No: {row_data[5]}")
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
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Create styles
        style = ttk.Style()
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Section.TLabelframe', padding=15, relief='solid')
        style.configure('Main.TFrame', background='white')
        
        # TC Details Section
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
        
        # Left Column (HT COIL, LT COIL)
        left_frame = ttk.Frame(inspection_frame)
        left_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 10))
        
        # Right Column (OTHER DETAILS)
        right_frame = ttk.Frame(inspection_frame)
        right_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(10, 0))
        
        # Store entries for later access
        entries = {}
        
        # Create sections
        sections = {
            "left": {
                "HT COIL": [
                    ("HT_COIL_A", "A"),
                    ("HT_COIL_B", "B"),
                    ("HT_COIL_C", "C")
                ],
                "LT COIL": [
                    ("LT_COIL_A", "A"),
                    ("LT_COIL_B", "B"),
                    ("LT_COIL_C", "C")
                ]
            },
            "right": {
                "OTHER DETAILS": [
                    ("COIL_PER_PHASE", "COIL PER PHASE"),
                    ("CU_ALU", "CU ALU"),
                    ("WT_OF_HT_COILS", "WT. OF HT COILS"),
                    ("WT_OF_LT_COILS", "WT. OF LT COILS"),
                    ("WASHER_RINGS", "WASHER RINGS"),
                    ("INSIDE_PAINT", "INSIDE PAINT"),
                    ("INSULATING_MATERIAL", "INSULATING MATERIAL"),
                    ("TESTING_CHARGES", "TESTING CHARGES"),
                    ("COIL_SE_DPC", "COIL SE/DPC"),
                    ("REMARKS", "REMARKS")
                ]
            }
        }
        
        # Function to create section
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
                
                entry = ttk.Entry(section_frame, width=30)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                entries[field_id] = entry
        
        # Create sections
        for title, fields in sections["left"].items():
            create_section(left_frame, title, fields)
        
        for title, fields in sections["right"].items():
            create_section(right_frame, title, fields)
        
        # Button frame
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=20)
        
        # Submit and Cancel buttons
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
                    
                    has_internal_data = any(row[28:45])  # Columns AC to AS
                    internal_status = "YES" if has_internal_data else "NO"
                    row_data.append(internal_status)
                    row_data.append("")
                    
                    item = self.tree.insert('', tk.END, values=row_data)
                    self.create_button_for_row(item, row[required_cols["TC NO"]], internal_status)
            
            if not self.tree.get_children():
                messagebox.showinfo("No Results", "No matching records found.")
                self.load_table_data()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error searching data: {str(e)}")
            print(f"Debug - Search error details: {str(e)}")

    def cleanup_buttons(self):
        """Clean up buttons when table is destroyed"""
        for button in self.buttons:
            button.destroy()
        self.buttons = []

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
                        status = values[-2]  # Internal Completed is second to last
                        self.create_button_for_row(item, tc_no, status)
        except Exception as e:
            print(f"Error updating button positions: {str(e)}")

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
            print(f"Debug - TC No: {tc_no}")

    def submit_inspection(self, entries, physical_date, row_data):
        try:
            # Validate all entries
            for field_id, entry in entries.items():
                if not entry.get().strip():
                    messagebox.showerror("Error", f"Please fill in all fields")
                    return
            
            # Collect data in order for columns AE to AU
            inspection_data = [
                physical_date.get(),  # Date of Internal
                # HT COIL
                entries["HT_COIL_A"].get(),
                entries["HT_COIL_B"].get(),
                entries["HT_COIL_C"].get(),
                # LT COIL
                entries["LT_COIL_A"].get(),
                entries["LT_COIL_B"].get(),
                entries["LT_COIL_C"].get(),
                # OTHER DETAILS
                entries["COIL_PER_PHASE"].get(),
                entries["CU_ALU"].get(),
                entries["WT_OF_HT_COILS"].get(),
                entries["WT_OF_LT_COILS"].get(),
                entries["WASHER_RINGS"].get(),
                entries["INSIDE_PAINT"].get(),
                entries["INSULATING_MATERIAL"].get(),
                entries["TESTING_CHARGES"].get(),
                entries["COIL_SE_DPC"].get(),
                entries["REMARKS"].get()
            ]
            
            # Update sheets
            self.update_sheets(row_data, inspection_data)
            
            # Show success message and close window
            messagebox.showinfo("Success", "Internal inspection data saved successfully!")
            
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

    def update_sheets(self, row_data, inspection_data):
        """Update both Master and Internal sheets"""
        try:
            # Update Master Sheet
            master_row = None
            master_data = self.master_sheet.get_all_values()
            
            for idx, row in enumerate(master_data):
                if row[5] == row_data[5]:  # Match TC NO
                    master_row = idx + 1
                    break
            
            if master_row:
                # Update columns AE to AU (30 to 46)
                range_name = f'AE{master_row}:AU{master_row}'
                self.master_sheet.update(range_name, [inspection_data])
                
            #     # Update Internal Sheet
            #     internal_sheet = setup_google_sheets().worksheet("INTERNAL")
                
            #     # Prepare complete row data
            #     internal_row_data = row_data[:9] + inspection_data
                
            #     # Check if TC already exists in Internal sheet
            #     internal_data = internal_sheet.get_all_values()
            #     internal_row = None
                
            #     for idx, row in enumerate(internal_data):
            #         if len(row) > 5 and row[5] == row_data[5]:  # Match TC NO
            #             internal_row = idx + 1
            #             break
                
            #     if internal_row:
            #         # Update existing row
            #         range_name = f'A{internal_row}:AS{internal_row}'
            #         internal_sheet.update(range_name, [internal_row_data])
            #     else:
            #         # Add new row
            #         internal_sheet.append_row(internal_row_data)
            # else:
            #     raise Exception("TC not found in master sheet")
                
        except Exception as e:
            raise Exception(f"Error updating sheets: {str(e)}")

def create_internal_form(sheet):
    InternalVerificationForm(sheet) 