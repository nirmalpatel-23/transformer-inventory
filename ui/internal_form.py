import sys
import os

# Add the parent directory to the system path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import save_to_sheets, get_worksheet, setup_google_sheets
from tkcalendar import DateEntry

class InternalVerificationForm:
    def __init__(self, master_sheet, mr_no=None):
        self.mr_no = mr_no
        self.window = tk.Toplevel()
        self.window.title("INTERNAL FORM")
        self.window.geometry("1200x800")
        
        # Initialize buttons list
        self.buttons = []
        
        # Store the master sheet reference
        self.master_sheet = master_sheet
        
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
        self.create_data_table(mr_no)

    def create_top_search_section(self):
        """Create search section at top of form"""
        search_frame = ttk.LabelFrame(self.main_container, text="Search Transformer", padding=10)
        search_frame.pack(fill=tk.X, pady=(0, 20), padx=20)
        
        fields_frame = ttk.Frame(search_frame)
        fields_frame.pack(pady=5)
        
        ttk.Label(fields_frame, text="TC No.:").pack(side=tk.LEFT, padx=5)
        self.tc_search = ttk.Entry(fields_frame, width=30)
        self.tc_search.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            fields_frame, 
            text="Search",
            command=self.search_and_update_table
        ).pack(side=tk.LEFT, padx=20)
        
        ttk.Button(
            fields_frame,
            text="Close Form",
            command=self.window.destroy,
            width=15
        ).pack(side=tk.LEFT, padx=20)

    def create_data_table(self, mr_no=None):
        """Create a table to display MR NOs and associated data."""
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
            "Internal Completed": 120,
            "EXECUTE": 100
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], anchor='center', stretch=False)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind events
        self.tree.bind('<Map>', lambda e: self.load_table_data(mr_no))
        self.tree.bind('<Destroy>', lambda e: self.cleanup_buttons())
        self.tree.bind('<<TreeviewSelect>>', self.update_button_positions)
        self.tree.bind('<Configure>', self.update_button_positions)
        self.tree.bind('<MouseWheel>', self.update_button_positions)

    def load_table_data(self, mr_no):
        """Load table data for given MR NO"""
        try:
            for button in self.buttons:
                button.destroy()
            self.buttons = []
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            data = self.master_sheet.get_all_values()
            
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
                "TC NO": 5, "MAKE": 6, "TC CAPACITY": 7, "JOB NO": 8
            }
            
            for row in data[2:]:
                if str(mr_no) == row[2]:
                    row_data = [row[required_cols[col]] for col in required_cols]
                    
                    # Check internal inspection completion (columns AE to AU / 31:47)
                    has_internal_data = any(row[31:47])
                    internal_status = "YES" if has_internal_data else "NO"
                    row_data.append(internal_status)
                    row_data.append("")
                    
                    item = self.tree.insert('', tk.END, values=row_data)
                    self.create_button_for_row(item, row[5], internal_status)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error loading table data: {str(e)}")

    def create_button_for_row(self, item, tc_no, status):
        """Create action button for table row"""
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
        """Create inspection form window"""
        inspection_window = tk.Toplevel(self.window)
        inspection_window.title(f"Internal Inspection Form - TC No: {row_data[5]}")
        inspection_window.geometry("1000x800")
        
        # Center window
        inspection_window.update_idletasks()
        width = inspection_window.winfo_width()
        height = inspection_window.winfo_height()
        x = (inspection_window.winfo_screenwidth() // 2) - (width // 2)
        y = (inspection_window.winfo_screenheight() // 2) - (height // 2)
        inspection_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Create scrollable frame
        main_frame = ttk.Frame(inspection_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        canvas = tk.Canvas(main_frame, bg='white')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Create styles
        style = ttk.Style()
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Section.TLabelframe', padding=15, relief='solid')
        
        # TC Details Section
        details_frame = ttk.LabelFrame(
            scrollable_frame, 
            text="Transformer Details", 
            padding=15,
            style='Section.TLabelframe'
        )
        details_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        # Display TC details
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
        
        # Date Section
        date_frame = ttk.LabelFrame(
            scrollable_frame, 
            text="Inspection Date", 
            padding=15,
            style='Section.TLabelframe'
        )
        date_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        internal_date = DateEntry(
            date_frame,
            width=15,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            font=("Arial", 10),
            date_pattern='dd/mm/yyyy'
        )
        internal_date.pack(pady=5)
        
        # Create inspection sections frame
        inspection_frame = ttk.Frame(scrollable_frame)
        inspection_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        # Left and Right frames
        left_frame = ttk.Frame(inspection_frame)
        left_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 10))
        
        right_frame = ttk.Frame(inspection_frame)
        right_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(10, 0))
        
        # Store entries
        entries = {}
        
        # Define sections with fields
        sections = {
            "left": {
                "HT COIL": [
                    ("HT_COIL_A", "HT - A"),
                    ("HT_COIL_B", "HT - B"),
                    ("HT_COIL_C", "HT - C")
                ],
                "LT COIL": [
                    ("LT_COIL_A", "LT - A"),
                    ("LT_COIL_B", "LT - B"),
                    ("LT_COIL_C", "LT - C")
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
                
                # Create radio buttons for specific fields
                if field_label == "CU ALU":
                    cu_alu_var = tk.StringVar(value="CU")  # Default value
                    cu_radio = ttk.Radiobutton(section_frame, text="CU", variable=cu_alu_var, value="CU")
                    cu_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    alu_radio = ttk.Radiobutton(section_frame, text="ALU", variable=cu_alu_var, value="ALU")
                    alu_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = cu_alu_var  # Store the variable for later access

                elif field_label == "INSIDE PAINT":
                    inside_paint_var = tk.StringVar(value="NR")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=inside_paint_var, value="Reqd")
                    yes_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    no_radio = ttk.Radiobutton(section_frame, text="NR", variable=inside_paint_var, value="NR")
                    no_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = inside_paint_var  # Store the variable for later access

                elif field_label == "INSULATING MATERIAL":
                    inside_paint_var = tk.StringVar(value="NR")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=inside_paint_var, value="Reqd")
                    yes_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    no_radio = ttk.Radiobutton(section_frame, text="NR", variable=inside_paint_var, value="NR")
                    no_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = inside_paint_var  # Store the variable for later access
                
                elif field_label == "TESTING CHARGES":
                    inside_paint_var = tk.StringVar(value="NR")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=inside_paint_var, value="Reqd")
                    yes_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    no_radio = ttk.Radiobutton(section_frame, text="NR", variable=inside_paint_var, value="NR")
                    no_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = inside_paint_var  # Store the variable for later access
                
                elif field_label == "COIL SE/DPC":
                    inside_paint_var = tk.StringVar(value="SE")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="SE", variable=inside_paint_var, value="SE")
                    yes_radio.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                    no_radio = ttk.Radiobutton(section_frame, text="DPC", variable=inside_paint_var, value="DPC")
                    no_radio.grid(row=i, column=1, padx=5, pady=5, sticky="e")
                    entries[field_id] = inside_paint_var  # Store the variable for later access
                
                elif field_label == "LT - A":
                    inside_paint_var = tk.StringVar(value="NR")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=inside_paint_var, value="Reqd")
                    yes_radio.grid(row=i, column=1, padx=10, pady=5, sticky="w")  # Column 1 for LT - A
                    no_radio = ttk.Radiobutton(section_frame, text="NR", variable=inside_paint_var, value="NR")
                    no_radio.grid(row=i, column=5, padx=10, pady=5, sticky="e")  # Column 1 for LT - A
                    entries[field_id] = inside_paint_var  # Store the variable for later access

                elif field_label == "LT - B":
                    inside_paint_var = tk.StringVar(value="NR")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=inside_paint_var, value="Reqd")
                    yes_radio.grid(row=i, column=1, padx=10, pady=5, sticky="w")  # Column 2 for LT - B
                    no_radio = ttk.Radiobutton(section_frame, text="NR", variable=inside_paint_var, value="NR")
                    no_radio.grid(row=i, column=5, padx=10, pady=5, sticky="e")  # Column 2 for LT - B
                    entries[field_id] = inside_paint_var  # Store the variable for later access

                elif field_label == "LT - C":
                    inside_paint_var = tk.StringVar(value="NR")  # Default value
                    yes_radio = ttk.Radiobutton(section_frame, text="Reqd", variable=inside_paint_var, value="Reqd")
                    yes_radio.grid(row=i, column=1, padx=10, pady=5, sticky="w")  # Column 3 for LT - C
                    no_radio = ttk.Radiobutton(section_frame, text="NR", variable=inside_paint_var, value="NR")
                    no_radio.grid(row=i, column=5, padx=10, pady=5, sticky="e")  # Column 3 for LT - C
                    entries[field_id] = inside_paint_var  # Store the variable for later access
                                
                else:
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
        
        submit_btn = ttk.Button(
            button_frame,
            text="Submit",
            command=lambda: self.submit_inspection(entries, internal_date, row_data),
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

    def submit_inspection(self, entries, internal_date, row_data):
        """Submit inspection data"""
        try:
            # Validate entries
            for field_id, entry in entries.items():
                if not entry.get().strip():
                    messagebox.showerror("Error", f"Please fill in all fields")
                    return
            
            # Collect data for columns AE to AU (31 to 47)
            inspection_data = [
                internal_date.get(),
                entries["HT_COIL_A"].get(),
                entries["HT_COIL_B"].get(),
                entries["HT_COIL_C"].get(),
                entries["LT_COIL_A"].get(),
                entries["LT_COIL_B"].get(),
                entries["LT_COIL_C"].get(),
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
            
            messagebox.showinfo("Success", "Internal inspection data saved successfully!")
            
            # Refresh table
            self.load_table_data(self.mr_no)
            
            # Close inspection window
            for widget in self.window.winfo_children():
                if isinstance(widget, tk.Toplevel):
                    widget.destroy()
                    break
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error submitting inspection: {str(e)}")

    def update_sheets(self, row_data, inspection_data):
        """Update master sheet with inspection data"""
        try:
            master_row = None
            master_data = self.master_sheet.get_all_values()
            
            for idx, row in enumerate(master_data):
                if row[5] == row_data[5]:  # Match TC NO
                    master_row = idx + 1
                    break
            
            if master_row:
                # Update columns AE to AU (31 to 47)
                range_name = f'AE{master_row}:AU{master_row}'
                self.master_sheet.update(range_name, [inspection_data])
            else:
                raise Exception("TC not found in master sheet")
                
        except Exception as e:
            raise Exception(f"Error updating sheets: {str(e)}")

    def cleanup_buttons(self):
        for button in self.buttons:
            button.destroy()
        self.buttons = []

    def update_button_positions(self, event=None):
        try:
            for button in self.buttons:
                button.place_forget()
            
            for item in self.tree.get_children():
                bbox = self.tree.bbox(item, "EXECUTE")
                if bbox:
                    values = self.tree.item(item)['values']
                    if values:
                        tc_no = values[4]  # TC NO is at index 4
                        status = values[-2]  # Internal Completed is second to last
                        self.create_button_for_row(item, tc_no, status)
        except Exception as e:
            print(f"Error updating button positions: {str(e)}")

    def search_and_update_table(self):
        try:
            tc_no = self.tc_search.get().strip()
            
            if not tc_no:
                self.load_table_data(self.mr_no)
                return
            
            for button in self.buttons:
                button.destroy()
            self.buttons = []
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            data = self.master_sheet.get_all_values()
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
                "TC NO": 5, "MAKE": 6, "TC CAPACITY": 7, "JOB NO": 8
            }
            
            for row in data[2:]:
                if str(tc_no).lower() in str(row[5]).lower():  # TC NO is at index 5
                    row_data = [row[required_cols[col]] for col in required_cols]
                    
                    has_internal_data = any(row[31:47])
                    internal_status = "YES" if has_internal_data else "NO"
                    row_data.append(internal_status)
                    row_data.append("")
                    
                    item = self.tree.insert('', tk.END, values=row_data)
                    self.create_button_for_row(item, row[5], internal_status)
            
            if not self.tree.get_children():
                messagebox.showinfo("No Results", "No matching records found.")
                self.load_table_data(self.mr_no)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error searching data: {str(e)}")

    def start_enrollment(self, tc_no):
        try:
            data = self.master_sheet.get_all_values()
            row_data = None
            
            for row in data[2:]:
                if str(row[5]).strip() == str(tc_no).strip():
                    row_data = row
                    break
            
            if row_data is None:
                messagebox.showerror("Error", f"TC No. {tc_no} not found in master sheet")
                return
            
            self.create_inspection_window(row_data)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error starting enrollment: {str(e)}")

def create_internal_form(sheet):
    InternalVerificationForm(sheet) 