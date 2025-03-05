import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from config.sheets_setup import setup_google_sheets, get_worksheet
from ui.enrollment_form import EnrollmentForm  # Ensure this import is correct
from ui.physical_form import PhysicalVerificationForm  # Ensure this import is correct

class TestingVerificationForm:
    def __init__(self, sheet):
        self.window = tk.Toplevel()
        self.window.title("TESTING FORM")
        self.window.geometry("1200x800")
        
        # Initialize buttons list
        self.buttons = []
        
        # Store the master sheet reference
        self.master_sheet = get_worksheet(sheet)  # Ensure this returns the correct worksheet object
        
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
        self.create_data_table(self.main_container)

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

    def create_data_table(self, container):
        table_frame = ttk.Frame(container)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)
        
        columns = (
            "DIVISION", "MR NO", "LOT NO", "DATE", "EXECUTE"
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
        self.tree.bind('<<TreeviewSelect>>', self.show_details)
        self.tree.bind('<Configure>', self.update_button_positions)

    def load_table_data(self):
        try:
            # Check if the master_sheet is valid
            if not self.master_sheet:
                print("Master sheet is not valid.")
                return

            # Get data from master sheet
            data = self.master_sheet.get_all_values()  # This should be called on the worksheet object
            
            # Check if the treeview exists
            if not self.tree.winfo_exists():
                print("Treeview does not exist.")
                return  # Exit the function if the treeview is not valid

            # Clear existing items and buttons
            for button in self.buttons:
                button.destroy()
            self.buttons = []
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Find required column indices
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
            }
            
            grouped_data = {}
            for row in data[2:]:
                mr_no = row[required_cols["MR NO"]]
                if mr_no not in grouped_data:
                    grouped_data[mr_no] = [row[required_cols[col]] for col in required_cols]
                    # Check if testing is completed (columns AT to BH)
                    has_testing_data = any(row[45:60])  # Columns AT to BH
                    testing_status = "YES" if has_testing_data else "NO"
                    grouped_data[mr_no].append(testing_status)
                    grouped_data[mr_no].append("")  # Empty string for EXECUTE column
            
            # Insert grouped data into the tree
            for mr_no, row_data in grouped_data.items():
                item = self.tree.insert('', tk.END, values=row_data)
                # Create button
                self.create_button_for_row(item, mr_no, row_data[-2])  # MR NO and Testing status
                
        except Exception as e:
            messagebox.showerror("Error", f"Error loading table data: {str(e)}")
            print(f"Debug - Error details: {str(e)}")

    def create_button_for_row(self, item, mr_no, status):
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
                    command=lambda mr=mr_no: self.start_enrollment(mr),
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
                        mr_no = values[1]  # MR NO is at index 1
                        status = values[-2]  # Testing Completed is second to last
                        self.create_button_for_row(item, mr_no, status)
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
            print(type(self.master_sheet))  # Add this line before calling get_all_values()
            data = self.master_sheet.get_all_values()
            required_cols = {
                "DIVISION": 0, "MR NO": 2, "LOT NO": 3, "DATE": 4,
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
                    self.create_button_for_row(item, row[required_cols["MR NO"]], testing_status)
            
            if not self.tree.get_children():
                messagebox.showinfo("No Results", "No matching records found.")
                self.load_table_data()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error searching data: {str(e)}")
            print(f"Debug - Search error details: {str(e)}")

    def start_enrollment(self, mr_no):
        """Start the enrollment process for a MR NO"""
        try:
            # Get row data from master sheet
            print(type(self.master_sheet))  # Add this line before calling get_all_values()
            data = self.master_sheet.get_all_values()
            row_data = None
            
            # Skip header rows and find the matching MR NO
            for row in data[2:]:  # Skip first two rows
                if str(row[2]).strip() == str(mr_no).strip():  # MR NO is in column C (index 2)
                    row_data = row
                    break
            
            if row_data is None:
                messagebox.showerror("Error", f"MR No. {mr_no} not found in master sheet")
                return
            
            # Create a new physical verification form with the designated MR NO
            PhysicalVerificationForm(row_data)  # Pass the row data to the physical form
            
        except Exception as e:
            messagebox.showerror("Error", f"Error starting enrollment: {str(e)}")
            print(f"Debug - Start enrollment error: {str(e)}")

    def create_bill_window(self, row_data):
        """Create a new window to display a physical bill for the selected MR NO"""
        bill_window = tk.Toplevel(self.window)
        bill_window.title(f"Physical Bill - MR No: {row_data[2]}")
        bill_window.geometry("1200x800")  # Adjust size as needed
        
        # Create main container for the bill
        bill_frame = ttk.Frame(bill_window, padding=20)
        bill_frame.pack(fill=tk.BOTH, expand=True)

        # Add a title for the bill
        title_label = ttk.Label(bill_frame, text="Physical Bill", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # Create a section for the details
        detail_labels = [
            "Division", "MR No", "Lot No", "Date", "TC No", "Make", "TC Capacity", "Job No", "Physical Completed"
        ]
        
        # Create a frame for the details
        details_frame = ttk.LabelFrame(bill_frame, text="Details", padding=10)
        details_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        for i, label in enumerate(detail_labels):
            ttk.Label(details_frame, text=f"{label}:").grid(row=i, column=0, sticky="w", padx=5, pady=5)
            ttk.Label(details_frame, text=row_data[i]).grid(row=i, column=1, sticky="w", padx=5, pady=5)

        # Add additional details if needed
        additional_details = [
            "HT Bushing", "HT Metal Part", "HT Connector", 
            "LT Bushing", "LT Metal Part", "LT Connector",
            "Gauge Glass", "Oil Position", "Outside Paint"
        ]
        
        # Assuming you have a way to get these details from row_data or another source
        for i, detail in enumerate(additional_details):
            ttk.Label(details_frame, text=f"{detail}:").grid(row=i + len(detail_labels), column=0, sticky="w", padx=5, pady=5)
            ttk.Label(details_frame, text=row_data[i + len(detail_labels)]).grid(row=i + len(detail_labels), column=1, sticky="w", padx=5, pady=5)

        # Add a close button
        close_button = ttk.Button(bill_frame, text="Close", command=bill_window.destroy)
        close_button.pack(pady=10)

    def show_details(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item[0])['values']
            mr_no = item_values[1]  # Assuming MR NO is at index 1
            self.display_mr_details(mr_no)

    def display_mr_details(self, mr_no):
        """Display details of the selected MR NO in a new window."""
        details_window = tk.Toplevel(self.window)
        details_window.title(f"Details for MR No: {mr_no}")
        details_window.geometry("600x400")

        # Add a title label
        title_label = ttk.Label(details_window, text="Hello", font=("Arial", 14, "bold"))
        title_label.pack(pady=(10, 20))  # Add some padding

        # Add a description label
        description_label = ttk.Label(details_window, text="Here are the details for the selected MR No:", font=("Arial", 12))
        description_label.pack(pady=(0, 10))  # Add some padding

        # Fetch data for the selected MR NO
        print(type(self.master_sheet))  # Add this line before calling get_all_values()
        data = self.master_sheet.get_all_values()
        details = [row for row in data[2:] if row[2] == mr_no]  # Assuming MR NO is at index 2

        # Create a text widget to display details
        details_text = tk.Text(details_window, wrap=tk.WORD)
        details_text.pack(expand=True, fill=tk.BOTH)

        # Insert details into the text widget
        for row in details:
            details_text.insert(tk.END, f"{row}\n")

        # Create a new table frame for additional details
        table_frame = ttk.Frame(details_window)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)
        
        columns = (
            "DIVISION", "MR NO", "LOT NO", "DATE", "EXECUTE"
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
        self.tree.bind('<<TreeviewSelect>>', self.show_details)
        self.tree.bind('<Configure>', self.update_button_positions)

        # Add a back button
        back_button = ttk.Button(details_window, text="Back", command=details_window.destroy)
        back_button.pack(pady=10)

        details_window.transient(self.window)  # Keep the new window on top of the main window
        details_window.grab_set()  # Make the new window modal

def create_final_bill(sheet):
    # Ensure 'sheet' is a valid object with the 'get_all_values' method
    TestingVerificationForm(sheet)