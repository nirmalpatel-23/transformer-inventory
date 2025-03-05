import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import get_worksheet
from ui.physical_form import PhysicalVerificationForm

class PhysicalVerificationForm1:
    def __init__(self, master_sheet):
        self.master_sheet = master_sheet
        self.window = tk.Toplevel()
        self.window.title("Physical Verification Form")
        self.window.geometry("800x600")

        # Create main container
        self.main_container = ttk.Frame(self.window)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Add title
        title_label = ttk.Label(self.main_container, text="Physical Verification by MR NO", font=("Arial", 16, "bold"))
        title_label.pack(pady=(10, 20))

        # Create the table frame
        self.create_data_table()

    def create_data_table(self):
        """Create a table to display MR NOs and associated data."""
        # Create frame for table
        table_frame = ttk.Frame(self.main_container)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Create Treeview with style
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)

        # Create Treeview
        columns = (
            "MR NO", "DIVISION", "LOT NO", "DATE", "Physical Completed", "Total TCs", "VIEW"
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
            "MR NO": 100,
            "DIVISION": 100,
            "LOT NO": 100,
            "DATE": 100,
            "Physical Completed": 120,
            "Total TCs": 100,
            "VIEW": 80
        }

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100), anchor='center')

        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Pack the table and scrollbar
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Load data into the table
        self.load_data()

    def load_data(self):
        """Load MR NO data from the master sheet."""
        try:
            data = self.master_sheet.get_all_values()
            mr_data = {}

            # Group data by MR NO and count TCs
            for row in data[2:]:  # Skip header rows
                if row:
                    mr_no = row[2]  # Assuming MR NO is at index 2
                    if mr_no not in mr_data:
                        mr_data[mr_no] = {
                            "DIVISION": row[0],
                            "LOT NO": row[3],
                            "DATE": row[4],
                            "Physical Completed": "NO",  # Default to NO
                            "TC Count": 0  # Initialize TC count
                        }
                    mr_data[mr_no]["TC Count"] += 1  # Increment TC count

                    # Check if columns J to AD (indices 9 to 29) are filled
                    physical_data = row[9:30]  # Get data from columns J to AD
                    if any(cell.strip() for cell in physical_data):  # Check if all cells are filled
                        mr_data[mr_no]["Physical Completed"] = "YES"  # Set to YES if all are filled

            # Insert grouped data into the tree
            for mr_no, details in mr_data.items():
                self.tree.insert('', tk.END, values=(mr_no, details["DIVISION"], details["LOT NO"], details["DATE"], details["Physical Completed"], details["TC Count"], "View"))

            # Bind the view button
            self.tree.bind('<Double-1>', self.on_view)

        except Exception as e:
            messagebox.showerror("Error", f"Error loading data: {str(e)}")

    def on_view(self, event):
        """Handle the view button click."""
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item[0])['values']
            mr_no = item_values[0]  # Get the MR NO
            # Open the physical form for the selected MR NO
            print(mr_no)
            PhysicalVerificationForm(self.master_sheet, mr_no)

def create_physical_verification_form(sheet):
    # master_sheet = get_worksheet(sheet)  # Ensure this returns the correct worksheet object
    PhysicalVerificationForm1(sheet) 