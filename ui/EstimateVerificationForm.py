import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import get_worksheet
from tkcalendar import DateEntry
import gspread
from gspread.utils import rowcol_to_a1

class EstimateVerification:
    def __init__(self, sheet, mr_no=None):
        self.mr_no = mr_no
        self.window = tk.Toplevel()
        self.window.title("Estimate Details")
        self.window.geometry("800x600")
        
        # Store the master sheet reference
        self.master_sheet = sheet
        
        # Get the ESTIMATE sheet
        try:
            self.estimate_sheet = self.master_sheet.spreadsheet.worksheet("ESTIMATE")
        except gspread.WorksheetNotFound:
            messagebox.showerror("Error", "ESTIMATE sheet not found in the spreadsheet")
            self.window.destroy()
            return
        
        # Create main container
        self.main_container = ttk.Frame(self.window)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create scrollable canvas
        self.canvas = tk.Canvas(self.main_container)
        self.scrollbar = ttk.Scrollbar(self.main_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack the scrollbar and canvas
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Store entry widgets in a list
        self.entry_sets = []
        
        # Add save button
        save_button = ttk.Button(
            self.main_container,
            text="Save to ESTIMATE Sheet",
            command=self.save_to_estimate_sheet
        )
        save_button.pack(pady=10)
        
        # Load data for the given MR NO
        self.load_data(mr_no)

    def create_entry_set(self, index):
        """Create a set of entries for a single TC"""
        # Create frame for this set of entries
        details_frame = ttk.LabelFrame(
            self.scrollable_frame,
            text=f"Transformer {index + 1}",
            padding=10
        )
        details_frame.pack(fill=tk.X, expand=True, pady=(0, 20))

        # Create the entries with headers
        headers = [
            ("ESTIMATES NO.", ""),
            ("Make", ""),
            ("X'mer Sr No", ""),
            ("Rating KVA", ""),
            ("Bolted/Sealed", ""),
            ("Winnding", ""),
            ("SE/DPC", "")
        ]

        entries = {}
        # Create grid for headers and values
        for row, (header, _) in enumerate(headers):
            # Header label (Column 0)
            header_label = ttk.Label(
                details_frame,
                text=header,
                font=("Arial", 10, "bold")
            )
            header_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")

            # Value entry (Column 1)
            value_entry = ttk.Entry(
                details_frame,
                width=30,
                state='readonly'
            )
            value_entry.grid(row=row, column=1, padx=10, pady=5, sticky="w")
            
            # Store entry widget in dictionary
            attr_name = header.lower().replace(" ", "_").replace("/", "_").replace("'", "").replace(".", "")
            entries[attr_name] = value_entry

        return entries

    def load_data(self, mr_no):
        """Load data from master sheet for the given MR NO"""
        try:
            if not mr_no:
                return

            # Clear existing entry sets
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.entry_sets.clear()

            # Get data from master sheet
            data = self.master_sheet.get_all_values()
            
            # Find all rows with matching MR NO
            matching_rows = []
            for row in data[2:]:  # Skip header rows
                if row[2] == str(mr_no):
                    matching_rows.append(row)

            # Create entry sets for each matching row
            for index, row in enumerate(matching_rows):
                entries = self.create_entry_set(index)
                self.entry_sets.append(entries)
                
                # Fill in the data
                entries['estimates_no'].configure(state='normal')
                entries['estimates_no'].delete(0, tk.END)
                entries['estimates_no'].insert(0, row[8])  # JOB NO
                entries['estimates_no'].configure(state='readonly')
                
                entries['make'].configure(state='normal')
                entries['make'].delete(0, tk.END)
                entries['make'].insert(0, row[6])  # MAKE
                entries['make'].configure(state='readonly')
                
                entries['xmer_sr_no'].configure(state='normal')
                entries['xmer_sr_no'].delete(0, tk.END)
                entries['xmer_sr_no'].insert(0, row[5])  # TC NO
                entries['xmer_sr_no'].configure(state='readonly')
                
                entries['rating_kva'].configure(state='normal')
                entries['rating_kva'].delete(0, tk.END)
                entries['rating_kva'].insert(0, row[7])  # TC CAPACITY
                entries['rating_kva'].configure(state='readonly')
                
                # Get B/S value from physical inspection data
                bs_value = row[25]  # BS column from physical data
                bolted_sealed = "BOLTED" if bs_value == "B" else "SEALED"
                entries['bolted_sealed'].configure(state='normal')
                entries['bolted_sealed'].delete(0, tk.END)
                entries['bolted_sealed'].insert(0, bolted_sealed)
                entries['bolted_sealed'].configure(state='readonly')
                
                # Get winding type from internal inspection data
                winding_value = row[35]  # CU/ALU column from internal data
                winding = "CU" if winding_value == "CU" else "AL"
                entries['winnding'].configure(state='normal')
                entries['winnding'].delete(0, tk.END)
                entries['winnding'].insert(0, winding)
                entries['winnding'].configure(state='readonly')

                # Get SE/DPC value safely
                se_dpc_value = row[45] 
                entries['se_dpc'].configure(state='normal')
                entries['se_dpc'].delete(0, tk.END)
                entries['se_dpc'].insert(0, se_dpc_value)
                entries['se_dpc'].configure(state='readonly')

            # If no matching rows found
            if not matching_rows:
                messagebox.showinfo("Info", f"No data found for MR NO: {mr_no}")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading data: {str(e)}")
            print(f"Debug - Error details: {str(e)}")

    def save_to_estimate_sheet(self):
        """Save the transformer details to the ESTIMATE sheet"""
        try:
            if not self.entry_sets:
                messagebox.showinfo("Info", "No data to save")
                return

            # Prepare batch update for clearing data
            batch_clear_data = []
            for row in range(1, 8):  # Rows 1 to 7
                row_data = [""] * 157  # 157 columns (168-12+1)
                batch_clear_data.append(row_data)
            
            # Clear all data in one batch update
            range_to_clear = f"L1:FR7"
            self.estimate_sheet.update(range_to_clear, batch_clear_data)

            # Prepare batch update for new data
            all_data = []
            max_col = 12  # Start from column L

            # Initialize the data array with empty strings
            for row in range(7):  # 7 rows for each transformer's data
                all_data.append([""] * (157))  # 157 columns total

            # Fill in the data for each transformer
            for index, entries in enumerate(self.entry_sets):
                col_index = index * 5  # Each transformer takes 3 columns
                
                # Get data for this transformer
                transformer_data = [
                    entries['estimates_no'].get(),
                    entries['make'].get(),
                    entries['xmer_sr_no'].get(),
                    entries['rating_kva'].get(),
                    entries['bolted_sealed'].get(),
                    entries['winnding'].get(),
                    entries['se_dpc'].get()
                ]

                # Fill in the data in the correct positions
                for row, value in enumerate(transformer_data):
                    # Fill the same value in 3 consecutive columns
                    for i in range(3):
                        if col_index + i < 157:  # Ensure we don't exceed array bounds
                            all_data[row][col_index + i] = value

            # Update all data in one batch
            range_to_update = f"L1:FR7"
            self.estimate_sheet.update(range_to_update, all_data)

            messagebox.showinfo("Success", "Data saved to ESTIMATE sheet successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data to ESTIMATE sheet: {str(e)}")
            print(f"Debug - Error details: {str(e)}") 