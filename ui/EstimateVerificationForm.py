import tkinter as tk
from tkinter import ttk, messagebox
from config.sheets_setup import get_worksheet
from tkcalendar import DateEntry
import gspread
from gspread.utils import rowcol_to_a1

class EstimateVerification:
    # Define quantity mappings as a class attribute
    QTY_MAPPING = {
        # Physical Inspection Items
        'HT_SIDE_BUSHING': '2',
        'HT_SIDE_METAL': '2',
        'HT_SIDE_CONNECTOR': '3',
        'LT_SIDE_BUSHING': '6',
        'LT_SIDE_METAL': '5',
        'LT_SIDE_CONNECTOR': '4',
        'GUAGE_GLASS': 'Reqd',
        'OUTSIDE_PAINT': 'Reqd',
        'BOLT_NUTS': 'NR',
        'ROD_GASKET': '5',
        'TOP_GASKET': '5',
        'NAME_PLATE': 'NR',
        'BREATHER': 'Reqd',
        # Internal Inspection Items
        'COILS_HT': '',
        'WEIGHT_HT': '',
        # 'CU_ALU_HT': '',
        # 'COILS_LT': '',
        'WEIGHT_LT': '',
        # 'CU_ALU_LT': '',
        'INSIDE_PAINT': '',
        'WASHER_RINGS': '',
        'TESTING_CHARGES': 'Reqd',
        'INSU_MATERIAL': ''
    }

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
                bs_value = row[26]  # BS column from physical data
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

            # Initialize empty data array
            all_data = [['' for _ in range(157)] for _ in range(43)]

            # Fill in the data for each transformer
            for index, entries in enumerate(self.entry_sets):
                # Calculate column positions
                qty_col = index * 5 + 1   # K=1, P=6, U=11 for Qty
                val_col = qty_col + 1     # L=2, Q=7, V=12 for Values

                # Get main transformer data (rows 1-7)
                main_data = [
                    entries['estimates_no'].get(),
                    entries['make'].get(),
                    entries['xmer_sr_no'].get(),
                    entries['rating_kva'].get(),
                    entries['bolted_sealed'].get(),
                    entries['winnding'].get(),
                    entries['se_dpc'].get()
                ]

                # Save main transformer data (rows 1-7)
                for row, value in enumerate(main_data):
                    all_data[row][val_col] = value

                # Get data from master sheet for this transformer
                row_data = self.master_sheet.get_all_values()
                matching_row = None
                for r in row_data[2:]:  # Skip header rows
                    if len(r) > 8 and r[8] == entries['estimates_no'].get():  # Check JOB NO column
                        matching_row = r
                        break

                if matching_row:
                    # Physical Inspection Data (rows 8-20)
                    physical_data = [
                        ('HT_SIDE_BUSHING', 20, 11),    # Column 10 for HT SIDE - BUSHING
                        ('HT_SIDE_METAL', 21, 12),      # Column 11 for HT SIDE - METAL PART
                        ('HT_SIDE_CONNECTOR', 22, 13), # Column 12 for HT SIDE - HT CONNECTOR
                        ('LT_SIDE_BUSHING', 23, 14),   # Column 13 for LT SIDE - BUSHING
                        ('LT_SIDE_METAL', 24, 15),     # Column 14 for LT SIDE - METAL PART
                        ('LT_SIDE_CONNECTOR', 25, 16), # Column 15 for LT SIDE - LT CONNECTOR
                        ('GUAGE_GLASS', 18, 17),       # Column 16 for GUAGE GLASS
                        ('OUTSIDE_PAINT', 19, 20),     # Column 19 for OUTSIDE PAINT
                        ('BOLT_NUTS', 17, 21),         # Column 20 for BOLT & NUTS
                        ('ROD_GASKET', 11, 22),        # Column 21 for ROD GASKET
                        ('TOP_GASKET', 10, 23),        # Column 22 for TOP GASKET
                        ('NAME_PLATE', 27, 24),        # Column 23 for NAME PLATE
                        ('BREATHER', 26, 25)           # Column 24 for BREATHER
                    ]

                    # Save physical inspection data with Qty values
                    for key, row, col in physical_data:
                        qty = matching_row[col-1] if col-1 < len(matching_row) else ''
                        if qty:
                            all_data[row][qty_col] = qty

                    # Internal Inspection Data (rows 24-33)
                    internal_data = [
                        ('COILS_HT', 30, 38),          # Column 37 for COILS_HT
                        ('WEIGHT_HT', 31, 40),         # Column 39 for WEIGHT_HT
                        ('WEIGHT_LT', 36, 41),         # Column 40 for WEIGHT_LT
                        ('INSIDE_PAINT', 41, 43),      # Column 42 for INSIDE_PAINT
                        ('WASHER_RINGS', 42, 42),      # Column 41 for WASHER_RINGS
                        ('TESTING_CHARGES', 8, 45),   # Column 44 for TESTING_CHARGES
                        ('INSU_MATERIAL', 12, 44)      # Column 43 for INSU_MATERIAL
                    ]

                    # Save internal inspection data with Qty values
                    for key, row, col in internal_data:
                        if col > 0:  # Skip if column number is 0 (placeholder)
                            qty = matching_row[col-1] if col-1 < len(matching_row) else ''
                            if qty:
                                all_data[row][qty_col] = qty

                    # Add static data for Qty Values section
                    static_qty_data = [
                        (7, "Qty"),   # Row 9 (index 8) - K Column
                        (9, "Reqd"),   # Row 9 (index 8) - K Column
                        (13, "Reqd"),  # Row 13 (index 12) - K Column
                        (14, "Reqd")   # Row 14 (index 13) - K Column
                    ]
                    
                    # Save static Qty values
                    for row, value in static_qty_data:
                        all_data[row][qty_col] = value

                    # HT and LT Coils Data (rows 31 and 36)
                    # Calculate column positions for coils data
                    # For first transformer: L=2, M=3, N=4
                    # For second transformer: Q=7, R=8, S=9
                    # For third transformer: V=12, W=13, X=14
                    coils_base_col = val_col  # Start from L, Q, V
                    
                    # Add static A, B, C labels for HT Coils (Row 30)
                    static_ht_labels = ["A", "B", "C"]
                    for i, label in enumerate(static_ht_labels):
                        all_data[29][coils_base_col + i] = label  # Row 30 (index 29)
                    
                    # HT Coils Data (Row 31)
                    # AF=31, AG=32, AH=33 for HT Coils A, B, C
                    ht_coils_data = [
                        (30, 32),  # HT Coil A (AF)
                        (30, 33),  # HT Coil B (AG)
                        (30, 34)   # HT Coil C (AH)
                    ]
                    
                    # Add static A, B, C labels for LT Coils (Row 35)
                    static_lt_labels = ["A", "B", "C"]
                    for i, label in enumerate(static_lt_labels):
                        all_data[34][coils_base_col + i] = label  # Row 35 (index 34)
                    
                    # LT Coils Data (Row 36)
                    # AI=34, AJ=35, AK=36 for LT Coils A, B, C
                    lt_coils_data = [
                        (35, 35),  # LT Coil A (AI)
                        (35, 36),  # LT Coil B (AJ)
                        (35, 37)   # LT Coil C (AK)
                    ]

                    # Save HT Coils data
                    for row, col in ht_coils_data:
                        value = matching_row[col-1] if col-1 < len(matching_row) else ''
                        if value:
                            all_data[row][coils_base_col] = value
                            coils_base_col += 1

                    # Reset coils_base_col for LT Coils
                    coils_base_col = val_col

                    # Save LT Coils data
                    for row, col in lt_coils_data:
                        value = matching_row[col-1] if col-1 < len(matching_row) else ''
                        if value:
                            all_data[row][coils_base_col] = value
                            coils_base_col += 1

            # Update the sheet in one batch
            self.estimate_sheet.update('J1:FR43', all_data)
            messagebox.showinfo("Success", "Data saved to ESTIMATE sheet successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Error saving data to ESTIMATE sheet: {str(e)}")
            print(f"Debug - Error details: {str(e)}") 