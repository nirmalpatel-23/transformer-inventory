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
            all_data = [['' for _ in range(157)] for _ in range(50)]

            # Add static "Amount" to L8, Q8, V8 for each transformer
            for index in range(len(self.entry_sets)):
                # Calculate the column position for this transformer
                # L=2, Q=7, V=12 (incrementing by 5 for each transformer)
                amount_col = 2 + (index * 5)
                all_data[7][amount_col] = "Amount"  # Row 8 (index 7)

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

                # Save main transformer data (rows 1-7) to column K (index 1)
                for row, value in enumerate(main_data):
                    all_data[row][qty_col] = value  # Store in K column (one column to the left)

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

                    # Add conditional BOLT_NUTS value based on column 26
                    bolt_nuts_value = "Reqd" if matching_row[26] == "S" else "NR"  # Column 26 (index 25)
                    all_data[16][qty_col] = bolt_nuts_value  # Row 16 (index 15)

                    # Add static data to column J for specific rows
                    all_data[7][qty_col - 1] = "Rate"  # Row 16 (index 15)
                    all_data[16][qty_col - 1] = "1452"  # Row 16 (index 15)
                    all_data[26][qty_col - 1] = "309"   # Row 26 (index 25)
                    all_data[27][qty_col - 1] = "143"   # Row 27 (index 26)

                    # Add capacity-based rate formulas in column J (and corresponding columns)
                    # Calculate the formula column (one column before Qty column)
                    formula_col = qty_col - 1  # J=0, O=5, T=10
                    
                    # Get the transformer capacity from the entry
                    capacity = entries['rating_kva'].get()
                    
                    # Get values from master sheet columns AM (index 38) and AT (index 45)
                    am_value = matching_row[38] if len(matching_row) > 38 else ''  # Column AM
                    at_value = matching_row[45] if len(matching_row) > 45 else ''  # Column AT
                    
                    # Debug prints
                    print(f"Transformer {index + 1} - AM Value:", am_value, "Type:", type(am_value))
                    print(f"Transformer {index + 1} - AT Value:", at_value, "Type:", type(at_value))
                    print(f"Transformer {index + 1} - AM Value stripped:", am_value.strip(), "AT Value stripped:", at_value.strip())
                    
                    # Get values from ESTIMATE sheet for C31, C33, C34, F34, C38, C39, C40, C41, and D36
                    estimate_row_data = self.estimate_sheet.get_all_values()
                    c31_value = estimate_row_data[30][2] if len(estimate_row_data) > 30 else ''  # C31 is row 31, column C (index 2)
                    c33_value = estimate_row_data[32][2] if len(estimate_row_data) > 32 else ''  # C33 is row 33, column C (index 2)
                    c34_value = estimate_row_data[33][2] if len(estimate_row_data) > 33 else ''  # C34 is row 34, column C (index 2)
                    f34_value = estimate_row_data[33][5] if len(estimate_row_data) > 33 else ''  # F34 is row 34, column F (index 5)
                    c38_value = estimate_row_data[37][2] if len(estimate_row_data) > 37 else ''  # C38 is row 38, column C (index 2)
                    c39_value = estimate_row_data[38][2] if len(estimate_row_data) > 38 else ''  # C39 is row 39, column C (index 2)
                    c40_value = estimate_row_data[39][2] if len(estimate_row_data) > 39 else ''  # C40 is row 40, column C (index 2)
                    c41_value = estimate_row_data[40][2] if len(estimate_row_data) > 40 else ''  # C41 is row 41, column C (index 2)
                    d36_value = estimate_row_data[35][3] if len(estimate_row_data) > 35 else ''  # D36 is row 36, column D (index 3)
                    
                    print(f"Transformer {index + 1} - C31 Value:", c31_value)
                    print(f"Transformer {index + 1} - C33 Value:", c33_value)
                    print(f"Transformer {index + 1} - C34 Value:", c34_value)
                    print(f"Transformer {index + 1} - F34 Value:", f34_value)
                    print(f"Transformer {index + 1} - C38 Value:", c38_value)
                    print(f"Transformer {index + 1} - C39 Value:", c39_value)
                    print(f"Transformer {index + 1} - C40 Value:", c40_value)
                    print(f"Transformer {index + 1} - C41 Value:", c41_value)
                    print(f"Transformer {index + 1} - D36 Value:", d36_value)
                    
                    # Calculate the target column for this transformer
                    # First transformer: J (index 0), Second: O (index 5), Third: T (index 10)
                    target_col = index * 5  # 0 for first, 5 for second, 10 for third
                    
                    # Store the J33, J34, J38, J39, J40, and J41 values in variables
                    j33_value = None
                    j34_value = None
                    j38_value = None
                    j39_value = None
                    j40_value = None
                    j41_value = None
                    
                    # Set J33 value based on conditions
                    if am_value.strip().upper() == "AL" and at_value.strip().upper() == "DPC":
                        print(f"Transformer {index + 1} - Condition met: Setting J33 to C33 value")
                        j33_value = c33_value  # Store C33 value
                        print(f"Transformer {index + 1} - J33 will be set to:", j33_value)
                    else:
                        print(f"Transformer {index + 1} - Condition not met: Setting J33 to C31 value")
                        j33_value = c31_value  # Store C31 value
                        print(f"Transformer {index + 1} - J33 will be set to:", j33_value)
                    
                    # Set J34 value based on AM value condition
                    if am_value.strip().upper() == "AL":
                        print(f"Transformer {index + 1} - AM is AL: Setting J34 to C34 value")
                        j34_value = c34_value  # Store C34 value
                        print(f"Transformer {index + 1} - J34 will be set to:", j34_value)
                    else:
                        print(f"Transformer {index + 1} - AM is not AL: Setting J34 to F34 value")
                        j34_value = f34_value  # Store F34 value
                        print(f"Transformer {index + 1} - J34 will be set to:", j34_value)
                    
                    # Set J38 value based on AM and AT value conditions
                    if am_value.strip().upper() == "AL" and at_value.strip().upper() == "DPC":
                        print(f"Transformer {index + 1} - AM is AL and AT is DPC: Setting J38 to C38 value")
                        j38_value = c38_value  # Store C38 value
                        print(f"Transformer {index + 1} - J38 will be set to:", j38_value)
                    else:
                        print(f"Transformer {index + 1} - Conditions not met: Setting J38 to D36 value")
                        j38_value = d36_value  # Store D36 value
                        print(f"Transformer {index + 1} - J38 will be set to:", j38_value)
                    
                    # Set J39 value based on AM and AT value conditions
                    if am_value.strip().upper() == "AL" and at_value.strip().upper() == "DPC":
                        print(f"Transformer {index + 1} - AM is AL and AT is DPC: Setting J39 to C39 value")
                        j39_value = c39_value  # Store C39 value
                        print(f"Transformer {index + 1} - J39 will be set to:", j39_value)
                    else:
                        print(f"Transformer {index + 1} - Conditions not met: Setting J39 to D36 value")
                        j39_value = d36_value  # Store D36 value
                        print(f"Transformer {index + 1} - J39 will be set to:", j39_value)
                    
                    # Set J40 value based on AM and AT value conditions
                    if am_value.strip().upper() == "AL" and at_value.strip().upper() == "DPC":
                        print(f"Transformer {index + 1} - AM is AL and AT is DPC: Setting J40 to C40 value")
                        j40_value = c40_value  # Store C40 value
                        print(f"Transformer {index + 1} - J40 will be set to:", j40_value)
                    else:
                        print(f"Transformer {index + 1} - Conditions not met: Setting J40 to 0")
                        j40_value = "0"  # Store 0
                        print(f"Transformer {index + 1} - J40 will be set to:", j40_value)
                    
                    # Set J41 value based on AM and AT value conditions
                    if am_value.strip().upper() == "AL" and at_value.strip().upper() == "DPC":
                        print(f"Transformer {index + 1} - AM is AL and AT is DPC: Setting J41 to C41 value")
                        j41_value = c41_value  # Store C41 value
                        print(f"Transformer {index + 1} - J41 will be set to:", j41_value)
                    else:
                        print(f"Transformer {index + 1} - Conditions not met: Setting J41 to 0")
                        j41_value = "0"  # Store 0
                        print(f"Transformer {index + 1} - J41 will be set to:", j41_value)
                    
                    # Set the values in the appropriate columns
                    all_data[32][target_col] = j33_value  # Set J33/O33/T33
                    all_data[33][target_col] = j34_value  # Set J34/O34/T34
                    all_data[37][target_col] = j38_value  # Set J38/O38/T38
                    all_data[38][target_col] = j39_value  # Set J39/O39/T39
                    all_data[39][target_col] = j40_value  # Set J40/O40/T40
                    all_data[40][target_col] = j41_value  # Set J41/O41/T41
                    
                    print(f"Transformer {index + 1} - Final J33 value being set in column {target_col}:", j33_value)
                    print(f"Transformer {index + 1} - Final J34 value being set in column {target_col}:", j34_value)
                    print(f"Transformer {index + 1} - Final J38 value being set in column {target_col}:", j38_value)
                    print(f"Transformer {index + 1} - Final J39 value being set in column {target_col}:", j39_value)
                    print(f"Transformer {index + 1} - Final J40 value being set in column {target_col}:", j40_value)
                    print(f"Transformer {index + 1} - Final J41 value being set in column {target_col}:", j41_value)
                    
                    mappingObj = {
                        "200 KVA": 7,
                        "100 KVA": 6,
                        "75 KVA": 5,
                        "63 KVA": 5,
                        "50 KVA": 5,
                        "25 KVA": 4,
                        "16 KVA": 3,
                        "10 KVA": 3,
                        "5 KVA": 2,
                    }
                                        
                    # Define the rows where formulas should be added
                    formula_rows = list(range(8, 15)) + list(range(17, 26)) + [41] + [42]  # Rows 9-15, 18-26, and 42
                    
                    estimate_row_data = self.estimate_sheet.get_all_values()
                    
                    # Create the formula for each row
                    for row in formula_rows:
                        print('row', row)
                        rate_row = row + 1
                        total = 0
                        for capacity, col_index in mappingObj.items():
                            if entries['rating_kva'].get() == capacity:
                                print(estimate_row_data[row][col_index])
                                # print(matching_row, (col_letter))
                                # print(matching_row, matching_row[col_letter], (col_letter))
                                rate_value = estimate_row_data[row][col_index] or '0'
                                print(rate_value)
                                total = rate_value if rate_value else '0'
                        all_data[row][qty_col - 1] = total  # Save the calculated total in the column before Qty


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

            # After all transformers are processed, apply the "Reqd" logic
            print("\nApplying 'Reqd' logic to all transformers...")
            for index, entries in enumerate(self.entry_sets):
                # Calculate column positions for this transformer
                qty_col = index * 5 + 1   # K=1, P=6, U=11 for Qty
                val_col = qty_col + 1     # L=2, Q=7, V=12 for Values
                
                # Define the rows that need the "Reqd" logic
                reqd_logic_rows = [8, 9, 12, 13, 14, 16, 17, 18, 19, 26, 27, 41]  # 0-based indices
                
                print(f"\nProcessing Transformer {index + 1}:")
                for row in reqd_logic_rows:
                    # Get the value from K column (qty_col)
                    k_value = all_data[row][qty_col]
                    # Get the value from J column (qty_col - 1)
                    j_value = all_data[row][qty_col - 1]
                    
                    print(f"Row {row + 1}:")
                    print(f"K column value: '{k_value}'")
                    print(f"J column value: '{j_value}'")
                    
                    # Apply the logic
                    if k_value and k_value.strip().upper() == "REQD":
                        print(f"K{row + 1} is 'Reqd', setting L{row + 1} to J{row + 1} value: {j_value}")
                        all_data[row][val_col] = j_value
                    else:
                        print(f"K{row + 1} is not 'Reqd', setting L{row + 1} to 0")
                        all_data[row][val_col] = "0"
                    
                    print(f"Final value set in L{row + 1}: {all_data[row][val_col]}")

                # Calculate K33 for this transformer
                print(f"\nCalculating K33 for Transformer {index + 1}:")
                try:
                    # Get values from L31, M31, N31 (columns val_col, val_col+1, val_col+2)
                    l31_value = all_data[30][val_col] or '0'     # L31
                    m31_value = all_data[30][val_col + 1] or '0' # M31
                    n31_value = all_data[30][val_col + 2] or '0' # N31
                    
                    # Get value from K32
                    k32_value = all_data[31][qty_col] or '0'     # K32
                    
                    print(f"L31 value: {l31_value}")
                    print(f"M31 value: {m31_value}")
                    print(f"N31 value: {n31_value}")
                    print(f"K32 value: {k32_value}")
                    
                    # Convert values to float and calculate sum
                    l31_num = float(l31_value)
                    m31_num = float(m31_value)
                    n31_num = float(n31_value)
                    k32_num = float(k32_value)
                    
                    # Calculate sum of L31, M31, N31
                    sum_lmn = l31_num + m31_num + n31_num
                    print(f"Sum of L31+M31+N31: {sum_lmn}")
                    
                    # Calculate final result and round to 2 decimal places
                    k33_result = round(sum_lmn * k32_num, 2)
                    print(f"K33 result (sum * K32): {k33_result}")
                    
                    # Store result in K33
                    all_data[32][qty_col] = str(k33_result)
                    print(f"Final value set in K33: {k33_result}")
                    
                    # Calculate L33 (K33 * J33)
                    print(f"\nCalculating L33 for Transformer {index + 1}:")
                    try:
                        # Get value from J33 (qty_col - 1)
                        j33_value = all_data[32][qty_col - 1] or '0'
                        print(f"J33 value: {j33_value}")
                        
                        # Convert J33 to float and calculate L33
                        j33_num = float(j33_value)
                        l33_result = round(k33_result * j33_num, 2)
                        print(f"L33 result (K33 * J33): {l33_result}")
                        
                        # Store result in L33
                        all_data[32][val_col] = str(l33_result)
                        print(f"Final value set in L33: {l33_result}")
                        
                        # Calculate L34 (K33 * J34)
                        print(f"\nCalculating L34 for Transformer {index + 1}:")
                        try:
                            # Get value from J34 (qty_col - 1)
                            j34_value = all_data[33][qty_col - 1] or '0'
                            print(f"J34 value: {j34_value}")
                            
                            # Convert J34 to float and calculate L34
                            j34_num = float(j34_value)
                            l34_result = round(k33_result * j34_num, 2)
                            print(f"L34 result (K33 * J34): {l34_result}")
                            
                            # Store result in L34
                            all_data[33][val_col] = str(l34_result)
                            print(f"Final value set in L34: {l34_result}")
                            
                            # Calculate K38 (count of 'R' in L36, M36, N36)
                            print(f"\nCalculating K38 for Transformer {index + 1}:")
                            try:
                                # Get values from L36, M36, N36 (columns val_col, val_col+1, val_col+2)
                                l36_value = all_data[35][val_col] or ''     # L36
                                m36_value = all_data[35][val_col + 1] or '' # M36
                                n36_value = all_data[35][val_col + 2] or '' # N36
                                
                                print(f"L36 value: '{l36_value}'")
                                print(f"M36 value: '{m36_value}'")
                                print(f"N36 value: '{n36_value}'")
                                
                                # Count 'R' values
                                r_count = 0
                                if l36_value.strip().upper() == 'R':
                                    r_count += 1
                                if m36_value.strip().upper() == 'R':
                                    r_count += 1
                                if n36_value.strip().upper() == 'R':
                                    r_count += 1
                                
                                print(f"Number of 'R' values found: {r_count}")
                                
                                # Store result in K38
                                all_data[37][qty_col] = str(r_count)
                                print(f"Final value set in K38: {r_count}")
                                
                                # Calculate K39 (count of 'D' in L36, M36, N36)
                                print(f"\nCalculating K39 for Transformer {index + 1}:")
                                try:
                                    # Count 'D' values
                                    d_count = 0
                                    if l36_value.strip().upper() == 'D':
                                        d_count += 1
                                    if m36_value.strip().upper() == 'D':
                                        d_count += 1
                                    if n36_value.strip().upper() == 'D':
                                        d_count += 1
                                    
                                    print(f"Number of 'D' values found: {d_count}")
                                    
                                    # Store result in K39
                                    all_data[38][qty_col] = str(d_count)
                                    print(f"Final value set in K39: {d_count}")
                                    
                                    # Calculate K40 (count of 'R' in L36, M36, N36)
                                    print(f"\nCalculating K40 for Transformer {index + 1}:")
                                    try:
                                        # Count 'R' values
                                        r_count_k40 = 0
                                        if l36_value.strip().upper() == 'R':
                                            r_count_k40 += 1
                                        if m36_value.strip().upper() == 'R':
                                            r_count_k40 += 1
                                        if n36_value.strip().upper() == 'R':
                                            r_count_k40 += 1
                                        
                                        print(f"Number of 'R' values found for K40: {r_count_k40}")
                                        
                                        # Store result in K40
                                        all_data[39][qty_col] = str(r_count_k40)
                                        print(f"Final value set in K40: {r_count_k40}")
                                        
                                        # Calculate K41 (count of 'D' in L36, M36, N36)
                                        print(f"\nCalculating K41 for Transformer {index + 1}:")
                                        try:
                                            # Count 'D' values
                                            d_count_k41 = 0
                                            if l36_value.strip().upper() == 'D':
                                                d_count_k41 += 1
                                            if m36_value.strip().upper() == 'D':
                                                d_count_k41 += 1
                                            if n36_value.strip().upper() == 'D':
                                                d_count_k41 += 1
                                            
                                            print(f"Number of 'D' values found for K41: {d_count_k41}")
                                            
                                            # Store result in K41
                                            all_data[40][qty_col] = str(d_count_k41)
                                            print(f"Final value set in K41: {d_count_k41}")
                                            
                                            # Calculate L38, L39, L40, L41
                                            print(f"\nCalculating L38-L41 for Transformer {index + 1}:")
                                            try:
                                                # Get K37 value
                                                k37_value = all_data[36][qty_col] or '0'
                                                print(f"K37 value: {k37_value}")
                                                
                                                # Convert K37 to float
                                                k37_num = float(k37_value)
                                                
                                                # Calculate L38 (J38 * K38 * K37)
                                                j38_value = all_data[37][qty_col - 1] or '0'
                                                k38_value = all_data[37][qty_col] or '0'
                                                l38_result = round(float(j38_value) * float(k38_value) * k37_num, 2)
                                                all_data[37][val_col] = str(l38_result)
                                                print(f"L38 result (J38 * K38 * K37): {l38_result}")
                                                
                                                # Calculate L39 (J39 * K39 * K37)
                                                j39_value = all_data[38][qty_col - 1] or '0'
                                                k39_value = all_data[38][qty_col] or '0'
                                                l39_result = round(float(j39_value) * float(k39_value) * k37_num, 2)
                                                all_data[38][val_col] = str(l39_result)
                                                print(f"L39 result (J39 * K39 * K37): {l39_result}")
                                                
                                                # Calculate L40 (J40 * K40 * K37)
                                                j40_value = all_data[39][qty_col - 1] or '0'
                                                k40_value = all_data[39][qty_col] or '0'
                                                l40_result = round(float(j40_value) * float(k40_value) * k37_num, 2)
                                                all_data[39][val_col] = str(l40_result)
                                                print(f"L40 result (J40 * K40 * K37): {l40_result}")
                                                
                                                # Calculate L41 (J41 * K41 * K37)
                                                j41_value = all_data[40][qty_col - 1] or '0'
                                                k41_value = all_data[40][qty_col] or '0'
                                                l41_result = round(float(j41_value) * float(k41_value) * k37_num, 2)
                                                all_data[40][val_col] = str(l41_result)
                                                print(f"L41 result (J41 * K41 * K37): {l41_result}")
                                                
                                            except ValueError as e:
                                                print(f"Error calculating L38-L41: {str(e)}")
                                                all_data[37][val_col] = "0"  # Set L38 to 0
                                                all_data[38][val_col] = "0"  # Set L39 to 0
                                                all_data[39][val_col] = "0"  # Set L40 to 0
                                                all_data[40][val_col] = "0"  # Set L41 to 0
                                            
                                        except Exception as e:
                                            print(f"Error calculating K41: {str(e)}")
                                            all_data[40][qty_col] = "0"  # Set to 0 if there's any error
                                            all_data[37][val_col] = "0"  # Also set L38 to 0
                                            all_data[38][val_col] = "0"  # Also set L39 to 0
                                            all_data[39][val_col] = "0"  # Also set L40 to 0
                                            all_data[40][val_col] = "0"  # Also set L41 to 0
                                            
                                    except Exception as e:
                                        print(f"Error calculating K40: {str(e)}")
                                        all_data[39][qty_col] = "0"  # Set to 0 if there's any error
                                        all_data[40][qty_col] = "0"  # Also set K41 to 0 if K40 calculation failed
                                        all_data[37][val_col] = "0"  # Also set L38 to 0
                                        all_data[38][val_col] = "0"  # Also set L39 to 0
                                        all_data[39][val_col] = "0"  # Also set L40 to 0
                                        all_data[40][val_col] = "0"  # Also set L41 to 0
                                        
                                except Exception as e:
                                    print(f"Error calculating K39: {str(e)}")
                                    all_data[38][qty_col] = "0"  # Set to 0 if there's any error
                                    all_data[39][qty_col] = "0"  # Also set K40 to 0 if K39 calculation failed
                                    all_data[40][qty_col] = "0"  # Also set K41 to 0 if K39 calculation failed
                                    all_data[37][val_col] = "0"  # Also set L38 to 0
                                    all_data[38][val_col] = "0"  # Also set L39 to 0
                                    all_data[39][val_col] = "0"  # Also set L40 to 0
                                    all_data[40][val_col] = "0"  # Also set L41 to 0
                                    
                            except Exception as e:
                                print(f"Error calculating K38: {str(e)}")
                                all_data[37][qty_col] = "0"  # Set to 0 if there's any error
                                all_data[38][qty_col] = "0"  # Also set K39 to 0 if K38 calculation failed
                                all_data[39][qty_col] = "0"  # Also set K40 to 0 if K38 calculation failed
                                all_data[40][qty_col] = "0"  # Also set K41 to 0 if K38 calculation failed
                                all_data[37][val_col] = "0"  # Also set L38 to 0
                                all_data[38][val_col] = "0"  # Also set L39 to 0
                                all_data[39][val_col] = "0"  # Also set L40 to 0
                                all_data[40][val_col] = "0"  # Also set L41 to 0
                                
                        except ValueError as e:
                            print(f"Error calculating L34: {str(e)}")
                            all_data[33][val_col] = "0"  # Set to 0 if there's any error
                            all_data[37][qty_col] = "0"  # Also set K38 to 0 if L34 calculation failed
                            all_data[38][qty_col] = "0"  # Also set K39 to 0 if L34 calculation failed
                            all_data[39][qty_col] = "0"  # Also set K40 to 0 if L34 calculation failed
                            all_data[40][qty_col] = "0"  # Also set K41 to 0 if L34 calculation failed
                            
                    except ValueError as e:
                        print(f"Error calculating L33: {str(e)}")
                        all_data[32][val_col] = "0"  # Set to 0 if there's any error
                        all_data[33][val_col] = "0"  # Also set L34 to 0 if L33 calculation failed
                        all_data[37][qty_col] = "0"  # Also set K38 to 0 if L33 calculation failed
                        all_data[38][qty_col] = "0"  # Also set K39 to 0 if L33 calculation failed
                        all_data[39][qty_col] = "0"  # Also set K40 to 0 if L33 calculation failed
                        all_data[40][qty_col] = "0"  # Also set K41 to 0 if L33 calculation failed
                        
                except ValueError as e:
                    print(f"Error calculating K33: {str(e)}")
                    all_data[32][qty_col] = "0"  # Set to 0 if there's any error
                    all_data[32][val_col] = "0"  # Also set L33 to 0 if K33 calculation failed
                    all_data[33][val_col] = "0"  # Also set L34 to 0 if K33 calculation failed
                    all_data[37][qty_col] = "0"  # Also set K38 to 0 if K33 calculation failed
                    all_data[38][qty_col] = "0"  # Also set K39 to 0 if K33 calculation failed
                    all_data[39][qty_col] = "0"  # Also set K40 to 0 if K33 calculation failed
                    all_data[40][qty_col] = "0"  # Also set K41 to 0 if K33 calculation failed

            # Apply multiplication logic for specific cells
            print("\nApplying multiplication logic for specific cells...")
            for index, entries in enumerate(self.entry_sets):
                # Calculate column positions for this transformer
                qty_col = index * 5 + 1   # K=1, P=6, U=11 for Qty
                val_col = qty_col + 1     # L=2, Q=7, V=12 for Values
                
                # Define the rows that need multiplication logic
                multiply_rows = [10, 11, 20, 21, 22, 23, 24, 25, 42]  # 0-based indices for rows 11,12,21,22,23,24,25,26,43
                
                print(f"\nProcessing multiplication for Transformer {index + 1}:")
                for row in multiply_rows:
                    # Get the value from J column (qty_col - 1)
                    j_value = all_data[row][qty_col - 1]
                    # Get the value from K column (qty_col)
                    k_value = all_data[row][qty_col]
                    
                    print(f"Row {row + 1}:")
                    print(f"J column value: '{j_value}'")
                    print(f"K column value: '{k_value}'")
                    
                    # Convert values to float for multiplication, defaulting to 0 if invalid
                    try:
                        j_num = float(j_value) if j_value else 0
                        k_num = float(k_value) if k_value else 0
                        result = round(j_num * k_num, 2)  # Round to 2 decimal places
                        print(f"Multiplying {j_num} * {k_num} = {result}")
                        all_data[row][val_col] = str(result)
                    except ValueError:
                        print(f"Invalid number format, setting to 0")
                        all_data[row][val_col] = "0"
                    
                    print(f"Final value set in L{row + 1}: {all_data[row][val_col]}")

            # After all other calculations and before the final sheet update
            print("\nCalculating final sums for L16 (L9-L15) and L29 (L17-L28)...")
            for index, entries in enumerate(self.entry_sets):
                try:
                    # Calculate column positions for this transformer
                    # For first transformer: L=2, second: Q=7, third: V=12, etc.
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Calculate sum for L16 (L9 to L15)
                    sum_l16 = 0
                    print(f"\nTransformer {index + 1} - Final sum calculation for L16 (L9-L15):")
                    print(f"Using column index: {val_col} for transformer {index + 1}")
                    
                    # Get all values from L9 to L15
                    for row in range(8, 15):  # Rows 9-15 (0-based indices 8-14)
                        value = all_data[row][val_col] or '0'
                        try:
                            num_value = float(value)
                            print(f"Row {row + 1} value: {value} (converted to {num_value})")
                            sum_l16 += num_value
                        except ValueError:
                            print(f"Warning: Invalid number format in row {row + 1}: '{value}', using 0")
                            sum_l16 += 0
                    
                    # Store the final sum in L16
                    all_data[15][val_col] = str(round(sum_l16, 2))
                    print(f"Final sum for L16: {sum_l16}")
                    
                    # Calculate sum for L29 (L17 to L28)
                    sum_l29 = 0
                    print(f"\nTransformer {index + 1} - Final sum calculation for L29 (L17-L28):")
                    
                    # Get all values from L17 to L28
                    for row in range(16, 28):  # Rows 17-28 (0-based indices 16-27)
                        value = all_data[row][val_col] or '0'
                        try:
                            num_value = float(value)
                            print(f"Row {row + 1} value: {value} (converted to {num_value})")
                            sum_l29 += num_value
                        except ValueError:
                            print(f"Warning: Invalid number format in row {row + 1}: '{value}', using 0")
                            sum_l29 += 0
                    
                    # Store the final sum in L29
                    all_data[28][val_col] = str(round(sum_l29, 2))
                    print(f"Final sum for L29: {sum_l29}")
                    
                    # Format L16 and L29 cells to bold
                    # Convert column index to letter (L=12, Q=17, V=22, AA=27, etc.)
                    col_letter = chr(ord('A') + (val_col % 26))
                    if val_col >= 26:
                        col_letter = 'A' + col_letter  # For columns AA and beyond
                    
                    # Format L16
                    cell_range_16 = f"{col_letter}16"
                    print(f"Formatting cell {cell_range_16} to bold")
                    self.estimate_sheet.format(cell_range_16, {
                        "textFormat": {
                            "bold": True
                        }
                    })
                    
                    # Format L29
                    cell_range_29 = f"{col_letter}29"
                    print(f"Formatting cell {cell_range_29} to bold")
                    self.estimate_sheet.format(cell_range_29, {
                        "textFormat": {
                            "bold": True
                        }
                    })
                    
                except Exception as e:
                    print(f"Error calculating final sums for Transformer {index + 1}: {str(e)}")
                    all_data[15][val_col] = "0"  # Set L16 to 0
                    all_data[28][val_col] = "0"  # Set L29 to 0

            # Calculate the sum from L33 to L43 and store it in L44, Q44, V44, AA44
            print("\nCalculating sum for L44 (L33-L43) for each transformer...")
            for index, entries in enumerate(self.entry_sets):
                try:
                    # Calculate column positions for this transformer
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Initialize sum for L44
                    sum_l44 = 0
                    print(f"\nTransformer {index + 1} - Final sum calculation for L44 (L33-L43):")
                    
                    # Get all values from L33 to L43
                    for row in range(32, 43):  # Rows 33-43 (0-based indices 32-42)
                        value = all_data[row][val_col] if row < len(all_data) and val_col < len(all_data[row]) else '0'
                        try:
                            # Convert to float, if it fails, treat as 0
                            num_value = float(value) if value.replace('.', '', 1).isdigit() else 0
                            print(f"Row {row + 1} value: {value} (converted to {num_value})")
                            sum_l44 += num_value
                        except ValueError:
                            print(f"Warning: Invalid number format in row {row + 1}: '{value}', using 0")
                            sum_l44 += 0
                    
                    print(all_data, len(all_data), val_col )
                    # Store the final sum in the appropriate L44 cell for this transformer
                    if 43 < len(all_data) and val_col < len(all_data[43]):
                        all_data[43][val_col] = str(round(sum_l44, 2))  # Round to 2 decimal places
                    print(f"Final sum for L44: {sum_l44}")

                except Exception as e:
                    print(f"Error calculating final sum for Transformer {index + 1}: {str(e)}")
                    if 43 < len(all_data) and val_col < len(all_data[43]):
                        all_data[43][val_col] = "0"  # Set L44 to 0 in case of error

            # Calculate the sum of L29 and L44 and store it in L45, Q45, V45, AA45 for each transformer
            print("\nCalculating sum for L45, Q45, V45, AA45 (L29 + L44) for each transformer...")
            for index, entries in enumerate(self.entry_sets):
                try:
                    # Calculate column positions for this transformer
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Initialize sum for L45
                    sum_l45 = 0
                    print(f"\nTransformer {index + 1} - Final sum calculation for L45 (L29 + L44):")
                    
                    # Get values from L29 and L44
                    l29_value = all_data[28][val_col] if 28 < len(all_data) and val_col < len(all_data[28]) else '0'
                    l44_value = all_data[43][val_col] if 43 < len(all_data) and val_col < len(all_data[43]) else '0'
                    
                    # Convert values to float and calculate the sum
                    try:
                        l29_num = float(l29_value) if l29_value else 0
                        l44_num = float(l44_value) if l44_value else 0
                        sum_l45 = round(l29_num + l44_num, 2)
                        print(f"L29 value: {l29_value} (converted to {l29_num})")
                        print(f"L44 value: {l44_value} (converted to {l44_num})")
                        print(f"Sum for L45: {sum_l45}")
                    except ValueError:
                        print(f"Warning: Invalid number format in L29 or L44, using 0 for sum.")
                        sum_l45 = 0
                    
                    # Store the final sum in L45
                    if 44 < len(all_data) and val_col < len(all_data[44]):
                        all_data[44][val_col] = str(sum_l45)  # Store the sum in L45
                    else:
                        print(f"Warning: Unable to store sum in L45, index out of range.")

                except Exception as e:
                    print(f"Error calculating final sum for L45 for Transformer {index + 1}: {str(e)}")
                    if 44 < len(all_data) and val_col < len(all_data[44]):
                        all_data[44][val_col] = "0"  # Set L45 to 0 in case of error

            # Set Discount Value in L46, Q46, V46, AA46 for each transformer
            print("\nSetting Discount Value in L46, Q46, V46, AA46 to 0.00 for each transformer...")
            for index in range(len(self.entry_sets)):
                try:
                    # Calculate column positions for this transformer
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Set the Discount Value
                    discount_value = "0.00"
                    print(f"Setting Discount Value for Transformer {index + 1} in L46: {discount_value}")
                    
                    # Store the Discount Value in L46
                    if 45 < len(all_data) and val_col < len(all_data[45]):
                        all_data[45][val_col] = discount_value  # Store the discount value in L46
                    else:
                        print(f"Warning: Unable to store Discount Value in L46, index out of range.")

                except Exception as e:
                    print(f"Error setting Discount Value for Transformer {index + 1}: {str(e)}")
                    if 45 < len(all_data) and val_col < len(all_data[45]):
                        all_data[45][val_col] = "0.00"  # Set to "0.00" in case of error

            # Calculate the value for L47 (L45 + L16 - L46) for each transformer
            print("\nCalculating value for L47 (L45 + L16 - L46) for each transformer...")
            for index in range(len(self.entry_sets)):
                try:
                    # Calculate column positions for this transformer
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Initialize L47 value
                    l47_value = 0.00
                    print(f"\nTransformer {index + 1} - Calculating L47 value:")
                    
                    # Get values from L45, L16, and L46
                    l45_value = all_data[44][val_col] if 44 < len(all_data) and val_col < len(all_data[44]) else '0'
                    l16_value = all_data[15][val_col] if 15 < len(all_data) and val_col < len(all_data[15]) else '0'
                    l46_value = all_data[45][val_col] if 45 < len(all_data) and val_col < len(all_data[45]) else '0'
                    
                    # Convert values to float and calculate L47
                    try:
                        l45_num = float(l45_value) if l45_value else 0
                        l16_num = float(l16_value) if l16_value else 0
                        l46_num = float(l46_value) if l46_value else 0
                        
                        l47_value = round(l45_num + l16_num - l46_num, 2)
                        print(f"L45 value: {l45_value} (converted to {l45_num})")
                        print(f"L16 value: {l16_value} (converted to {l16_num})")
                        print(f"L46 value: {l46_value} (converted to {l46_num})")
                        print(f"Calculated L47 value: {l47_value}")
                    except ValueError:
                        print(f"Warning: Invalid number format in L45, L16, or L46, using 0 for L47 calculation.")
                        l47_value = 0.00
                    
                    # Store the final value in L47
                    if 46 < len(all_data) and val_col < len(all_data[46]):
                        all_data[46][val_col] = str(l47_value)  # Store the calculated value in L47
                    else:
                        print(f"Warning: Unable to store value in L47, index out of range.")

                except Exception as e:
                    print(f"Error calculating L47 for Transformer {index + 1}: {str(e)}")
                    if 46 < len(all_data) and val_col < len(all_data[46]):
                        all_data[46][val_col] = "0.00"  # Set L47 to "0.00" in case of error

            # Calculate the value for L48 (4% of L47) for each transformer
            print("\nCalculating value for L48 (4% of L47) for each transformer...")
            for index in range(len(self.entry_sets)):
                try:
                    # Calculate column positions for this transformer
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Initialize L48 value
                    l48_value = 0.00
                    print(f"\nTransformer {index + 1} - Calculating L48 value:")
                    
                    # Get value from L47
                    l47_value = all_data[46][val_col] if 46 < len(all_data) and val_col < len(all_data[46]) else '0'
                    
                    # Convert L47 value to float and calculate L48
                    try:
                        l47_num = float(l47_value) if l47_value else 0
                        l48_value = round(l47_num * 0.04, 2)  # Calculate 4% of L47
                        print(f"L47 value: {l47_value} (converted to {l47_num})")
                        print(f"Calculated L48 value (4% of L47): {l48_value}")
                    except ValueError:
                        print(f"Warning: Invalid number format in L47, using 0 for L48 calculation.")
                        l48_value = 0.00
                    
                    # Store the final value in L48
                    if 47 < len(all_data) and val_col < len(all_data[47]):
                        all_data[47][val_col] = str(l48_value)  # Store the calculated value in L48
                    else:
                        print(f"Warning: Unable to store value in L48, index out of range.")

                except Exception as e:
                    print(f"Error calculating L48 for Transformer {index + 1}: {str(e)}")
                    if 47 < len(all_data) and val_col < len(all_data[47]):
                        all_data[47][val_col] = "0.00"  # Set L48 to "0.00" in case of error

            # Calculate the value for L49 (L47 + L48) for each transformer
            print("\nCalculating value for L49 (L47 + L48) for each transformer...")
            for index in range(len(self.entry_sets)):
                try:
                    # Calculate column positions for this transformer
                    val_col = 2 + (index * 5)  # L=2, Q=7, V=12, AA=27, etc.
                    
                    # Initialize L49 value
                    l49_value = 0.00
                    print(f"\nTransformer {index + 1} - Calculating L49 value:")
                    
                    # Get values from L47 and L48
                    l47_value = all_data[46][val_col] if 46 < len(all_data) and val_col < len(all_data[46]) else '0'
                    l48_value = all_data[47][val_col] if 47 < len(all_data) and val_col < len(all_data[47]) else '0'
                    
                    # Convert values to float and calculate L49
                    try:
                        l47_num = float(l47_value) if l47_value else 0
                        l48_num = float(l48_value) if l48_value else 0
                        
                        l49_value = round(l47_num + l48_num, 2)  # Calculate L49 as L47 + L48
                        print(f"L47 value: {l47_value} (converted to {l47_num})")
                        print(f"L48 value: {l48_value} (converted to {l48_num})")
                        print(f"Calculated L49 value (L47 + L48): {l49_value}")
                    except ValueError:
                        print(f"Warning: Invalid number format in L47 or L48, using 0 for L49 calculation.")
                        l49_value = 0.00
                    
                    # Store the final value in L49
                    if 48 < len(all_data) and val_col < len(all_data[48]):
                        all_data[48][val_col] = str(l49_value)  # Store the calculated value in L49
                    else:
                        print(f"Warning: Unable to store value in L49, index out of range.")

                except Exception as e:
                    print(f"Error calculating L49 for Transformer {index + 1}: {str(e)}")
                    if 48 < len(all_data) and val_col < len(all_data[48]):
                        all_data[48][val_col] = "0.00"  # Set L49 to "0.00" in case of error

            # Update the sheet with all calculations
            self.estimate_sheet.update('J1:FR50', all_data)
            messagebox.showinfo("Success", "Data saved to ESTIMATE sheet successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data to ESTIMATE sheet: {str(e)}")
            print(f"Debug - Error details: {str(e)}") 