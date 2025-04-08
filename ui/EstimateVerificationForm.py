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

            # Update the sheet in one batch
            self.estimate_sheet.update('J1:FR43', all_data)
            messagebox.showinfo("Success", "Data saved to ESTIMATE sheet successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data to ESTIMATE sheet: {str(e)}")
            print(f"Debug - Error details: {str(e)}") 