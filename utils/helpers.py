import re
import tkinter as tk
from tkinter import messagebox

def get_last_lot_number(sheet, division):
    try:
        all_values = sheet.get_all_values()
        
        if all_values:
            all_values = all_values[1:]
        
        division_rows = [row for row in all_values if row[0] == division]
        
        if not division_rows:
            return f"{division[0]}1"
        
        lot_numbers = [row[3] for row in division_rows if row[3]]
        
        if not lot_numbers:
            return f"{division[0]}1"
        
        max_number = 0
        prefix = division[0].upper()
        
        for lot in lot_numbers:
            match = re.match(f"{prefix}(\d+)", lot)
            if match:
                number = int(match.group(1))
                max_number = max(max_number, number)
        
        return f"{prefix}{max_number + 1}"
    
    except Exception as e:
        messagebox.showerror("Error", f"Failed to get last lot number: {e}")
        return f"{division[0]}1"

def on_division_select(event, division_var, lot_entry, sheet):
    selected_division = division_var.get()
    if selected_division:
        next_lot_number = get_last_lot_number(sheet, selected_division)
        lot_entry.configure(state='normal')
        lot_entry.delete(0, tk.END)
        lot_entry.insert(0, next_lot_number)
        lot_entry.configure(state='readonly') 