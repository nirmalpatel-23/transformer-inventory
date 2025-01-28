import gspread
from tkinter import messagebox
from oauth2client.service_account import ServiceAccountCredentials

# Define sheet names and their configurations
SPREADSHEET_NAME = "MASTER SHEET"
SHEET_CONFIGS = {
    "MASTER": {
        "name": "MASTER SHEET",
        "forms": ["add_form", "second_form", "physical_form"]
    },
    
}

def get_worksheet(sheet_name):
    """Get a specific worksheet"""
    try:
        spreadsheet = setup_google_sheets()
        if spreadsheet:
            return spreadsheet.worksheet(sheet_name)
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Failed to get worksheet: {e}")
        return None

def setup_google_sheets():
    """
    Connect to Google Sheets and return the spreadsheet
    Returns:
        gspread.Spreadsheet: The main spreadsheet object
    """
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        client = gspread.authorize(credentials)
        return client.open(SPREADSHEET_NAME)
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to connect to Google Sheets: {e}")
        return None

def save_to_sheets(data, form_name, additional_data=None):
    """
    Save data to appropriate sheets based on form type
    Args:
        data: List of data to save
        form_name: Name of the form submitting data
        additional_data: Any additional data needed for saving
    """
    try:
        # Determine which sheets need this data
        for config in SHEET_CONFIGS.values():
            if form_name in config["forms"]:
                worksheet = get_worksheet(config["name"])
                if worksheet:
                    # If it's second_form data, we need to update existing rows
                    if form_name == "second_form" and additional_data:
                        start_row = additional_data["start_row"]
                        worksheet.update(f"F{start_row}:I{start_row + len(data) - 1}", data)
                    else:
                        worksheet.append_rows(data)

        return True
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save data: {e}")
        return False 