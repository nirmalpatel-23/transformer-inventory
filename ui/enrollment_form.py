import tkinter as tk
from tkinter import ttk, messagebox

class EnrollmentForm:
    def __init__(self, row_data):
        self.window = tk.Toplevel()
        self.window.title("Enrollment Form")
        self.window.geometry("800x400")
        
        # Create a frame for the enrollment details
        self.frame = ttk.Frame(self.window, padding=20)
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Add title
        title_label = ttk.Label(self.frame, text="Enrollment Details", font=("Arial", 16, "bold"))
        title_label.grid(row=0, columnspan=2, pady=(0, 20))

        # Define the detail labels
        detail_labels = [
            "DIVISION", "TRUCK NO", "MR NO", "LOT NO", 
            "DATE", "TC NO", "MAKE", "TC CAPACITY", "JOB NO"
        ]

        # Create labels and entries for each detail
        self.detail_entries = {}
        for i, label in enumerate(detail_labels):
            ttk.Label(self.frame, text=f"{label}:").grid(row=i + 1, column=0, sticky="w", padx=5, pady=5)
            entry = ttk.Entry(self.frame, width=30)
            entry.grid(row=i + 1, column=1, padx=5, pady=5)
            entry.insert(0, row_data[i] if i < len(row_data) else "")  # Populate with data
            entry.configure(state='readonly')  # Make it read-only
            self.detail_entries[label] = entry

        # Add Enroll and Enrolled buttons
        self.enroll_button = ttk.Button(self.frame, text="Enroll", command=self.enroll)
        self.enroll_button.grid(row=len(detail_labels) + 1, column=0, padx=5, pady=20)

        self.enrolled_button = ttk.Button(self.frame, text="Enrolled", command=self.mark_as_enrolled)
        self.enrolled_button.grid(row=len(detail_labels) + 1, column=1, padx=5, pady=20)

    def enroll(self):
        # Logic for enrollment
        messagebox.showinfo("Enrollment", "Enrollment successful!")
        self.window.destroy()  # Close the window after enrollment

    def mark_as_enrolled(self):
        # Logic for marking as enrolled
        messagebox.showinfo("Enrollment", "Marked as enrolled!")
        self.window.destroy()  # Close the window after marking as enrolled 