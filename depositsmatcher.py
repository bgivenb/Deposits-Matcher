import tkinter as tk
from tkinter import messagebox
from itertools import combinations
from PIL import Image, ImageTk
import sys
import os

# Program: DepositsMatcher
# Author: Given Borthwick

class DepositsMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DepositsMatcher - by Given Borthwick")
        
        # Title with icon at the top
        title_frame = tk.Frame(root)
        title_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        # Text "Deposits" Label
        title_label_left = tk.Label(title_frame, text="Deposits", font=("Arial", 16, "bold"))
        title_label_left.pack(side=tk.LEFT)

        # Determine the path for the icon image
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, "deposits.png")
        else:
            icon_path = "deposits.png"

        # Load and display icon in the title
        try:
            icon_image = Image.open(icon_path)  # Load the image
            icon_image = icon_image.resize((24, 24), Image.LANCZOS)  # Resize with high-quality resampling
            icon_photo = ImageTk.PhotoImage(icon_image)  # Convert to PhotoImage for Tkinter
            icon_label = tk.Label(title_frame, image=icon_photo)
            icon_label.image = icon_photo  # Keep a reference to avoid garbage collection
            icon_label.pack(side=tk.LEFT, padx=5)
        except Exception as e:
            print("Error loading icon:", e)

        # Text "Matcher" Label
        title_label_right = tk.Label(title_frame, text="Matcher", font=("Arial", 16, "bold"))
        title_label_right.pack(side=tk.LEFT)

        # Help button
        help_button = tk.Button(title_frame, text="Help", command=self.show_help)
        help_button.pack(side=tk.RIGHT, padx=10)

        # Input for number of deposits for List A
        tk.Label(root, text="Number of Deposits for List A:").grid(row=1, column=0, padx=10, pady=5)
        self.num_deposits_a_entry = tk.Entry(root)
        self.num_deposits_a_entry.grid(row=1, column=1, padx=10, pady=5)

        # Input for number of deposits for List B
        tk.Label(root, text="Number of Deposits for List B:").grid(row=2, column=0, padx=10, pady=5)
        self.num_deposits_b_entry = tk.Entry(root)
        self.num_deposits_b_entry.grid(row=2, column=1, padx=10, pady=5)

        # Button to generate deposit fields for List A and List B
        self.generate_button = tk.Button(root, text="Generate Deposit Fields", command=self.generate_fields)
        self.generate_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

        # Placeholders for deposit fields and verify button
        self.list_a_entries = []
        self.list_b_entries = []
        self.verify_button = None

        # Footer at the bottom
        footer_label = tk.Label(root, text="Created by Given Borthwick", font=("Arial", 10, "italic"))
        footer_label.grid(row=1000, column=0, columnspan=2, padx=10, pady=20)  # Position at the bottom with a high row number

    def show_help(self):
        # Display a help message
        help_text = (
            "DepositsMatcher is a tool for accounting and bookkeeping to match "
            "deposit amounts between two lists (List A and List B). "
            "\n\nFeatures:\n"
            "- Input deposits for List A and List B.\n"
            "- Find matching subset totals between the two lists.\n"
            "- Highlight discrepancies if there are unmatched amounts.\n\n"
            "This tool helps users compare deposit entries to ensure accuracy."
        )
        messagebox.showinfo("Help - DepositsMatcher", help_text)

    def generate_fields(self):
        # Clear previous fields
        for entry in self.list_a_entries + self.list_b_entries:
            entry.destroy()
        self.list_a_entries.clear()
        self.list_b_entries.clear()
        
        if self.verify_button:
            self.verify_button.destroy()
        
        # Get the number of deposit fields to create for List A
        try:
            num_deposits_a = int(self.num_deposits_a_entry.get())
            if num_deposits_a <= 0:
                raise ValueError("Number of deposits for List A must be positive.")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid positive integer for the number of deposits in List A.")
            return

        # Get the number of deposit fields to create for List B
        try:
            num_deposits_b = int(self.num_deposits_b_entry.get())
            if num_deposits_b <= 0:
                raise ValueError("Number of deposits for List B must be positive.")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid positive integer for the number of deposits in List B.")
            return
        
        # Create deposit fields for List A
        tk.Label(self.root, text="List A Deposits:").grid(row=4, column=0, columnspan=2, padx=10, pady=5)
        for i in range(num_deposits_a):
            label = tk.Label(self.root, text=f"A{i+1}:")
            label.grid(row=5+i, column=0, padx=10, pady=5)
            entry = tk.Entry(self.root)
            entry.grid(row=5+i, column=1, padx=10, pady=5)
            self.list_a_entries.append(entry)

        # Create deposit fields for List B
        tk.Label(self.root, text="List B Deposits:").grid(row=5+num_deposits_a, column=0, columnspan=2, padx=10, pady=5)
        for i in range(num_deposits_b):
            label = tk.Label(self.root, text=f"B{i+1}:")
            label.grid(row=6+num_deposits_a+i, column=0, padx=10, pady=5)
            entry = tk.Entry(self.root)
            entry.grid(row=6+num_deposits_a+i, column=1, padx=10, pady=5)
            self.list_b_entries.append(entry)
        
        # Create a button to verify the sums
        self.verify_button = tk.Button(self.root, text="Find Maximum Matching Sum", command=self.find_max_matching_sum)
        self.verify_button.grid(row=6+num_deposits_a+num_deposits_b, column=0, columnspan=2, padx=10, pady=10)

    def find_max_matching_sum(self):
        try:
            deposits_a = [float(entry.get()) for entry in self.list_a_entries]
            deposits_b = [float(entry.get()) for entry in self.list_b_entries]
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for all deposits.")
            return

        # Function to get all subset sums of a list
        def get_subset_sums(lst):
            subset_sums = set()
            for r in range(1, len(lst) + 1):
                for subset in combinations(lst, r):
                    subset_sums.add(sum(subset))
            return subset_sums

        # Get all subset sums for lists A and B
        subset_sums_a = get_subset_sums(deposits_a)
        subset_sums_b = get_subset_sums(deposits_b)

        # Find the maximum matching sum between subset sums of A and B
        matching_sums = subset_sums_a.intersection(subset_sums_b)
        if matching_sums:
            max_matching_sum = max(matching_sums)
            remaining_a = sum(deposits_a) - max_matching_sum
            remaining_b = sum(deposits_b) - max_matching_sum
            result_message = f"Max matching sum: {max_matching_sum}\nRemaining sum in List A: {remaining_a}\nRemaining sum in List B: {remaining_b}"
        else:
            result_message = "No matching subset sums found."

        # Display the result in a popup window
        messagebox.showinfo("Result", result_message)

# Main loop
root = tk.Tk()
app = DepositsMatcherApp(root)
root.mainloop()
