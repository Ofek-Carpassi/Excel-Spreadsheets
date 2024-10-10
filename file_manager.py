import json # Used to save and load data in JSON format
from tkinter import filedialog, messagebox # Used to open file dialog and show message boxes
import csv # Used to export data to CSV format

# FileHandling class
class FileManager:
    # Initialize the object with the GUI object and the current file
    def __init__(self, gui):
        self.gui = gui
        self.current_file = None

    # Save the current file as a new file in JSON format
    def save_as(self):
        # Get the file path from the user
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        # If the user selected a file path
        if file_path:
            self.current_file = file_path # Set the current file to the selected file path
            self.save() # Call the save method
    
    # Save the current file in JSON format
    def save(self):
        if not self.current_file:
            self.save_as() # Call the save_as method
        else:
            # Get the data from the spreadsheet 
            data = self.gui.get_spreadsheet_data_with_styles()
            # Save the data in the current file
            with open(self.current_file, 'w') as file:
                # Write the data to the file
                json.dump(data, file)

    # Open a file from the json format
    def open_file(self):
        # If the user wants to save the current file before opening a new one
        if messagebox.askyesno("Open File", "Do you want to save the current file before opening a new one?"):
            self.save()
        # Get the file path from the user
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        # If the user selected a file path
        if file_path:
            # Set the current file to the selected file path
            self.current_file = file_path
            # Open the file
            with open(file_path, 'r') as file:
                # Read the data from the file
                data = json.load(file)
                # Clear the spreadsheet
                self.gui.clear_spreadsheet(keep_headers=True)
                # Load the data into the spreadsheet
                self.gui.load_spreadsheet_data_with_styles(data)

    # Create a new file
    def create_new_file(self):
        # If the user wants to save the current file before creating a new one
        if messagebox.askyesno("Create New File", "Do you want to save the current file before creating a new one?"):
            self.save()
        # Clear the current file
        self.current_file = None
        # Clear the spreadsheet
        self.gui.clear_spreadsheet(keep_headers=True)

    # Export the data to a CSV file
    def export_to_csv(self):
        # Get the file path from the user
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            # Get the data from the spreadsheet
            data = self.gui.get_spreadsheet_data()
            # Save the data in the CSV file
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)
                # Write the data to the file
                for row in data:
                    writer.writerow(row)