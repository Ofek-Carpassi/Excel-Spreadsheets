import json
from tkinter import filedialog, messagebox

class FileManager:
    def __init__(self, gui):
        self.gui = gui
        self.current_file = None

    def save_as(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if file_path:
            self.current_file = file_path
            self.save()

    def save(self):
        if not self.current_file:
            self.save_as()
        else:
            data = self.gui.get_spreadsheet_data()
            with open(self.current_file, 'w') as file:
                json.dump(data, file)

    def open_file(self):
        if messagebox.askyesno("Open File", "Do you want to save the current file before opening a new one?"):
            self.save()
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if file_path:
            self.current_file = file_path
            with open(file_path, 'r') as file:
                data = json.load(file)
                self.gui.clear_spreadsheet(keep_headers=True)
                self.gui.load_spreadsheet_data(data)

    def create_new_file(self):
        if messagebox.askyesno("Create New File", "Do you want to save the current file before creating a new one?"):
            self.save()
        self.current_file = None
        self.gui.clear_spreadsheet(keep_headers=True)

    def export_to_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            data = self.gui.get_spreadsheet_data()
            with open(file_path, 'w') as file:
                for row in data:
                    file.write(",".join(row) + "\n")