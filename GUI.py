import tkinter as tk
from file_manager import FileManager
from tkinter import messagebox

class Cell(tk.Entry):
    def __init__(self, parent, text="", width=6, height=2, **kwargs):
        super().__init__(parent, width=width, **kwargs)
        self.insert(0, text)
        self.config(relief="solid", bd=1)  # Bolder outline

    def get_text(self):
        return self.get()

    def set_text(self, text):
        self.delete(0, tk.END)
        self.insert(0, text)

class HeaderCell(tk.Label):
    def __init__(self, parent, text="", width=6, height=1, **kwargs):
        super().__init__(parent, text=text, width=width, height=height, **kwargs)
        self.config(font=("Arial", 12, "bold"), anchor="center", relief="solid", bd=1)  # Bold and centered

    def get_text(self):
        return self.cget("text")

    def set_text(self, text):
        self.config(text=text)

class GUI:
    def __init__(self, windowSize=(800, 600)):
        self.windowSize = windowSize
        self.file_manager = FileManager(self)

    def create_window(self):
        self.window = tk.Tk()
        self.window.title("Excel Spreadsheet")
        self.window.geometry(f"{self.windowSize[0]}x{self.windowSize[1]}")
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)  # Intercept the window close event

    def create_spreadsheet(self, cellSize=(90, 30), rows=25, columns=9):
        frame = tk.Frame(self.window)
        frame.pack(fill=tk.BOTH, expand=True)

        if rows > 30:
            v_scroll = tk.Scrollbar(frame, orient=tk.VERTICAL)
            v_scroll.pack(side=tk.LEFT, fill=tk.Y)

        h_scroll = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.spreadsheet = tk.Canvas(frame, yscrollcommand=v_scroll.set if rows > 30 else None, xscrollcommand=h_scroll.set)
        self.spreadsheet.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if rows > 30:
            v_scroll.config(command=self.spreadsheet.yview)
        h_scroll.config(command=self.spreadsheet.xview)

        inner_frame = tk.Frame(self.spreadsheet)
        self.spreadsheet.create_window((0, 0), window=inner_frame, anchor='nw')

        self.cells = []
        for i in range(rows):
            row_cells = []
            for j in range(columns):
                if i == 0 and j > 0:
                    cell = HeaderCell(inner_frame, text=chr(64+j), width=cellSize[0]//10, height=cellSize[1]//20, bg="white", fg="black")
                elif j == 0 and i > 0:
                    cell = HeaderCell(inner_frame, text=str(i), width=cellSize[0]//10, height=cellSize[1]//20, bg="white", fg="black")
                else:
                    cell = Cell(inner_frame, width=cellSize[0]//10, height=cellSize[1]//20)
                cell.grid(row=i, column=j, sticky='nsew')
                row_cells.append(cell)
            self.cells.append(row_cells)

        inner_frame.update_idletasks()
        self.spreadsheet.config(scrollregion=self.spreadsheet.bbox(tk.ALL))

    def create_menu(self):
        menu_bar = tk.Menu(self.window)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Save as", command=self.file_manager.save_as)
        file_menu.add_command(label="Save", command=self.file_manager.save)
        file_menu.add_command(label="Open", command=self.file_manager.open_file)
        file_menu.add_command(label="Create new file", command=self.file_manager.create_new_file)
        file_menu.add_command(label="Export to CSV", command=self.file_manager.export_to_csv)  # New menu option
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.window.config(menu=menu_bar)

    def get_spreadsheet_data(self):
        data = []
        for row in self.cells:
            row_data = [cell.get_text() for cell in row]
            data.append(row_data)
        return data

    def load_spreadsheet_data(self, data):
        for i, row in enumerate(data):
            for j, text in enumerate(row):
                self.cells[i][j].set_text(text)

    def clear_spreadsheet(self, keep_headers=False):
        for i, row in enumerate(self.cells):
            for j, cell in enumerate(row):
                if keep_headers and (i == 0 or j == 0):
                    continue
                cell.set_text("")

    def on_closing(self):
        if messagebox.askyesno("Exit", "Do you want to save the current file before exiting?"):
            self.file_manager.save()
        self.window.destroy()

    def run(self):
        self.window.mainloop()

def main():
    gui = GUI()
    gui.create_window()
    gui.create_menu()
    gui.create_spreadsheet()
    gui.run()