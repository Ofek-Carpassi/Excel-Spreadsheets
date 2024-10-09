import tkinter as tk

class Cell:
    def __init__(self, master, text="", width=10, height=2, borderwidth=1, relief="solid", bg="white", fg="black"):
        self.master = master
        self.label = tk.Label(master, text=text, borderwidth=borderwidth, relief=relief, width=width, height=height, bg=bg, fg=fg)
        self.label.bind("<Double-Button-1>", self.edit_cell)
        self.text_var = tk.StringVar()
        self.entry = tk.Entry(master, textvariable=self.text_var, width=width, borderwidth=borderwidth, relief=relief, bg=bg, fg=fg)
        self.entry.bind("<Return>", self.save_edit)

    def grid(self, row, column, sticky='nsew'):
        self.label.grid(row=row, column=column, sticky=sticky)
        self.row = row
        self.column = column

    def set_text(self, text):
        self.label.config(text=text)

    def get_text(self):
        return self.label.cget("text")

    def set_bg(self, color):
        self.label.config(bg=color)
        self.entry.config(bg=color)

    def set_fg(self, color):
        self.label.config(fg=color)
        self.entry.config(fg=color)

    def edit_cell(self, event):
        self.text_var.set(self.get_text())
        self.entry.grid(row=self.row, column=self.column, sticky='nsew')
        self.entry.focus_set()

    def save_edit(self, event):
        text = self.text_var.get()
        if text.startswith("bg="):
            color = text.split("=")[1]
            self.set_bg(color)
        elif text.startswith("txtcolor="):
            color = text.split("=")[1]
            self.set_fg(color)
        else:
            self.set_text(text)
            # Adjust the size of the row/column if the text is too long
            text_length = len(text)
            if text_length > self.label.cget("width"):
                self.label.config(width=text_length)
                self.entry.config(width=text_length)
                self.master.grid_columnconfigure(self.column, minsize=text_length * 10)  # Adjust column width
                self.master.grid_rowconfigure(self.row, minsize=text_length * 2)  # Adjust row height
        self.entry.grid_forget()

class GUI:
    def __init__(self, windowSize=(800, 600)):
        self.windowSize = windowSize

    # Function to create the main window
    def create_window(self):
        self.window = tk.Tk()
        self.window.title("Excel Spreadsheet")
        self.window.geometry(f"{self.windowSize[0]}x{self.windowSize[1]}")

    # Function to create a spreadsheet - a grid of cells
    def create_spreadsheet(self, cellSize=(100, 30), rows=30, columns=30):
        # Create a frame to hold the canvas and scrollbars
        frame = tk.Frame(self.window)
        frame.pack(fill=tk.BOTH, expand=True)

        # Create vertical scrollbar on the left
        if rows > 30:
            v_scroll = tk.Scrollbar(frame, orient=tk.VERTICAL)
            v_scroll.pack(side=tk.LEFT, fill=tk.Y)

        # Create horizontal scrollbar at the bottom
        h_scroll = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # Create the canvas
        self.spreadsheet = tk.Canvas(frame, yscrollcommand=v_scroll.set if rows > 30 else None, xscrollcommand=h_scroll.set)
        self.spreadsheet.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure the scrollbars to control the canvas
        if rows > 30:
            v_scroll.config(command=self.spreadsheet.yview)
        h_scroll.config(command=self.spreadsheet.xview)

        # Create an inner frame to hold the grid of cells
        inner_frame = tk.Frame(self.spreadsheet)
        self.spreadsheet.create_window((0, 0), window=inner_frame, anchor='nw')

        # Create the grid of cells
        self.cells = []
        for i in range(rows):
            row_cells = []
            for j in range(columns):
                if i == 0 and j > 0:
                    # First row, add column letters
                    cell = Cell(inner_frame, text=chr(64+j), width=cellSize[0]//10, height=cellSize[1]//20, bg="white", fg="black")
                elif j == 0 and i > 0:
                    # First column, add row numbers
                    cell = Cell(inner_frame, text=str(i), width=cellSize[0]//10, height=cellSize[1]//20, bg="white", fg="black")
                else:
                    # Other cells
                    cell = Cell(inner_frame, width=cellSize[0]//10, height=cellSize[1]//20)
                cell.grid(row=i, column=j, sticky='nsew')
                row_cells.append(cell)
            self.cells.append(row_cells)

        # Configure the scroll region
        inner_frame.update_idletasks()
        self.spreadsheet.config(scrollregion=self.spreadsheet.bbox(tk.ALL))

    def create_menu(self):
        pass

    def run(self):
        self.window.mainloop()