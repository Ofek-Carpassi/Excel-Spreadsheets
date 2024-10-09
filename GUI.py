import tkinter as tk

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
        for i in range(rows):
            for j in range(columns):
                if i == 0 and j > 0:
                    # First row, add column letters
                    cell = tk.Label(inner_frame, text=chr(64+j), borderwidth=1, relief="solid", width=cellSize[0]//10, height=cellSize[1]//20, bg="white", fg="black")
                elif j == 0 and i > 0:
                    # First column, add row numbers
                    cell = tk.Label(inner_frame, text=str(i), borderwidth=1, relief="solid", width=cellSize[0]//10, height=cellSize[1]//20, bg="white", fg="black")
                else:
                    # Other cells
                    cell = tk.Label(inner_frame, borderwidth=1, relief="solid", width=cellSize[0]//10, height=cellSize[1]//20)
                cell.grid(row=i, column=j, sticky='nsew')

        # Configure the scroll region
        inner_frame.update_idletasks()
        self.spreadsheet.config(scrollregion=self.spreadsheet.bbox(tk.ALL))

    def create_menu(self):
        pass

    def run(self):
        self.window.mainloop()