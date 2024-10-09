import GUI as gui

def main():
    gui_obj = gui.GUI()
    gui_obj.create_window()
    gui_obj.create_spreadsheet()

    gui_obj.cells[2][3].set_text("Hello")
    print(gui_obj.cells[2][3].get_text())

    gui_obj.cells[2][3].set_bg("red")
    gui_obj.cells[2][3].set_fg("white")

    gui_obj.run()

if __name__ == "__main__":
    main()