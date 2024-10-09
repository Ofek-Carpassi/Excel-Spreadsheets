import GUI as gui

def main():
    gui_obj = gui.GUI()
    gui_obj.create_window()
    gui_obj.create_spreadsheet()
    gui_obj.run()

if __name__ == "__main__":
    main()