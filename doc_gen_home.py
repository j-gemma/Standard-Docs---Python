from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import date
from math import floor
import item_spec_backend as spec_backend
import item_spec_frontend as spec_frontend
import cover_frontend

HOME_SIZE = (500, 200)
COVER_WINDOW_SIZE = (675, 550)
ITEM_WINDOW_SIZE = (650, 250)

class Home(Tk):
    def __init__(self, title, size):
        super().__init__()
        self.title(title)
        self.geometry(f'{size[0]}x{size[1]}')
        self.setup_menu()
        self.protocol("WM_DELETE_WINDOW", self.quit_app)  # Handle close button

    def show_window(self):
        self.deiconify()  # Unhides the window
        self.lift()       # Brings the window to the front
        
    def setup_menu(self):
        self.menu = HomeMenu(self)
        self.menu.grid(row=0, column=0)

    def quit_app(self):
        # Quit the application
        self.quit()
        self.destroy()

class HomeMenu(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        buttons = [
            ("Item Spec", lambda: self.item_spec(), 'Generate itemized specification from AQ Excel export'),
            ("Cutbook Covers", lambda: self.cut_covers(), 'Batch generate cutbook covers for a single project'),
            ("General Spec", None, 'Select std. details to include in general spec - *NOT YET IMPLEMENTED*'),
            ("Cutbook", None, 'Generate cutbook from AQ Excel - *NOT YET IMPLEMENTED*')
        ]

        for i, (text, command, desc) in enumerate(buttons):
            button = ttk.Button(self, text=text, command=command)
            button.grid(row=i, column=0, sticky='we')
            label = ttk.Label(self, text=f"{desc}")
            label.grid(row=i, column=1, sticky='we')

        pad_config(self, 5)

    def item_spec(self):
        self.parent.withdraw()
        win = spec_frontend.ItemSpec('Item Specification', ITEM_WINDOW_SIZE)
        win.lift()
        def close_and_show_home():
            win.destroy()  # Close the current window
            self.parent.show_window()  # Show the home window
        win.protocol("WM_DELETE_WINDOW", close_and_show_home)
        
    def cut_covers(self):
        self.parent.withdraw()
        win = cover_frontend.CutCover('Cutbook Covers', COVER_WINDOW_SIZE)
        win.lift()
        def close_and_show_home():
            win.destroy()  # Close the current window
            self.parent.show_window()  # Show the home window
        win.protocol("WM_DELETE_WINDOW", close_and_show_home)

def pad_config(widget, pad_amt):
    for child in widget.winfo_children():
        child.grid_configure(padx=pad_amt, pady=pad_amt)

if __name__ == "__main__":
    app = Home('Document Generator', HOME_SIZE)
    app.lift()
    app.mainloop()
