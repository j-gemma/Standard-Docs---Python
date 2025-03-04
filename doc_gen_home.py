from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import date
from math import floor
import item_spec_backend as spec_backend
import item_spec_frontend as spec_frontend
import cover_frontend

# Window sizes for different sections
HOME_SIZE = (500, 200)  # Main home window size (width x height)
COVER_WINDOW_SIZE = (675, 550)  # Cutbook Covers window size
ITEM_WINDOW_SIZE = (650, 250)  # Item Specification window size

class Home(Tk):
    """
    Main application window class inheriting from Tk.
    It serves as the entry point for the user interface.
    """
    def __init__(self, title, size):
        super().__init__()
        self.title(title)  # Set the window title
        self.geometry(f'{size[0]}x{size[1]}')  # Set the initial size of the window
        self.setup_menu()  # Initialize the menu section
        self.protocol("WM_DELETE_WINDOW", self.quit_app)  # Handle window close events

    def show_window(self):
        """Bring the window to the front and make it visible."""
        self.deiconify()  # Unhide the window if it was minimized or hidden
        self.lift()  # Ensure the window is on top of other windows

    def setup_menu(self):
        """Set up the main menu displayed in the home window."""
        self.menu = HomeMenu(self)  # Create an instance of the HomeMenu class
        self.menu.grid(row=0, column=0)  # Place it in the first row and column

    def quit_app(self):
        """Quit the application gracefully by destroying the window."""
        self.quit()
        self.destroy()


class HomeMenu(ttk.Frame):
    """
    A frame to hold the main menu of the application.
    Contains buttons to access specific functionalities like Item Spec and Cutbook Covers.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent  # Reference to the parent (main window)
        self.create_widgets()  # Add the menu buttons and labels


    """
        Create and layout the menu buttons and descriptions.
    """
    # Define menu buttons: (Button text, Command function, Description)
    def create_widgets(self):
            
        buttons = [
            ("Item Spec", lambda: self.item_spec(), 'Generate itemized specification from AQ Excel export'),
            ("Cutbook Covers", lambda: self.cut_covers(), 'Batch generate cutbook covers for a single project'),
            ("General Spec", None, 'Select std. details to include in general spec - *NOT YET IMPLEMENTED*'),
            ("Cutbook", None, 'Generate cutbook from AQ Excel - *NOT YET IMPLEMENTED*')
        ]

        # Iterate over each button, create it, and layout along with its description
        for i, (text, command, desc) in enumerate(buttons):
            button = ttk.Button(self, text=text, command=command)  # Create a button
            button.grid(row=i, column=0, sticky='we')  # Place it in the menu grid
            label = ttk.Label(self, text=f"{desc}")  # Create a descriptive label
            label.grid(row=i, column=1, sticky='we')  # Place the label next to the button

        # Add padding around buttons and labels
        pad_config(self, 5)
    
    """
        Open the Item Specification window and hide the main menu.
    """
    def item_spec(self):
        self.parent.withdraw()  # Hide the home window
        win = spec_frontend.ItemSpec('Item Specification', ITEM_WINDOW_SIZE)  # Open the new window
        win.lift()

        # Define a function to handle the window close event
        def close_and_show_home():
            win.destroy()  # Close the current window
            self.parent.show_window()  # Show the home window again
       
        win.protocol("WM_DELETE_WINDOW", close_and_show_home)  # Link the function to the close button
    
    """
        Open the Cutbook Covers window and hide the main menu.
    """
    def cut_covers(self):
    
        self.parent.withdraw()  # Hide the home window
        win = cover_frontend.CutCover('Cutbook Covers', COVER_WINDOW_SIZE)  # Open the new window
        win.lift()

        # Define a function to handle the window close event
        def close_and_show_home():
            win.destroy()  # Close the current window
            self.parent.show_window()  # Show the home window again
        
        win.protocol("WM_DELETE_WINDOW", close_and_show_home)  # Link the function to the close button


"""
    Add consistent padding around all child widgets in a frame.
    :param widget: The parent widget whose children need padding.
    :param pad_amt: The padding amount (pixels).
"""
def pad_config(widget, pad_amt):
    
    for child in widget.winfo_children():
        child.grid_configure(padx=pad_amt, pady=pad_amt)  # Apply padding to each child

# Entry point of the program: Create the home window and start the Tkinter main event loop
if __name__ == "__main__": 
    app = Home('Document Generator', HOME_SIZE)  # Create the main application window
    app.lift()  # Bring it to the front
    app.mainloop()  # Start the event loop
