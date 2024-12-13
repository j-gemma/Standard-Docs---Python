from tkinter import *
from tkinter import ttk, messagebox, filedialog
import item_spec_backend as spec_backend
from datetime import date

class ItemSpec(Toplevel):
    def __init__(self, title, size):
        super().__init__()
        self.setup_window(title, size)
        self.add_menu()

    def add_menu(self):
        menu = ItemOptions(self)
        menu.grid(column=0, row=0)

    def setup_window(self, title, size):
        self.title(title)
        x = (self.winfo_screenwidth() - size[0]) // 2
        y = (self.winfo_screenheight() - size[1]) // 2
        self.geometry(f'{size[0]}x{size[1]}+{x}+{y}')

class ItemOptions(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.grid(column=0, row=0, sticky='nsew')
        self.create_widgets()

    def create_widgets(self):
        today = date.today().strftime('%B %d, %Y')
        today_formatted = date.today().strftime('%y%m%d')

        you_r_here = ttk.Label(self, text='Today is ' + today + ' [' + today_formatted + ']')
        you_r_here.grid(row=0, column=0, columnspan=2, sticky='ew')

        ttk.Label(self, text='Excel file exported from AutoQuotes:').grid(column=0, row=1, sticky='e')
        ttk.Label(self, text='File to be saved at:').grid(column=0, row=3, sticky='e')
        ttk.Label(self, text='Save as:').grid(column=0, row=5, sticky='e')

        ttk.Button(self, text='Pick Excel File', command=self.xl_pick).grid(column=0, row=2, sticky='e')
        ttk.Button(self, text='Pick Directory', command=self.dir_pick).grid(column=0, row=4, sticky='e')

        ttk.Button(self, text="Generate", command=self.make_spec).grid(row=9, column=0, columnspan=2)

        self.xl_path = StringVar()
        self.xl_path.trace_add('write', self.xl_path_callback)
        ttk.Label(self, textvariable=self.xl_path, width=55, relief='sunken').grid(column=1, row=1, sticky='ew')

        self.dir_path = StringVar()
        self.dir_path.trace_add('write', self.dir_path_callback)
        ttk.Label(self, textvariable=self.dir_path, width=55, relief='sunken').grid(column=1, row=3, sticky='ew')

        self.filename = StringVar()
        self.filename.trace_add('write', self.filename_callback)
        self.filename.set(today_formatted + ' Foodservice Section 114000')
        ttk.Entry(self, textvariable=self.filename).grid(column=1, row=5, sticky='ew')

        self.to_compose = StringVar()
        self.to_compose.set('compose')
        self.to_compose.trace_add('write', self.checkbox_callback)
        chk_items_only = ttk.Checkbutton(self, text='ITEMS ONLY', onvalue='nocompose', offvalue='compose', variable=self.to_compose)
        chk_items_only.grid(row=8, column=0, sticky=E)
        pad_config(self, 5)

    def dir_pick(self):
        self.dir_path.set(filedialog.askdirectory(title='File to be saved at:'))

    def xl_pick(self):
        self.xl_path.set(filedialog.askopenfilename(title='Excel file from AQ', filetypes=[('Excel Files', '*.xlsx')]))

    def xl_path_callback(self, *args):
        print(self.xl_path.get())

    def checkbox_callback(self, *args):
        new_value = self.to_compose.get()
        print(f"Checkbox new value: {new_value}")
        # You can add any logic here to respond to the change.
    
    def dir_path_callback(self, *args):
        print(self.dir_path.get())

    def filename_callback(self, *args):
        print(self.filename.get())

    def make_spec(self):
        ready = True
        err_message = []
        
        sheet = self.xl_path.get()
        if sheet == '' or sheet is None:
            ready = False
            err_message.append('Excel File (From AQ)')
            
        name = self.filename.get()
        if name == '' or name is None:
            ready = False
            err_message.append('Filename')

        save_here = self.dir_path.get()
        if save_here == '' or save_here is None:
            ready = False
            err_message.append('Save Location')

        to_compose = self.to_compose.get()
        if to_compose is None or to_compose == '':
            ready = False
            err_message.append('Compose')

        if not ready:
            err = 'The following fields are missing information: '
            for i in err_message:
                err += i + ', '
            err = err[:len(err) - 2]
            print(err)
            messagebox.showerror(message=err, title='Error')
        else:
            errors = spec_backend.main(sheet, name, save_here, to_compose)
            print(errors)
            if errors:
                error_items = ', '.join(errors)
                messagebox.showwarning(title='NIKEC Errors Encountered', \
                                       message='NIKEC items must be formatted as "NIKEC/ BY..." or "NIKEC, BY..."'+ \
                                             '\n\nError occured in item(s): ' + error_items) 
                messagebox.showinfo(title='Saved!', message='Your spec was created using the valid data available.\n\n' + \
                                    'Please review, revise in AutoQuotes, and regenerate spec items as needed.\n\n' +\
                                    'File saved to: ' + save_here)
            else: 
              messagebox.showinfo(title='Saved!', message='Files saved to ' +\
                                  save_here)

def pad_config(container, pad_amt):
    for child in container.winfo_children():
        child.grid_configure(padx=pad_amt, pady=pad_amt)
