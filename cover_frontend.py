from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import date
import cover_backend
import cover_widgets_custom as cw

class CutCover(Toplevel):
    def __init__(self, title, size):
        super().__init__()
        self.setup_window(title, size)
        self.add_menu()
        
    def add_menu(self):
        menu = CoverOptions(self)
        menu.grid(column=0, row=0)

    def setup_window(self, title, size):
        self.title(title)
        x = (self.winfo_screenwidth() - size[0]) // 2
        y = (self.winfo_screenheight() - size[1]) // 2
        self.geometry(f'{size[0]}x{size[1]}+{x}+{y}')

class CoverOptions(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.existing_entries = {}
        self.grid(column=0, row=0, sticky='nsew')
        self.create_widgets()
        
    def create_widgets(self):

        today = date.today().strftime('%B %d, %Y')
        today_formatted = date.today().strftime('%y%m%d')
        self.today_is = 'Today is ' + today + ' [' + today_formatted + ']'

        self.date_string = StringVar()
        self.date_string.set(self.today_is)
        you_r_here = ttk.Label(self, textvariable=self.date_string)
        you_r_here.grid(row=0, column=1) 

        info = ttk.Label(self, text='*Information below should correctly autopopulate when you choose a directory.*')
        info.grid(column=1, row=3, sticky=E)

        ttk.Label(self, text='File(s) to be saved at:').grid(column=0, row=2, sticky=E)
        ttk.Label(self, text='Project Location:').grid(column=0, row=1, sticky=E)
        ttk.Button(self, text='Edit', command=self.modal_input_box).grid(row=1, column =2, sticky=E)
        
        self.dir_path = StringVar()
        self.dir_path.trace_add('write', self.dir_path_callback)
        directory = ttk.Label(self, textvariable=self.dir_path, width=55, relief='sunken')
        directory.grid(column=1, row=2, sticky='ew')

        ttk.Button(self, text="Generate Covers", command=self.make_covers).grid(row=20, column=1)
        ttk.Button(self, text='Pick Directory', command=self.dir_pick).grid(column=2, row=2, sticky='sew')
        
        self.area_table = cw.AreaTable(self)
        self.area_table.grid(column=1, row=6, sticky='nsew')

        self.main_buttons = cw.ThreeButtons(self)
        self.main_buttons.grid(column=1, row=14, sticky='nsew')

        self.entry_lbl = ttk.Label(self, text='Area | Abbreviation')
        self.entry_lbl.grid(column=1, row=13, sticky='ns')

        self.entry_frame = cw.AreaEntry(self)
        self.entry_frame.grid(column=1, row=12)

        self.status = StringVar()
        self.status.trace_add('write', self.status_callback)
        self.entry_status = ttk.Label(self, textvariable=self.status)
        self.entry_status.grid(column=2, row=12, sticky='nw')

        self.proj_frame = cw.ProjectEntry(self)
        self.proj_frame.grid(column=1, row=4)
        
        self.loc = StringVar()
        self.loc_lbl = ttk.Label(self, textvariable=self.loc, relief='sunken')
        self.loc_lbl.grid(column=1, row=1, sticky='new')

        pad_config(self, 5)
        self.modal_input_box()

        
    def modal_input_box(self):
        self.dlg = cw.InputBox(self, 'Location', "Where is this project located? (i.e. 'Rockville, MD')", (300, 150))
        self.dlg.protocol("WM_DELETE_WINDOW", self.dismiss)
        self.dlg.transient(self)
        self.dlg.wait_visibility()
        self.dlg.grab_set()
        self.dlg.wait_window()
        if self.dlg.loc != None:
            self.loc.set(self.dlg.loc)
            
        if self.dlg.date != None:
            self.date_string.set(self.today_is + ' | Date on covers: '\
                             + self.dlg.date)

    def dismiss(self):
        self.dlg.grab_release()
        self.dlg.destroy()
    
    def dir_pick(self):
        path = filedialog.askdirectory(title='File(s) to be saved at:')
        self.dir_path.set(path)
        self.proj_frame.proj_num.set(path[path.index('/') + 1:path.index('-') - 1])
        self.proj_frame.job_title.set(path[path.index('-') + 2:path.index('/', 3)])
        
    def dir_path_callback(self, *args):
        print(self.dir_path.get())

    def status_callback(self, *args):
        print(self.status.get())

    def is_unique_entry(self, area_name, area_abbr):
        # Check if the area name or abbreviation already exists
        name_lower = area_name.lower()
        abbr_lower = area_abbr.lower()
        if name_lower in self.existing_entries.values():
            messagebox.showinfo("Duplicate", "This area name already exists.")
            return False
        if abbr_lower in self.existing_entries.keys():
            messagebox.showinfo("Duplicate", "This abbreviation already exists.")
            return False
        # No duplicates found, add to tracking sets and return True
        
        return True

    def make_covers(self):
        ready = True
        err_message = []
        
        TEMPLATE_PATH = r"Q:\Standard Documents\Templates\CUTBOOK COVER.docx"
        
        areas = self.existing_entries

        if len(areas) == 0:
            ready = False
            err_message.append('Area List')
        
        proj_num = self.proj_frame.proj_num.get()

        if proj_num is None or proj_num == '':
            ready = False
            err_message.append('Project Number')
            
        proj_title = self.proj_frame.job_title.get()

        if proj_title is None or proj_title == '':
            ready = False
            err_message.append('Project Title')
        
        directory = self.dir_path.get()

        if directory is None or directory == '':
            ready = False
            err_message.append('Directory')
        
        cover_date = self.dlg.date

        if cover_date is None or cover_date == '':
            ready = False
            err_message.append('Cover Date')
            
        location = self.dlg.loc

        if location is None or location == '':
            ready = False
            err_message.append('Location')

        if not ready:
            err = 'The following fields are missing information: '
            for i in err_message:
                err += i + ', '
            err = err[:len(err) - 2]
            print(err)
            messagebox.showerror(message=err, title='Error')
        else:
            cover_backend.main(areas, proj_num, proj_title, directory, location, cover_date)
            messagebox.showinfo(title='Saved!', message='Files saved to ' +\
                                directory)

def pad_config(widget, pad_amt):
    for child in widget.winfo_children():
        child.grid_configure(padx=pad_amt, pady=pad_amt)
