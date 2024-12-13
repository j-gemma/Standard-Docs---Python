#custom widgets for cutbook cover frontend
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import date


class AreaTable(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, relief='sunken', borderwidth=2)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.create_widgets()

    def create_widgets(self):
        self.table = ttk.Treeview(self, columns=('area', 'abbr'), show='headings')
        self.table['selectmode'] = 'extended'
        self.table.heading('area', text='Area')
        self.table.heading('abbr', text ='Abbr.')
        self.table.grid(row=0, column=0, sticky='news')
            
class ThreeButtons(ttk.Frame):
    def __init__(self, parent):
        self.parent = parent
        self.user_input = False
        self.confirming = False
        self.name_holder = ''
        self.abbr_holder = ''
        super().__init__(parent, borderwidth=2)
        self.create_widgets()

    def create_widgets(self):

        self.new = ttk.Button(self, text='Add', command=self.new_area)
        self.new.grid(row=0, column=0, sticky='nsew')

        self.edit = ttk.Button(self, text='Edit', command=self.edit_area)
        self.edit.grid(row=0, column=1, sticky='nsew', padx=5)

        self.remove = ttk.Button(self, text='Remove', command=self.remove_area)
        self.remove.grid(row=0, column=2, sticky='nsew')

        self.columnconfigure(0, weight=1)  
        self.columnconfigure(1, weight=1)  
        self.columnconfigure(2, weight=1)  

    def new_area(self):

        area_name = self.parent.entry_frame.area_name.get()
        area_abbr = self.parent.entry_frame.area_abbr.get()

        if area_name is None or area_name == '':
            messagebox.showerror("Error", "Area name cannot be empty.")
            return
        if len(area_abbr) > 3:
            messagebox.showerror("Error", "Area abbreviation cannot be longer than 3 characters.")
            return
        if not self.confirming:
            if self.parent.is_unique_entry(area_name, area_abbr):
                self.parent.existing_entries[area_abbr.lower()] = area_name.lower()
                inserted = self.parent.area_table.table.insert('', 'end', values=(area_name, area_abbr))
                print(self.parent.area_table.table.item(inserted))
                self.parent.entry_frame.area_name.set('')
                self.parent.entry_frame.area_abbr.set('')
              
    def edit_area(self):
        selected_item = self.parent.area_table.table.selection()
        

        if len(selected_item) == 1:
            area_name, area_abbr = self.parent.area_table.table.item(selected_item, "values")
            self.name_holder = area_name
            self.abbr_holder = area_abbr
            self.parent.entry_frame.area_name.set(area_name)
            self.parent.entry_frame.area_abbr.set(area_abbr)
            self.remove_area()
            self.parent.area_table.table['selectmode'] = 'none'
            self.parent.status.set('Editing...') 
            self.edit['text'] = 'Confirm'
            self.edit['command'] = self.confirm_edit
            self.remove['command'] = ''
            self.new['text'] = 'Cancel'
            self.new['command'] = self.cancel_edit
        elif len(selected_item) > 1:
            messagebox.showinfo('Edit', "Only one item may be edited at a time")
        else:
            messagebox.showinfo('Edit', "Please select an area to edit")

    def cancel_edit(self):
        self.parent.area_table.table['selectmode'] = 'extended'
        self.edit['text'] = 'Edit'
        self.edit['command'] = self.edit_area
        self.remove['command'] = self.remove_area
        self.new['text'] = 'New'
        self.new['command'] = self.new_area
        self.parent.entry_frame.area_name.set(self.name_holder)
        self.parent.entry_frame.area_abbr.set(self.abbr_holder)
        self.new_area()
        self.parent.status.set('')
        self.name_holder = ''
        self.abbr_holder = ''

    def confirm_edit(self):
        self.new_area()
        self.parent.area_table.table['selectmode'] = 'extended'
        self.edit['text'] = 'Edit'
        self.edit['command'] = self.edit_area
        self.remove['command'] = self.remove_area
        self.confirming = False
        self.new['text'] = 'New'
        self.new['command'] = self.new_area
        self.parent.status.set('')

    def remove_area(self):
        selected_items = self.parent.area_table.table.selection()

        if selected_items:
            for selected_item in selected_items:
                area_name, area_abbr = self.parent.area_table.table.item(selected_item, "values")
                area_name_lower = area_name.lower()
                area_abbr_lower = area_abbr.lower()
                self.parent.area_table.table.delete(selected_item)
                self.parent.existing_entries.pop(area_abbr_lower)
        else:
            messagebox.showinfo("Remove", "Please select an item to remove.")

class ProjectEntry(ttk.Frame):
    def __init__(self, parent):
        self.parent = parent
        super().__init__(parent, relief='groove', borderwidth=2)
        self.create_widgets()

    def create_widgets(self):
        job_lbl_lbl = ttk.Label(self, text='Project: ', relief='flat')
        job_lbl_lbl.grid(column=0, row=0, sticky='nw')

        self.job_title = StringVar()
        self.job_title.trace_add('write', self.job_title_callback)
        job_label = ttk.Entry(self, textvariable=self.job_title, width=35)
        job_label.grid(column=1, row=0, sticky='new')

        proj_num_lbl = ttk.Label(self, text='Number: ', relief='flat')
        proj_num_lbl.grid(column=2, row=0, sticky='new')

        self.proj_num = StringVar()
        self.proj_num.trace_add('write', self.proj_num_callback)
        proj_label = ttk.Entry(self, textvariable=self.proj_num, width=15)
        proj_label.grid(column=3, row=0, sticky='ne')

        self.columnconfigure(0, weight=1)  
        self.columnconfigure(1, weight=1)  
        self.columnconfigure(2, weight=1)  
        self.columnconfigure(3, weight=1)  

    def proj_num_callback(self, *args):
        print(self.proj_num.get())
        
    def job_title_callback(self, *args):
        print(self.job_title.get())

class InputBox(Toplevel):
    def __init__(self, parent, title, prompt, size):
        super().__init__(parent)
        self.setup_window(title, size)
        self.create_widgets(prompt)
        self.loc = None
        self.date = None

    def setup_window(self, title, size):
        self.title(title)
        x = (self.winfo_screenwidth() - size[0]) // 2
        y = (self.winfo_screenheight() - size[1]) // 2
        self.geometry(f'{size[0]}x{size[1]}+{x}+{y}')

    def create_widgets(self, prompt):
        ttk.Label(self, text=prompt).grid(row=0, column=0, sticky='nsew')
        self.location_var = StringVar()
        self.location_entry = ttk.Entry(self, textvariable=self.location_var)
        self.location_entry.grid(row=1, column=0, sticky='nsew')
        ttk.Button(self, text='OK', command=self.on_ok).grid(row=4, column=0, sticky='nsew')


        ttk.Label(self, text='Enter a date for covers to be generated')\
                        .grid(row=2, column=0, sticky='nsew')
        self.date_var = StringVar()
        self.date_entry = ttk.Entry(self, textvariable=self.date_var)
        self.date_entry.grid(row=3, column=0, sticky='nsew')
        pad_config(self, 2)

    def on_ok(self):
        loc = self.location_var.get().strip()
        if loc != '':
            self.loc = loc 

        date = self.date_var.get()
        if date != '':
            self.date = date
            
        self.destroy()

class AreaEntry(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, relief='groove', borderwidth=2)
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        self.area_name = StringVar()
        self.area_name.trace_add("write", self.area_name_callback)
        self.area = ttk.Entry(self, textvariable=self.area_name, width=25)
        self.area.grid(row=0, column=0)

        self.area_abbr = StringVar()
        self.area_abbr.trace_add("write", self.area_abbr_callback)
        self.abbr = ttk.Entry(self, textvariable=self.area_abbr, width=5)
        self.abbr.grid(row=0, column=1, sticky='ne')

        self.columnconfigure(0, weight=1)  
        self.columnconfigure(1, weight=1)  

    def area_name_callback(self, *args):
        print(self.area_name.get())

    def area_abbr_callback(self, *args):
        print(self.area_abbr.get())



def destroy():
    self.parent.destroy()

def ok(input_frame):
    input_frame.parent.loc = input_frame.entry.get()
    input_frame.parent.cover_date = input_frame.entry_2.get()
    input_frame.parent.destroy()

def cancel(input_frame):
    input_frame.parent.destroy()

def pad_config(widget, pad_amt):
    for child in widget.winfo_children():
        child.grid_configure(padx=pad_amt, pady=pad_amt)


