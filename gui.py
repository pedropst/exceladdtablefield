from tkinter import *
from tkinter import filedialog as fd

import main

class Gui():
    def __init__(self):
        self.script = None
        root = Tk()
        root.title("EXCEL - ASSIGN CATEGORY")
        root.geometry("800x200")

        Label(root, text="Select the new report file: ").grid(row=0, column=0, padx=10, pady=10, sticky='W')
        entry_select_file = Entry(root, width=20)
        entry_select_file.grid(row=0, column=1, padx=10, pady=10, sticky='W')
        Button(root, text='Select File', command=lambda: self.select_file(entry_select_file)).grid(row=0, column=2, padx=10, pady=10, sticky='W')

        Label(root, text="Select the master file: ").grid(row=1, column=0, padx=10, pady=10, sticky='W')
        entry_master_file = Entry(root, width=20)
        entry_master_file.grid(row=1, column=1, padx=10, pady=10, sticky='W')
        Button(root, text='Select File', command=lambda: self.select_file(entry_master_file)).grid(row=1, column=2, padx=10, pady=10, sticky='W')


        Button(root, text='Run', command=lambda: self.call(entry_select_file.get(), entry_master_file.get(), root)).grid(row=2, column=1, padx=10, pady=10, sticky='W')

        root.mainloop()

        
    def select_file(self, entry):
        filetypes = [('Excel files', '*.xlsx')]
        path = fd.askopenfilename(filetypes=filetypes)
        entry.delete(0, END)
        entry.insert(0, path)

    def call(self, path_report, path_master, root):
        self.script = main.Categories()
        self.script.call(path_report, path_master, root)
        


g = Gui()

