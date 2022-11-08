# TODO remove sticky to see what happens
# TODO select first browse button
# TODO file manager button

from tkinter import *
from tkinter import ttk
from tkinter import filedialog

# layout
#   Files
#       Template   |||| [BROWSE]
#       CSV        |||| [BROWSE]
#       Attachment |||| [BROWSE]
#   Testing Configuration
#       ???
#   [OK] [CANCEL]

# Positions
col_label = 1
col_entry = 2
col_button = 3
row_files = 1
row_template = 2

def main():

    root = Tk()
    root.title('Speadsheet Slicer-dicer')

    mainframe = ttk.Frame(root, padding='3 3 12 12')
    mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # Files
    ttk.Label(mainframe, text='Files').grid(column=col_label, row=row_files, sticky=W) # Span

    # Template
    template = StringVar()
    template_entry = ttk.Entry(mainframe, width=50, textvariable=template)
    template_entry.grid(column=col_entry, row=row_template, sticky=(W, E))
    ttk.Label(mainframe, text='Template').grid(column=col_label, row=row_template, sticky=W)
    ttk.Button(mainframe, text='Browse', command=lambda *args: None).grid(column=col_button, row=row_template, sticky=W)

    filename = filedialog t

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

main()
