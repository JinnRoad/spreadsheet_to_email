from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import sys

root = Tk()
root.title('Speadsheet Slicer-dicer')

files = {'email': StringVar(), 'csv': StringVar(), 'attachment': StringVar()}

def main():

    mainframe = ttk.Frame(root, padding='3 3 12 12')
    mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
    root.columnconfigure(0, weight=5)
    root.rowconfigure(0, weight=5)

    # Buttons
    b1 = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('email'))
    b2 = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('csv'))
    b3 = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('attachment'))
    b4 = ttk.Button(mainframe, text='Send Emails', command=send_emails)

    # Labels
    l1 = ttk.Label(mainframe, text='Files')
    l2 = ttk.Label(mainframe, text='Email Template')
    l3 = ttk.Label(mainframe, text='CSV File')
    l4 = ttk.Label(mainframe, text='Attachment')
    l5 = ttk.Label(mainframe, textvariable=files['email'])
    l6 = ttk.Label(mainframe, textvariable=files['csv'])
    l7 = ttk.Label(mainframe, textvariable=files['attachment'])

    # Label positions
    l1.grid(column=1, row=1, sticky=W)  # Title
    l2.grid(column=1, row=2, sticky=W)  # Email label
    b1.grid(column=2, row=2, sticky=W)  # Email browse button
    l5.grid(column=3, row=2, sticky=W)  # Email filename
    l3.grid(column=1, row=3, sticky=W)  # CSV label
    b2.grid(column=2, row=3, sticky=W)  # CSV browse button
    l6.grid(column=3, row=3, sticky=W)  # CSV filename
    l4.grid(column=1, row=4, sticky=W)  # Attachment label
    b3.grid(column=2, row=4, sticky=W)  # Attachment browse button
    l7.grid(column=3, row=4, sticky=W)  # Attachment filename
    b4.grid(column=1, row=5, sticky=W)  # Okay

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

def browse_button(file):
    """folder path = StringVar()"""
    files[file].set(filedialog.askopenfile(
        title=f'Select {file}',
        filetypes=(('All files', '*.*'),)
        ).name)

def send_emails():
    print(*files.items(), sep='\n', end='\n\n', flush=True)

main()
