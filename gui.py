from tkinter import *

# layout
#   template name
#   csv
#   attachment

window = Tk()
window.title('Spreadsheet to Email')
Entry(text='XYZ'),

elements = []

template = Label(text='template file')
elements.extend([template,])

for element in (
    Button(text='Send'),
    Label(text='label'),
    ):
    element.pack()

#print(email.insert('hi'))
#print(email.get())


window.geometry('1000x1000+10+20')

window.mainloop()
