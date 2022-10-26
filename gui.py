from tkinter import *

window = Tk()
window.title('Spreadsheet to Email')


for element in (
    Entry(text='Send'),
    Button(text='Send'),
    Label(text='label'),
    ):
    element.pack()

#print(email.insert('hi'))
#print(email.get())


#window.geometry('1000x1000+10+20')

window.mainloop()
