from tkinter import *
import Roll
import os

root = Tk()

e = Entry(root)


def sayHi():
    greet = Label(root, text="Hello " + str(e) + "!")
    greet.pack()

button1 = Button(root, text="Click Me!", command=sayHi)
button1.grid(row=0, column=0)


def openxl():
    os.startfile("Hi.txt")

button2 = Button(root, text="Open file!", command=openxl)
button2.grid(row=0, column=1)

root.mainloop()
