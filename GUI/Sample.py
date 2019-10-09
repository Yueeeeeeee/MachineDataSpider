from tkinter import *
from tkinter import ttk
import tkinter as tk

fenster = Tk()
fenster.title("Searching for the Machines")
fenster.geometry("500x1000")

Label_Height = Label(fenster, text = "Height of the workpiece",bg = "yellow", fg = "black")
Label_Width = Label(fenster, text = "Width of the workpiece", bg = "yellow", fg = "black")
Label_Length = Label(fenster, text = "Length of the workpiece", bg = "yellow", fg = "black")
Label_Type = Label(fenster, text = "Machine Type",  bg = "red", fg = "black")

Label_Unit1 = Label(fenster, text = "mm", fg = "black")
Label_Unit2 = Label(fenster, text = "mm", fg = "black")
Label_Unit3 = Label(fenster, text = "mm", fg = "black")

Label_Height.place(x = 1, y = 5 , width=200, height=15)
Label_Width.place(x = 1, y = 25 , width=200, height=15)
Label_Length.place(x = 1, y = 45 , width=200, height=15)

Label_Unit1.place(x = 213, y = 5 , width=120, height=15)
Label_Unit2.place(x = 213, y = 25 , width=120, height=15)
Label_Unit3.place(x = 213, y = 45 , width=120, height=15)

eingabe_Height = Entry(fenster)
eingabe_Height.place(x = 210, y = 5 , width=50, height=15)
eingabe_Width = Entry(fenster)
eingabe_Width.place(x = 210, y = 25 , width=50, height=15)
eingabe_Length = Entry(fenster)
eingabe_Length.place(x = 210, y = 45 , width=50, height=15)

Machine_Type = tk.StringVar()
MachineChosen = ttk.Combobox(fenster, width=200, textvariable=Machine_Type)
MachineChosen['values'] = ('Turning Machines', 'Machine centers/Milling Machines', 'Milling and Turning Combi-Machines')
MachineChosen.place(x = 213, y = 85 , width=220, height=25)

Label_Type.place(x = 1, y = 85 , width=200, height=15)

# def ausgabe():
#     if int(eingabe.get()) < 2:
#         print('Gut')
#     else:
#         print('Normal')


# Knopf1 = Button(fenster, text="HAHA", command = ausgabe)
# Knopf1.place(x = 20, y = 30 , width=120, height=25)


mainloop()