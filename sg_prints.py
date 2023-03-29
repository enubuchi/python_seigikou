import sys
import tkinter as tk
from tkinter.constants import PAGES

root = tk.Tk()
root.title(u"枚数入力")
root.geometry("250x150")

label1 = tk.Label(root,text="枚数入力")
label1.place(x=20,y=20)
entry1 = tk.Entry(root,width="6")
entry1.place(x=120,y=20)
label1mai = tk.Label(root,text="枚")
label1mai.place(x=180,y=20)

var = tk.IntVar()
var.set(1)

label2 = tk.Label(root,text="オンライン開催")
label2.place(x=20,y=50)
radio2y = tk.Radiobutton(
    text="する",
    value=1,
    variable=var
)
radio2y.place(x=20,y=70)
radio2n = tk.Radiobutton(
    text="しない",
    value=0,
    variable=var
)
radio2n.place(x=80,y=70)

button3=tk.Button(root,width=10,text="作成実行")
button3.place(x=90,y=110)

#pagezで、entry1.get()中の値を保持。オンライン開催する場合はvar.get()が1

def click():
    pagez = entry1.get()
    entry1.delete(0, tk.END)
    label0 = tk.Label(text=pagez) 
    label0.place(x=0, y=120)
    print901 = var.get()
    label901 = tk.Label(text=print901) 
    label901.place(x=190, y=120)  

button3["command"] = click

root.mainloop()

