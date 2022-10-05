from tkinter import *
import tkinter as tk


root = Tk()

canvas= tk.Canvas(root, width=600, height=300)
canvas.grid(columnspan=3, rowspan=3)
# label = Label(root, text="Test")
# label.pack()

br_wrd_txt = tk.StringVar()
br_wrd_txt.set("Word")
br_wrd_btn = tk.Button(root, textvariable=br_wrd_txt, bg="#005094", fg="white", height=2, width=5)

br_xl_txt = tk.StringVar()
br_xl_txt.set("Excel")
br_xl_btn = tk.Button(root, textvariable=br_xl_txt, bg="#026c37", fg="white", height=2, width=5)

br_xl_btn.grid(column=2, row=2)
br_wrd_btn.grid(column=0, row=2)

root.mainloop()