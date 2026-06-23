from ventana_poo import *
import tkinter as tk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("breeze.json")

if __name__ == "__main__":
    import threading

    com = Comunication(port="", baud=9600)
    sh = Shell(com)
    root = Window(com, sh)
    root.mainloop()
