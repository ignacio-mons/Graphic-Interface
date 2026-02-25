from ventana_poo import *
import tkinter as tk

# com = Comunication(port="COM1", baud=9600)
# com.conexion()
# bascula = Shell(com)

# print(bascula.real_weight())
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("Harlequin.json")


# if __name__ == "__main__":
#     com = Comunication(port="COM1", baud=9600)
#     com.conexion()
#     bascula = Shell(com)
#     root = Window(com, bascula)
#     # root.mainloop()
#     print(bascula.escribir_variable(40, 1))

if __name__ == "__main__":
    com = Comunication(port="", baud=9600)
    sh = Shell(com)
    root = Window(com, sh)
    root.mainloop()
