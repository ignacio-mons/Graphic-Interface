from ventana_poo import *
import tkinter as tk

# com = Comunication(port="COM1", baud=9600)
# com.conexion()
# bascula = Shell(com)

# print(bascula.real_weight())
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("breeze.json")

# if __name__ == "__main__":
#     com = Comunication(port="COM1", baud=9600)
#     com.conexion()
#     bascula = Shell(com)
#     root = Window(com, bascula)
#     # root.mainloop()
#     print(bascula.escribir_variable(40, 1))

if __name__ == "__main__":
    import threading

    # hilo_b = threading.Thread(target=read_bar, daemon=True)
    # # hilo_h = threading.Thread(target=read_hidro, daemon=True)
    # hilo_b.start()
    # hilo_h.start()
    com = Comunication(port="", baud=9600)
    sh = Shell(com)
    root = Window(com, sh)
    root.mainloop()
