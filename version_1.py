from ventana_poo import Window, Comunication, Shell, ruta
import customtkinter as ctk
import os


def ventana():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme(ruta(os.path.join("Icon", "midnight.json")))

    com = Comunication(port="", baud=9600)
    sh = Shell(com)
    root = Window(com, sh)
    root.mainloop()


if __name__ == "__main__":
    ventana()
