import tkinter as tk
import customtkinter as ctk
import serial
import csv
import time
from tkinter import ttk
from CTkMenuBarPlus import *
from datetime import datetime
from CTkMessagebox import CTkMessagebox
from tkinter.filedialog import asksaveasfilename, askopenfilename
from PIL import Image
import serial.tools.list_ports
import os
from fpdf import FPDF


def ruta(r_relativa):
    import sys, os

    try:
        b_path = sys._MEIPASS
    except Exception:
        b_path = os.path.abspath(".")
    return os.path.join(b_path, r_relativa)


class Comunication:
    """
    Establecer la comunicación serial con MT
    """

    def __init__(self, port: str, baud: int):
        self.port = port
        self.baud = baud
        self.ser = None
        self.conectado = True

    def conexion(self):
        """
        metodo de conexion
        """
        if self.ser and self.ser.is_open:
            self.ser.close()

        try:
            self.ser = serial.Serial(
                port=self.port,
                baudrate=self.baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                timeout=0.5,
            )
            time.sleep(2)
            self.ser.reset_input_buffer()
            print(f"M-Conexion: {self.port}")
            self.conectado = True
            return True
        except Exception as e:
            self.conectado = False
            print(f"Error: M-Conexion: {e}")
            return False

    def envio(self, data: str):
        """
        Envio generico
        """
        if self.ser and self.ser.is_open:
            try:
                self.ser.write((data + "\r\n").encode("ascii"))
                # print("M-Send")
            except Exception as e:
                print(f"Error M-Send: {e}")
        # else:
        # print("P cerrado")

    def respuesta(self):
        """
        Respuesta del write ascii
        """
        if self.ser and self.ser.is_open:
            return self.ser.readline().decode("ascii", errors="ignore").strip()

    def cerrar_puerto(self):
        if self.ser and self.ser.is_open:
            self.ser.close()


class Shell:
    """
    Manda las intrucciones zero, tara, peso
    """

    def __init__(self, comando: Comunication):
        self.comando = comando

    def peso_instantaneo(self):
        """
        Envia el peso al instante este o no estable
        """
        self.comando.envio("SI")
        lecture = self.comando.respuesta()
        if lecture:
            return lecture.split()
        return []

    def zero(self):
        """
        Envia el cero
        """
        self.comando.envio("Z")
        lecture = self.comando.respuesta()
        return print("Zero enviado")

    def tara(self):
        """
        Envia la tara
        """
        self.comando.envio("T")
        self.comando.respuesta()
        return print("Tara enviada")

    def quitar_tara(self):
        """
        Quita la tara si es que tiene
        """
        self.comando.envio("TAC")
        self.comando.respuesta()
        return print("Quite Tara")

    def peso_estable(self):
        """
        Envia el peso cuando esta estable
        """
        self.comando.envio("S")
        lecture = self.comando.respuesta()
        print("peso estable")
        if "S +" in lecture:
            return "OVERLOAD"
        if lecture:
            return lecture.split()
        return []

    def cali_cero(self):
        try:
            self.comando.envio("Z")
            lecture = self.comando.respuesta()
            return lecture
        except Exception as e:
            print(f"Error cali cero {e}")

    def consulta_datos(self):
        try:
            self.comando.envio("I2")
            lecture = self.comando.respuesta()
            return lecture
        except Exception as e:
            print(f"errro{e}")

    def leer_variable(self, indice):
        try:
            comando_i = f"R{indice}\r\n"
            self.comando.ser.write(comando_i.encode("ascii"))
            respuesta = self.comando.respuesta()

            if respuesta:
                partes = respuesta.split()
                if len(partes) >= 2:
                    return True, partes[1]
            return False, None
        except Exception as e:
            print(f"Error RR {e}")
            return False, None

    def escribir_variable(self, indice, valor):
        try:
            comando_i = f"W{indice} {valor}\r\n"
            self.comando.ser.write(comando_i.encode("ascii"))
            respuesta = self.comando.ser.read(1)

            if respuesta and respuesta[0] == 0x06:
                print(f"good Variable {indice}, {valor}")
                return True
            else:
                print(f"error {indice}")
                return False
        except Exception as e:
            print(f"Error escribir_v {e}")
            return False

    def calibrar_cero(self):
        lecture = self.peso_instantaneo()
        if len(lecture) >= 3:
            peso_actual = float(lecture[2])
            estado = lecture[1]

            if abs(peso_actual) > 5:
                return False, f"No vacia"
            if estado != "S":
                return False, "No estable"

        exito = self.escribir_variable(40, 1)
        if exito:
            time.sleep(2)

            lecture = self.peso_instantaneo()
            if len(lecture) >= 3:
                peso_final = float(lecture[2])
                return True, f"Calibrado {peso_final}"
            else:
                return True, f"Calibraco goos"
        else:
            return False, "Error calibracion"

    def obtener_tara(self):
        if self.comando.ser and self.comando.ser.is_open:
            try:
                self.comando.envio("TA")
                lecture = self.comando.respuesta()
                partes = lecture.split()

                if len(partes) >= 4:
                    return f"{partes[2]} {partes[3]}"
                return "0.0"
            except Exception as e:
                return f"error {e}"

        return "--"


class ventana_calibracion(ctk.CTkToplevel):
    def __init__(self, master, shell, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.shell = shell
        self.title("Ajustes")
        self.geometry("1300x700+150+50")
        self.grab_set()

        self.widgets()

    def widgets(self):
        # ctk.CTkLabel(self, text="Certificado", font=("Cambria", 20)).pack(pady=20)
        self.frame_formato = ctk.CTkFrame(self)
        self.frame_formato.pack(fill="both", expand=True, padx=5, pady=5)
        self.frame_formato.pack_propagate(False)

        # ----------------------------------------
        self.frame_pesas = ctk.CTkFrame(self.frame_formato, width=500, height=800)
        self.frame_pesas.grid(row=0, column=0, padx=5, pady=5)
        # self.frame_pesas.pack(side="left", padx=5, pady=5)
        self.frame_pesas.pack_propagate(False)

        # ----------------------------------------

        self.label_pesas = ctk.CTkLabel(
            self.frame_pesas,
            justify="center",
            text="Juego de Pesas",
            font=("Cambria", 20),
        )
        self.label_pesas.grid(row=0, column=0, padx=10, pady=10)

        # ------------------------------------------
        self.frame_pesa1 = ctk.CTkFrame(self.frame_pesas, width=400, height=100)
        self.frame_pesa1.grid(row=1, column=0, padx=10, pady=10)

        self.frame_pesa1.pack_propagate(False)
        # # ------------------------------------------

        self.label_marca_p1 = ctk.CTkLabel(
            self.frame_pesa1, text="Marca:", font=("Cambria", 12)
        )

        self.label_marca_p1.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.entrada_marca_p1 = ctk.CTkEntry(
            self.frame_pesa1,
            justify="center",
            placeholder_text="Pesa 1",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_marca_p1.grid(row=0, column=1, padx=5, pady=5)

        # --------------------------------------
        self.label_identificacion_p1 = ctk.CTkLabel(
            self.frame_pesa1, text="Identificación:", font=("Cambria", 12)
        )
        self.label_identificacion_p1.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entrada_identificacion_p1 = ctk.CTkEntry(
            self.frame_pesa1,
            justify="center",
            placeholder_text="ID Pesa 1",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_identificacion_p1.grid(row=0, column=3, padx=5, pady=5)

        # ------------------------------

        self.label_modelo_p1 = ctk.CTkLabel(
            self.frame_pesa1, text="Modelo:", font=("Cambria", 12)
        )
        self.label_modelo_p1.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entrada_modelo_p1 = ctk.CTkEntry(
            self.frame_pesa1,
            justify="center",
            placeholder_text="Modelo Pesa 1",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_modelo_p1.grid(row=1, column=1, padx=5, pady=5)

        # -------------------------------

        self.label_exactitud_p1 = ctk.CTkLabel(
            self.frame_pesa1, text="Clase de exactitud:", font=("Cambria", 12)
        )
        self.label_exactitud_p1.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entrada_exactitud_p1 = ctk.CTkEntry(
            self.frame_pesa1,
            justify="center",
            placeholder_text="Exactitud Pesa 1",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_exactitud_p1.grid(row=1, column=3, padx=5, pady=5)
        # --------------------------------
        self.label_serie_p1 = ctk.CTkLabel(
            self.frame_pesa1, text="Número de serie:", font=("Cambria", 12)
        )
        self.label_serie_p1.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entrada_serie_p1 = ctk.CTkEntry(
            self.frame_pesa1,
            justify="center",
            placeholder_text="Serie Pesa 1",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_serie_p1.grid(row=2, column=1, padx=5, pady=5)

        # ---------------------------------

        self.label_certificado_p1 = ctk.CTkLabel(
            self.frame_pesa1, text="Certificado:", font=("Cambria", 12)
        )
        self.label_certificado_p1.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.entrada_certificado_p1 = ctk.CTkEntry(
            self.frame_pesa1,
            justify="center",
            placeholder_text="Certificado Pesa 1",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_certificado_p1.grid(row=2, column=3, padx=5, pady=5)

        # --------BOTONES P1--------------------------
        self.button_guardar_p1 = ctk.CTkButton(
            self.frame_pesa1,
            text="Guardar",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
        )
        self.button_guardar_p1.grid(row=3, column=0, pady=5)

        self.button_abrir_p1 = ctk.CTkButton(
            self.frame_pesa1,
            text="Abrir",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
        )
        self.button_abrir_p1.grid(row=3, column=3, columnspan=4, pady=5)

        # # ------------------------
        self.frame_pesa2 = ctk.CTkFrame(self.frame_pesas, width=425, height=100)
        self.frame_pesa2.grid(row=2, column=0, padx=10, pady=10)
        self.frame_pesa2.pack_propagate(False)
        # ------------------------------------------
        # ++++++++++++++++++
        self.label_marca_p2 = ctk.CTkLabel(
            self.frame_pesa2, text="Marca:", font=("Cambria", 12)
        )

        self.label_marca_p2.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.entrada_marca_p2 = ctk.CTkEntry(
            self.frame_pesa2,
            justify="center",
            placeholder_text="Pesa 2",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_marca_p2.grid(row=0, column=1, padx=5, pady=5)

        # --------------------------------------
        self.label_identificacion_p2 = ctk.CTkLabel(
            self.frame_pesa2, text="Identificación:", font=("Cambria", 12)
        )
        self.label_identificacion_p2.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entrada_identificacion_p2 = ctk.CTkEntry(
            self.frame_pesa2,
            justify="center",
            placeholder_text="ID Pesa 2",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_identificacion_p2.grid(row=0, column=3, padx=5, pady=5)

        # ------------------------------

        self.label_modelo_p2 = ctk.CTkLabel(
            self.frame_pesa2, text="Modelo:", font=("Cambria", 12)
        )
        self.label_modelo_p2.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entrada_modelo_p2 = ctk.CTkEntry(
            self.frame_pesa2,
            justify="center",
            placeholder_text="Modelo Pesa 2",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_modelo_p2.grid(row=1, column=1, padx=5, pady=5)

        # -------------------------------

        self.label_exactitud_p2 = ctk.CTkLabel(
            self.frame_pesa2, text="Clase de exactitud:", font=("Cambria", 12)
        )
        self.label_exactitud_p2.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entrada_exactitud_p2 = ctk.CTkEntry(
            self.frame_pesa2,
            justify="center",
            placeholder_text="Exactitud Pesa 2",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_exactitud_p2.grid(row=1, column=3, padx=5, pady=5)
        # --------------------------------
        self.label_serie_p2 = ctk.CTkLabel(
            self.frame_pesa2, text="Número de serie:", font=("Cambria", 12)
        )
        self.label_serie_p2.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entrada_serie_p2 = ctk.CTkEntry(
            self.frame_pesa2,
            justify="center",
            placeholder_text="Serie Pesa 2",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_serie_p2.grid(row=2, column=1, padx=5, pady=5)

        # ---------------------------------

        self.label_certificado_p2 = ctk.CTkLabel(
            self.frame_pesa2, text="Certificado:", font=("Cambria", 12)
        )
        self.label_certificado_p2.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.entrada_certificado_p2 = ctk.CTkEntry(
            self.frame_pesa2,
            justify="center",
            placeholder_text="Certificado Pesa 2",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_certificado_p2.grid(row=2, column=3, padx=5, pady=5)

        # --------BOTONES P2--------------------------
        self.button_guardar_p2 = ctk.CTkButton(
            self.frame_pesa2,
            text="Guardar",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
        )
        self.button_guardar_p2.grid(row=3, column=0, pady=5)

        self.button_abrir_p2 = ctk.CTkButton(
            self.frame_pesa2,
            text="Abrir",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
        )
        self.button_abrir_p2.grid(row=3, column=3, columnspan=4, pady=5)

        # +++++++++++++++++++
        self.frame_pesa3 = ctk.CTkFrame(self.frame_pesas, width=425, height=100)
        self.frame_pesa3.grid(row=3, column=0, padx=10, pady=10)
        self.frame_pesa3.pack_propagate(False)
        # ------------------

        self.label_marca_p3 = ctk.CTkLabel(
            self.frame_pesa3, text="Marca:", font=("Cambria", 12)
        )

        self.label_marca_p3.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.entrada_marca_p3 = ctk.CTkEntry(
            self.frame_pesa3,
            justify="center",
            placeholder_text="Pesa 3",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_marca_p3.grid(row=0, column=1, padx=5, pady=5)

        # --------------------------------------
        self.label_identificacion_p3 = ctk.CTkLabel(
            self.frame_pesa3, text="Identificación:", font=("Cambria", 12)
        )
        self.label_identificacion_p3.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entrada_identificacion_p3 = ctk.CTkEntry(
            self.frame_pesa3,
            justify="center",
            placeholder_text="ID Pesa 3",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_identificacion_p3.grid(row=0, column=3, padx=5, pady=5)

        # ------------------------------

        self.label_modelo_p3 = ctk.CTkLabel(
            self.frame_pesa3, text="Modelo:", font=("Cambria", 12)
        )
        self.label_modelo_p3.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entrada_modelo_p3 = ctk.CTkEntry(
            self.frame_pesa3,
            justify="center",
            placeholder_text="Modelo Pesa 3",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_modelo_p3.grid(row=1, column=1, padx=5, pady=5)

        # -------------------------------

        self.label_exactitud_p3 = ctk.CTkLabel(
            self.frame_pesa3, text="Clase de exactitud:", font=("Cambria", 12)
        )
        self.label_exactitud_p3.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entrada_exactitud_p3 = ctk.CTkEntry(
            self.frame_pesa3,
            justify="center",
            placeholder_text="Exactitud Pesa 3",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_exactitud_p3.grid(row=1, column=3, padx=5, pady=5)
        # --------------------------------
        self.label_serie_p3 = ctk.CTkLabel(
            self.frame_pesa3, text="Número de serie:", font=("Cambria", 12)
        )
        self.label_serie_p3.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entrada_serie_p3 = ctk.CTkEntry(
            self.frame_pesa3,
            justify="center",
            placeholder_text="Serie Pesa 3",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_serie_p3.grid(row=2, column=1, padx=5, pady=5)

        # ---------------------------------

        self.label_certificado_p3 = ctk.CTkLabel(
            self.frame_pesa3, text="Certificado:", font=("Cambria", 12)
        )
        self.label_certificado_p3.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.entrada_certificado_p3 = ctk.CTkEntry(
            self.frame_pesa3,
            justify="center",
            placeholder_text="Certificado Pesa 3",
            width=100,
            font=("Cambria", 12),
        )
        self.entrada_certificado_p3.grid(row=2, column=3, padx=5, pady=5)

        # ----------------------------
        # --------BOTONES P3--------------------------
        self.button_guardar_p3 = ctk.CTkButton(
            self.frame_pesa3,
            text="Guardar",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
        )
        self.button_guardar_p3.grid(row=3, column=0, pady=5)

        self.button_abrir_p3 = ctk.CTkButton(
            self.frame_pesa3,
            text="Abrir",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
        )
        self.button_abrir_p3.grid(row=3, column=3, columnspan=4, pady=5)

        # # -------------------------------------------

        self.frame_ambiental = ctk.CTkFrame(self.frame_formato, width=425, height=100)
        self.frame_ambiental.grid(row=0, column=1, padx=5, pady=5)
        # self.frame_pesas.pack(side="left", padx=5, pady=5)
        self.frame_ambiental.pack_propagate(False)

        self.label_ambiental = ctk.CTkLabel(
            self.frame_ambiental,
            justify="center",
            text="Equipo Ambiental",
            font=("Cambria", 20),
        )
        self.label_ambiental.grid(row=0, column=0, padx=10, pady=10)
        # -------------------------------------------

        self.frame_ambiental1 = ctk.CTkFrame(
            self.frame_ambiental, width=425, height=100
        )
        self.frame_ambiental1.grid(row=1, column=0, padx=10, pady=10)
        self.frame_ambiental1.pack_propagate(False)

        # ---------------------------
        self.label_marca_am1 = ctk.CTkLabel(
            self.frame_ambiental1, text="Marca:", font=("Cambria", 12)
        )
        self.label_marca_am1.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entrada_marca_am1 = ctk.CTkEntry(
            self.frame_ambiental1, justify="center", placeholder_text="Marca", width=100
        )
        self.entrada_marca_am1.grid(row=0, column=1, padx=10, pady=10)
        # ---------------------------
        self.label_identificacion_am1 = ctk.CTkLabel(
            self.frame_ambiental1, text="Identificación:", font=("Cambria", 12)
        )
        self.label_identificacion_am1.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entrada_identificacion_am1 = ctk.CTkEntry(
            self.frame_ambiental1, justify="center", placeholder_text="ID", width=100
        )
        self.entrada_identificacion_am1.grid(row=0, column=3, padx=5, pady=5)

        # -------------------------------------------
        self.frame_ambiental2 = ctk.CTkFrame(
            self.frame_ambiental, width=425, height=100
        )
        self.frame_ambiental2.grid(row=2, column=0, padx=10, pady=10)
        # -------------------------------------------
        self.frame_ambiental2 = ctk.CTkFrame(
            self.frame_ambiental, width=425, height=100
        )
        self.frame_ambiental2.grid(row=3, column=0, padx=10, pady=10)


class repetibilidad(ctk.CTkToplevel):
    def __init__(self, master, shell, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.shell = shell
        self.title("Repetibilidad")
        self.geometry("800x700+150+50")
        self.grab_set()

        self.widgets()
        self.peso_al_momento()


class Window(ctk.CTk):
    """
    clase principal de la ventana, aqui se crean los widgets, los menus y se actualiza el peso al momento
    """

    def __init__(self, comunicacion: Comunication, shell: Shell):
        super().__init__()

        self.shell = shell
        self.ultimo_peso = 0
        self.comunicacion = comunicacion
        self.tara_activa = False
        self.val_tara = 0

        # ---------------------------
        self.peso_maximo = 58000
        self.alerta_peso_maximo = False
        # ---------------------------

        self.title("MT")
        self.geometry("1000x600+100+100")

        if os.path.exists(
            ruta(
                os.path.join(
                    "Icon",
                    "CNM_color_100dpi.ico",
                )
            )
            # r"Icon\weight_fat_body_overweight_unhealthy_obesity_diet_health_belly_icon_133507.ico"
        ):
            self.iconbitmap(
                ruta(
                    os.path.join(
                        "Icon",
                        "CNM_color_100dpi.ico",
                    )
                )
                # r"Icon\weight_fat_body_overweight_unhealthy_obesity_diet_health_belly_icon_133507.ico"
            )
        self.b_menus()
        self.widgets()
        # self.b_menus()
        # self.barra_menus()
        self.peso_al_momento()

    def widgets(self):
        """
        Son los metodos propios de la ventana
        """
        self.tabla_frame = ctk.CTkFrame(self, width=400, height=600)
        self.tabla_frame.pack(side="left", fill="both", padx=10, pady=10)

        # ------------TABLAAA----------------------
        self.style_tabla = ttk.Style()
        self.style_tabla.theme_use("clam")
        self.style_tabla.configure(
            "Treeview",
            background="#BD9D9D",
            foreground="black",
            rowheight=25,
            fieldbackground="#DFF2F5",
            font=("Cambria", 10),
        )

        self.tabla = ttk.Treeview(
            self.tabla_frame, columns=("incremento", "indicacion", "fecha", "hora")
        )

        self.tabla.heading("incremento", text="Incremento")
        self.tabla.heading("indicacion", text="Indicación")
        self.tabla.heading("fecha", text="Fecha")
        self.tabla.heading("hora", text="Hora")
        self.tabla.heading("#0", text="#")

        self.tabla.column("incremento", width=80, anchor="center")
        self.tabla.column("indicacion", width=100, anchor="center")
        self.tabla.column("fecha", width=100, anchor="center")
        self.tabla.column("hora", width=100, anchor="center")
        self.tabla.column("#0", width=40, anchor="center")

        self.tabla.pack(padx=20, pady=20)

        # ---------------BOTON DE REGISTRO--------------
        imagen_registro = Image.open(ruta(os.path.join("Icon", "boton-agregar.png")))
        # imagen_registro = Image.open(r"Icon\boton-agregar.png")
        icon_registro = ctk.CTkImage(
            light_image=imagen_registro, dark_image=imagen_registro, size=(30, 30)
        )

        self.button_registro = ctk.CTkButton(
            self.tabla_frame,
            text="Registrar",
            width=3,
            height=20,
            corner_radius=20,
            border_color="white",
            command=self.registrar_peso,
            image=icon_registro,
            compound="top",
        )
        self.button_registro.pack()

        # ----------- Frame Bottones de la Tabla-----------

        self.frame_buttons_tabla = ctk.CTkFrame(self.tabla_frame, width=400, height=150)
        self.frame_buttons_tabla.pack(expand=True)
        self.frame_buttons_tabla.pack_propagate(False)

        # ------------Botones de la tabla-------------------
        """
        Botones de la Tabla nuevo, eliminar etc
        """
        # --------ARREGLAR LOS BOTONES--------------------

        # +++++++++++++Boton guardar++++++++++++++
        # imagen_guardar = Image.open(r"Icon\guardar-el-archivo.png")
        imagen_guardar = Image.open(
            ruta(os.path.join("Icon", "guardar-el-archivo.png"))
        )
        icon_save = ctk.CTkImage(
            light_image=imagen_guardar, dark_image=imagen_guardar, size=(30, 30)
        )

        self.button_guardar = ctk.CTkButton(
            self.frame_buttons_tabla,
            text="Guardar",
            width=3,
            corner_radius=20,
            border_width=1,
            border_color="white",
            command=self.guardar_xlsx,
            image=icon_save,
            compound="top",
        )
        self.button_guardar.pack(side="left", expand=True)
        # +++++++++++++++++++++++++++++++++++++++++

        # -----------BOTON NUEVO-------------------
        imagen_new = Image.open(ruta(os.path.join("Icon", "anadir.png")))
        # imagen_new = Image.open(r"Icon\anadir.png")
        icon_new = ctk.CTkImage(
            light_image=imagen_new, dark_image=imagen_new, size=(30, 30)
        )
        self.button_nuevo = ctk.CTkButton(
            self.frame_buttons_tabla,
            text="Nuevo",
            width=3,
            height=20,
            corner_radius=20,
            border_width=1,
            border_color="white",
            command=self.limpiar,
            image=icon_new,
            compound="top",
        )
        self.button_nuevo.pack(side="left", expand=True)
        # -------------------------------------------------
        imagen_eliminar_ultimo = Image.open(ruta(os.path.join("Icon", "borrar.png")))
        # imagen_eliminar_ultimo = Image.open(r"Icon/borrar.png")
        icono_eliminar_ultimo = ctk.CTkImage(
            light_image=imagen_eliminar_ultimo,
            dark_image=imagen_eliminar_ultimo,
            size=(30, 30),
        )

        self.button_eliminar_ultimo = ctk.CTkButton(
            self.frame_buttons_tabla,
            text="Eliminar Ultimo",
            width=3,
            height=20,
            corner_radius=20,
            border_width=1,
            border_color="white",
            command=self.eliminar_ultimo,
            image=icono_eliminar_ultimo,
            compound="top",
        )
        self.button_eliminar_ultimo.pack(side="left", expand=True)
        # ******************************************************************

        # ------------------Frame PARA EL PUERTO-------------------------------

        # -----------------------Frame Principal-----------------------

        self.principal_frame = ctk.CTkFrame(self, width=50, height=60)
        self.principal_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.frame_peso = ctk.CTkFrame(self.principal_frame, width=300, height=120)
        self.frame_peso.pack(pady=40, padx=20, fill="both")
        self.frame_peso.pack_propagate(False)

        self.label_unidad = ctk.CTkLabel(
            self.frame_peso, text="kg", font=("Cambria", 50)
        )
        self.label_unidad.pack(side="right", padx=30)

        # ***********************************+

        self.frame_peso_al_momento = ctk.CTkFrame(
            self.frame_peso, width=400, height=150
        )
        self.frame_peso_al_momento.pack(pady=10, padx=10, fill="both", expand=True)
        self.frame_peso_al_momento.pack_propagate(False)

        self.frame_peso_al_momento.grid_columnconfigure(0, weight=1)
        self.frame_peso_al_momento.grid_columnconfigure(1, weight=3)
        self.frame_peso_al_momento.grid_rowconfigure((0, 1), weight=1)

        self.label_tara = ctk.CTkLabel(
            self.frame_peso_al_momento, text="", font=("Cambria", 15)
        )
        self.label_tara.grid(row=0, column=0, sticky="nw", padx=10, pady=10)

        self.label_neto = ctk.CTkLabel(
            self.frame_peso_al_momento, text="", font=("Cambria", 15)
        )
        self.label_neto.grid(row=1, column=0, sticky="sw", padx=10, pady=5)

        self.label_peso = ctk.CTkLabel(
            self.frame_peso_al_momento,
            text="",
            text_color="red",
            font=("Cambria", 90),
        )
        self.label_peso.grid(row=0, column=1, rowspan=2, sticky="nsew")

        # ----------- Frame de Botones Principales, tarar, zero etc

        self.frame_buttons_principales = ctk.CTkFrame(
            self.principal_frame, width=400, height=100
        )
        self.frame_buttons_principales.pack(pady=5)
        self.frame_buttons_principales.pack_propagate(False)

        # -----------BOTON DE TARAR--------------------
        imagen_tarar = Image.open(ruta(os.path.join("Icon", "aumentar.png")))
        # imagen_tarar = Image.open(r"Icon\aumentar.png")
        icon_tarar = ctk.CTkImage(
            light_image=imagen_tarar, dark_image=imagen_tarar, size=(30, 30)
        )

        self.button_tarar = ctk.CTkButton(
            self.frame_buttons_principales,
            text="Tara",
            width=3,
            height=20,
            corner_radius=40,
            border_color="white",
            border_width=1,
            command=self.shell.tara,
            image=icon_tarar,
            compound="top",
        )
        self.button_tarar.pack(side="left", expand=True)

        # ------------BOTON DE ZERO-------------
        imagen_zero = Image.open(ruta(os.path.join("Icon", "cero.png")))
        # imagen_zero = Image.open(r"Icon\cero.png")
        icon_zero = ctk.CTkImage(
            light_image=imagen_zero, dark_image=imagen_zero, size=(30, 30)
        )

        self.button_zero = ctk.CTkButton(
            self.frame_buttons_principales,
            text="Zero",
            width=3,
            height=20,
            corner_radius=40,
            border_color="white",
            border_width=1,
            command=self.shell.zero,
            image=icon_zero,
            compound="top",
        )
        self.button_zero.pack(side="left", expand=True)
        # -----------BOTON DE QUITAR TARA---------------

        imagen_q_tara = Image.open(ruta(os.path.join("Icon", "perdida-peso.png")))
        # imagen_q_tara = Image.open(r"Icon\perdida-peso.png")
        icon_q_tara = ctk.CTkImage(
            light_image=imagen_q_tara, dark_image=imagen_q_tara, size=(30, 30)
        )
        self.button_quitar_tara = ctk.CTkButton(
            self.frame_buttons_principales,
            text="Eliminar Tara",
            width=3,
            height=20,
            corner_radius=40,
            border_color="white",
            border_width=1,
            command=self.shell.quitar_tara,
            image=icon_q_tara,
            compound="top",
        )
        self.button_quitar_tara.pack(side="left", expand=True)

        # ------------FRAME PARA MOSTRAR EL PUERTO----------------

        self.frame_puerto = ctk.CTkFrame(self.principal_frame, width=300, height=50)
        self.frame_puerto.pack(side="bottom", pady=10)
        self.frame_puerto.pack_propagate(False)

        self.label_port = ctk.CTkLabel(
            self.frame_puerto, text="Puerto", font=("Cambria", 15)
        )
        self.label_port.pack(expand=True)

    def b_menus(self):
        """
        Crear los menus genericos
        """
        self.menu = CTkMenuBar(master=self)
        self.b1 = self.menu.add_cascade("Archivo")  # Menu de archivo
        self.b2 = self.menu.add_cascade("Puerto")  # Menu de puertos
        self.b3 = self.menu.add_cascade(
            "Catalogo", postcommand=self.ventana_calibracion
        )
        self.b4 = self.menu.add_cascade(
            "Repetibilidad", command=self.ventana_repetibilidad
        )

        # -------------Las opciones solo del menu archivo---------------
        self.dropdown_1 = CustomDropdownMenu(widget=self.b1)
        self.dropdown_1.add_option(
            option="Abrir",
            command=self.abrir_xlsx,
            icon=ruta(os.path.join("Icon", "carpeta-abierta.png")),
            # icon=r"Icon\carpeta-abierta.png",
            # icon=r"Icon\folder_open_24dp_E3E3E3_FILL0_wght400_GRAD0_opsz24.png",
        )
        self.dropdown_1.add_option(
            option="Guardar",
            command=self.guardar_xlsx,
            icon=ruta(os.path.join("Icon", "marcador.png")),
            # icon=r"Icon\marcador.png",
        )
        self.dropdown_1.add_option(
            option="Limpiar tabla",
            command=self.limpiar,
            icon=ruta(os.path.join("Icon", "cepillar.png")),
            # icon=r"Icon\cepillar.png",
        )
        self.dropdown_1.add_option(
            option="Salir",
            command=self.destroy,
            icon=ruta(os.path.join("Icon", "salida.png")),
        )
        self.actualizar_puerto()
        self.dropdown_puerto.add_option(
            option="Cerrar puerto",
            command=self.comunicacion.cerrar_puerto,
            icon=ruta(os.path.join("Icon", "cable-roto.png")),
            # icon=r"Icon\desenchufado.png",
        )

        # self.dropdown_3 = CustomDropdownMenu(widget=self.b3)
        # self.dropdown_3.add_option(
        #     option="Ajustes",
        #     command=self.ventana_calibracion,
        #     icon=ruta(os.path.join("Icon", "negocio.png")),
        # )

    def actualizar_puerto(self):
        self.dropdown_puerto = CustomDropdownMenu(widget=self.b2)
        puertos = serial.tools.list_ports.comports()
        if not puertos:
            self.dropdown_puerto.add_option(option="Sin puertos")
        else:
            for p in puertos:
                self.dropdown_puerto.add_option(
                    option=p.device,
                    command=lambda puerto=p.device: self.seleccionar_puerto(puerto),
                    icon=ruta(os.path.join("Icon", "vga_cable.png")),
                    # icon=r"Icon\vga_cable.png",
                )

    def seleccionar_puerto(self, puerto_elegido):
        print(f"puerto elegido {puerto_elegido}")
        self.comunicacion.port = puerto_elegido

        self.label_port.configure(
            text=f"Puerto: {self.comunicacion.port} | {self.comunicacion.baud}"
        )
        if self.comunicacion.conexion():
            # inf = f"conexion: {self.comunicacion.port} | baud: {self.comunicacion.baud}"
            self.label_port.configure(text_color="#4E915D")
            print("conectado")
        else:
            self.label_port.configure(text_color="#823131")
            print("port fallo")

    def conectar_puerto(self, info):
        """
        Conexion del puerto al seleccionar
        """
        try:
            if hasattr(self, "comunicacion") and self.comunicacion:
                self.comunicacion.cerrar_puerto()

            self.comunicacion = Comunication(port=info["Dispositivo"], baud=9600)
            self.comunicacion.conexion()
            if self.comunicacion.conectado:
                self.shell = Shell(self.comunicacion)

                CTkMessagebox(title="Conexion", message="Conectado", icon="check")
            else:
                CTkMessagebox(title="Conexion", message="Salio mal", icon="warning")

        except Exception as e:
            CTkMessagebox(
                title="Error conexion", message="Error al conectar", icon="warning"
            )

    def registrar_peso(self):
        """
        Registra el peso con el button siempre que la lectura sea estable
        """
        lecture = self.shell.peso_estable()
        if len(lecture) >= 3:
            peso_actual = float(lecture[2])
            incremento = peso_actual - self.ultimo_peso

            fecha = datetime.now().strftime("%d/%m/%Y")
            hora = datetime.now().strftime("%H:%M:%S")
            id_fila = len(self.tabla.get_children()) + 1
            self.tabla.insert(
                "",
                "end",
                text=str(id_fila),
                values=(incremento, peso_actual, fecha, hora),
            )
            # self.table.insert('', 'end', values=(id_incremento,incremento, fecha, hora))
            self.ultimo_peso = peso_actual
            self.tabla.yview_moveto(1)

    def guardar_xlsx(self):
        """
        Guarda los registros de la tabla en archivo excel
        """

        if not self.tabla.get_children():
            CTkMessagebox(title="Erro", message="Sin datos", icon="warning")
            return

        archivo = asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivo csv", "*.csv")],
            title="Pesossss",
        )
        if archivo:
            try:
                with open(archivo, mode="w", newline="", encoding="utf-8") as f:
                    escritor = csv.writer(f)
                    escritor.writerow(["Incremento", "Indicación", "Fecha", "Hora"])

                    for i in self.tabla.get_children():
                        val = self.tabla.item(i)["values"]
                        escritor.writerow(val)
                CTkMessagebox(title="Guardar", message="Guardado", icon="check")
            except Exception as e:
                print("Error")

    def abrir_xlsx(self):
        """
        Abre archivo con extension excel dentro de la interfaz
        """
        # if not self.table.get_children():
        #     CTkMessagebox(title='Erro',message='Sin datos', icon='warning' )
        #     return

        archivo = askopenfilename(
            filetypes=[("Archivo csv", "*.csv")], title="Pesossss"
        )
        if archivo:
            try:
                # self.limpiar()
                with open(archivo, mode="r", encoding="utf-8") as f:
                    lecture = csv.DictReader(f)
                    # escritor.writerow(['Incremento', 'Peso', 'Fecha','Hora'])

                    for i in lecture:
                        inc = i["Incremento"]
                        peso = i["Indicación"]
                        fecha = i["Fecha"]
                        hora = i["Hora"]

                        id_i = len(self.tabla.get_children()) + 1
                        self.tabla.insert(
                            "", "end", text=str(id_i), values=(inc, peso, fecha, hora)
                        )

                        items = self.tabla.get_children()
                        if items:
                            last_i = items[-1]
                            val = self.tabla.item(last_i)["values"]
                            self.ultimo_peso = float(val[1])
                    CTkMessagebox(title="Open", message="Exito", icon="check")
                    self.tabla.yview_moveto(1)
            except Exception as e:
                CTkMessagebox(title="Open", message="Warning", icon="warning")

    def eliminar_ultimo(self):
        """
        Elimina el ultimo elemento de la tabla
        """
        it = self.tabla.get_children()
        if not self.tabla.get_children():
            CTkMessagebox(title="Eliminar ultimo", message="Sin elementos")
            return
        i_last = it[-1]
        self.tabla.delete(i_last)

        new_i = self.tabla.get_children()

        if new_i:
            new_last = new_i[-1]
            val = self.tabla.item(new_last)["values"]
            self.ultimo_peso = float(val[1])
        else:
            self.ultimo_peso = 0

        CTkMessagebox(title="Eliminar ultimo", message="Eliminado", icon="check")

    def limpiar(self):
        """
        Elimina/limpia toda la tabla con los elementos
        """
        # self.tabla.delete(*self.tabla.get_children()) + 1
        for i in self.tabla.get_children():
            self.tabla.delete(i)

        self.ultimo_peso = 0

    def peso_al_momento(self):
        """
        Es la actualizacion del peso en tiempo real que se ve en la interfaz
        """
        lecture = self.shell.peso_instantaneo()

        tara_lecture = self.shell.obtener_tara()

        if len(lecture) >= 3:
            try:
                peso_r = int(float(lecture[2]))
                self.label_peso.configure(text=str(peso_r))
                self.label_neto.configure(
                    text=f"Neto: {peso_r} kg", text_color="#FF6EDE"
                )
            except ValueError:
                pass

        # if len(lecture) >= 3:
        #     bruto = int(lecture[2])
        #     try:
        #         p_tara = int(tara_lecture.split()[0])
        #     except:
        #         p_tara = 0

        #     neto = bruto - p_tara
        #     self.label_neto.configure(text=f"N:{neto} kg", text_color="#FF6EDE")

        if "+" in lecture:
            self.label_peso.configure(text="----", text_color="red")

            if not self.alerta_peso_maximo:
                self.alerta_peso_maximo = True
                CTkMessagebox(
                    title="Maximo peso",
                    message="Sobrecarga. Retire peso",
                    icon="cancel",
                    option_1="aceptar",
                )
        elif len(lecture) >= 3:
            self.alerta_peso_maximo = False

        if len(lecture) >= 3:
            wg = lecture[2]
            self.label_peso.configure(text=wg)

            if lecture[1] == "S":
                self.label_peso.configure(text_color="#158A30")
            else:
                self.label_peso.configure(text_color="#B52C19")

            valor_tara = self.shell.obtener_tara()

            self.label_tara.configure(text=f"Tara: {valor_tara}", text_color="#FF6021")
        self.after(100, self.peso_al_momento)

    def apli_cal_cero(self):
        res = self.shell.cali_cero()
        if res == "Z A":
            CTkMessagebox(title="Calibracion", message="Cero estable", icon="check")

        elif res == "Z I":
            CTkMessagebox(title="Calibracion", message="Inestable", icon="cancel")
        else:
            CTkMessagebox(
                title="Calibracion", message=f"No calibro {res}", icon="warning"
            )

    def ventana_calibracion(self):
        if (
            not hasattr(self, "calibracion_window")
            or not self.calibracion_window.winfo_exists()
        ):
            self.calibracion_window = ventana_calibracion(self, self.shell)
        else:
            self.calibracion_window.focus()

    def ventana_repetibilidad(self):
        if (
            not hasattr(self, "repetibilidad_window")
            or not self.repetibilidad_window.winfo_exists()
        ):
            self.repetibilidad_window = repetibilidad(self, self.shell)
        else:
            self.repetibilidad_window.focus()
