#Programa para barómetro H008
#Este programa se elabora por segunda vez para poder obtener los valores mas legíbles y de manera gráfica

import serial, threading, time, collections
import sys
from tkinter import *
from datetime import date, datetime
import math
import tkinter
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import numpy as np

global Lectura
#from PIL import ImageTk, Image
Lectura = 0
Lecturah = 0

Samples = 100 #Muestras
data = collections.deque([0]*Samples, maxlen=Samples) #Vector de muestras
sampleTime = 1000 #Tiempo de muestreo

#Límites de ejes
xmin = 0
xmax = Samples
ymin = 0
ymax = 30

higrometro=serial.Serial('COM8',9600)
barometro=serial.Serial('COM5',9600,parity='N')



def getSerialData(self, Lecturah,Samples, lines, lineValueText):
    data.append(str(Lecturah[21:26]))
    lines.set_data(range(Samples),data)
    #lineValueText.set_text(lineLabel +' = '+str(round(Lecturah[21:26]),2))

def grafica1():
    fig = plt.figure(figsize=(13,6))
    ax = plt.axes(xlim=(xmin,xmax), ylim=(ymin,ymax))
    plt.title('Temperatura')
    ax.set_xlabel('Muestras')
    ax.set_ylabel('°C')

    lineLabel = 'Temperatura'
    lines = ax.plot([], [], label='Temperatura')[0]
    lineValueText = ax.text(0.85, 0.95, ' ', transform=ax.transAxes)

    #data.append(str(Lecturah[21:26]))
    #lines.set_data(range(Samples),data)
    #lineValueText.set_text(lineLabel+' = '+str(round(value,2)))    

    anim = animation.FuncAnimation(fig, getSerialData, fargs=(Samples, round(float(Lecturah[21:25])), lineValueText), interval=sampleTime)
    plt.show()


#Iniciamos la programación orientada a objetos
def read_bar():
    barometro.flushInput()
    global Lectura
    barometro.write(b':clr\n')
    barometro.write(b':SENS')
    barometro.write(b':PRESsure?')
    barometro.write(b':SYST')
    barometro.write(b'?')
    LecturaBar=barometro.readline()
    Lectura=str(float(LecturaBar[11:].decode('ascii').rstrip('\n').rstrip('\r')))
    #print(str(datetime.now())+' '+Lectura+' Pa')
    v.label.config(text='%.1f'%(float(Lectura))+' Pa')
    v.ventana.update()
    v.ventana.after(10,read_bar)


def read_higro():
    higrometro.flushInput()
    global Lecturah, LecturaT2, LecturaT1
    global dato
    dato=higrometro.readline(56)
    Lecturah=dato.decode('ascii').replace(',',' ').replace('C','°C')
    LecturaT1=str(Lecturah[21:29])
    LecturaH1=str(Lecturah[31:39])
    LecturaT2=str(Lecturah[40:49])
    LecturaH2=str(Lecturah[51:59])
    if int(Lecturah.find('%'))<=33 :
        time.sleep(1)
        higrometro.close()
        time.sleep(1)
        higrometro.open()
    else:
        v.label5.config(text='Sensor 1:       '+str(LecturaT1)+'                  '+str(LecturaH1))
        v.label6.config(text='Sensor 2:       '+str(LecturaT2)+'                  '+str(LecturaH2))
        #v.label3.config(text=str(Lecturah[20:]))
        v.ventana.update()
        v.ventana.after(10,read_higro)
        

def calc_dens():
    global res, res1
    if int(Lecturah.find('%'))<=33 :
        res = 0
        res1 = 0
        return Lectura.set(res),Lecturah[31:36].set(res1),Lecturah[21:26].set(res),Lecturah[51:56].set(res),Lecturah[40:45].set(res1)
    else:
        res=(0.34848*float('%.1f'%(float(Lectura)/100))-0.009*(float(Lecturah[31:36]))*math.exp(0.061*(float(Lecturah[21:26]))))/(273.5+(float(Lecturah[21:26])))
        res1=(0.34848*float('%.1f'%(float(Lectura)/100))-0.009*(float(Lecturah[51:56]))*math.exp(0.061*(float(Lecturah[40:45]))))/(273.5+(float(Lecturah[40:45])))
        #print(str('%.1f'%(float(Lectura))))
        #print(Lecturah[31:36])
        #print(Lecturah[21:26])
        #res=float(Lecturah[21:26])+float(Lecturah[51:56])
        v.label9.config(text='Densidad del aire desde sensor 1:           '+ str(float('%.4f'%(res)))+ '  kg/m3')
        v.label10.config(text='Densidad del aire desde sensor 2:           '+ str(float('%.4f'%(res1)))+ '  kg/m3')
        v.ventana.update()
        v.ventana.after(10,calc_dens)
        #time.sleep(1)
    

def saveData():
    while(1):
        time.sleep(1)
        file = open('registro_ambiente.txt','a')
        file.write(str(datetime.now())+' '+str(Lecturah[20:55])+'  '+str('%.1f'%(float(Lectura)))+'  Pa   '+ 'rho a 1: '+str(float('%.4f'%(res)))+'    '+'rho a 2: '+str(float('%.4f'%(res1)))+' \n')
        time.sleep(59)



class GUI():
    def __init__(self):
        pass
    def CrearVentana(self):
        self.ventana=Tk()
        self.ventana.title('Laboratorio Grandes Masas')
        self.ventana.geometry('900x450')
        self.ventana.resizable(False,False)
        self.ventana.attributes('-transparentcolor', 'wheat1')
        self.ventana.config(bg='khaki3')
        #image = Image.open('logo_CENAM.png')
        #image = image.resize((50,50), Image.LANCZOS)
        #imagen = ImageTk.PhotoImage(image)
        #self.label3=Label(self.ventana, image=imagen, bd=0).pack()
        self.label=Label(self.ventana,text="",bg='khaki3',font=('Arial',16))
        self.label.place(x=115, y=80)
        self.label2=Label(self.ventana,text='Presión (DPI142):',fg='blue',bg='khaki3',font=('Arial',14))
        self.label2.place(x=115, y=50)
        self.b1=Button(self.ventana,text='Cerrar',command=self.cerrar,bg='khaki3',font=('bold',8))
        self.b1.place(x=5, y=10)
        self.label3=Label(self.ventana,text='Temperatura'+'       '+'Humedad Relativa',bg='khaki3',font=('Arial',16))
        self.label3.place(x=205,y=155)
        self.label4=Label(self.ventana,text='Temperatura y Humedad Relativa (Fluke):',fg='blue',bg='khaki3',font=('Arial',14))
        self.label4.place(x=115, y=125)
        self.label5=Label(self.ventana,text='',bg='khaki3',font=('Arial',14))
        self.label5.place(x=115, y=185)
        self.label6=Label(self.ventana,text='',bg='khaki3',font=('Arial',14))
        self.label6.place(x=115, y=215)
        self.label7=Label(self.ventana,text='H008: Laboratorio Grandes Masas',bg='khaki3',font=('Arial',17))
        self.label7.place(x=110, y=10)
        self.label8=Label(self.ventana,text='Densidad del aire',fg='blue',bg='khaki3',font=('Arial',14))
        self.label8.place(x=110, y=255)
        self.label9=Label(self.ventana,text='',bg='khaki3',font=('Arial',14))
        self.label9.place(x=115, y=285)
        self.label10=Label(self.ventana,text='',bg='khaki3',font=('Arial',14))
        self.label10.place(x=115, y=315)
        
        bar=threading.Thread(target=read_bar)
        higro=threading.Thread(target=read_higro)
        calc=threading.Thread(target=calc_dens)
        gf=threading.Thread(target=grafica1)
        bar.run()
        higro.run()
        calc.run()
        gf.run()

    def cerrar(self):
        self.ventana.destroy()
        Save.join()
        bar.join()
        higro.join()
        calc.join()
        gf.join()
       

def comenzar():
    v.CrearVentana()        

v=GUI()
hiloStart=threading.Thread(target=comenzar)
hiloStart.run()
Save=threading.Thread(target=saveData)
Save.start()



