#****************************************************************************************************************************************************************************************************************************************************************
# Nombre: Modificar.Py                                                                                                                                                                                                                                          *       
# Descripcion: Aplicación para modificar información de equipos en listado de equipos                                                                                                                                                                           *
#               El funcionamiento es simple, de la busqueda de un equipo en la app "Busqueda.py* salta la info y podemos modificar los mismos parametros sin tener que buscar nuevamente.
# #**************************************************************************************************************************************************************************************************************************************************************
import cmd
from dis import dis
import os as os
import sys as sys
from sys import *
from os import replace, system as system
import datetime,time
#****************************************************************************************************************************************************************************************************************************************************************
import tkinter as tk
from tkinter import *
from tkinter import Entry, Grid, Image, StringVar, Text, Variable, messagebox, ttk, scrolledtext, simpledialog, tix, font, commondialog
from tkinter.ttk import Progressbar, setup_master,Sizegrip,Entry
from tkinter.tix import STATUS, LabelEntry,LabelFrame,Meter, ButtonBox,PhotoImage,ComboBox
from tkinter.constants import *
from tokenize import Double, String
from turtle import color, delay, title
from typing import Any
#****************************************************************************************************************************************************************************************************************************************************************
import openpyxl
from openpyxl import Workbook,load_workbook
#****************************************************************************************************************************************************************************************************************************************************************
from Variables import *
#****************************************************************************************************************************************************************************************************************************************************************
#                                                                                                               Declaración de Variables                                                                                                                        *
#****************************************************************************************************************************************************************************************************************************************************************
new_user=os.environ['USERPROFILE']                                                                                                  # Identifico el "UserProfile" de la pc para poder encontrar las carpetas instaladas en el Setup.py                          *
user=new_user.replace("\\","/")                                                                                                     # Reemplazo "\" por "/" dado a que no reconocen la ruta en Python                                                           *

#path=open(new_user + '/Desktop/Codigo/Paths.txt','r')                                                                              # Busco el archivo donde se encuentran las rutas preestablecidas para encontrar los archivos. (Las cuales se pueden         *
path=open(user+'/Desktop/Soft/Path.txt')                                                                              
lineas=path.readlines()                                                                                                             # modificar en caso que asi se quiera)                                                                                      *

log=open(user  + (str(lineas[1])[:-1]),mode="a")                                                                                   # Ruta de Archivo donde se encuentran los Logs de eventos                                                                   *
#****************************************************************************************************************************************************************************************************************************************************************
#   Variables para uso del programa                                                                                                                                                                                                                             *
#****************************************************************************************************************************************************************************************************************************************************************
#****************************************************************************************************************************************************************************************************************************************************************
#****************************************************************************************************************************************************************************************************************************************************************
# Defina Interfaz Grafica                                                                                                                                                                                                                                       *
#****************************************************************************************************************************************************************************************************************************************************************
#****************************************************************************************************************************************************************************************************************************************************************
Mymod = Tk()                                                                                                                        # MyApp es el nombre de la planilla a visualizar                                                                            *                                                                                                             
Mymod.title("Modificar")                                                                                                             # Defino el titulo del programa                                                                                             *
H=500
W=1000
Mymod.minsize(W,H)                                                                                                                  # Defino el Tamaño minimo e inicial de plantilla                                                                            *
Mymod.frame()
Mymod.resizable(False,False)
Mymod.config(background=Fondo)                                                                                                      # Defino color de Fondo de pantalla "Myapp"                                                                                 *
#****************************************************************************************************************************************************************************************************************************************************************
#   Lineas de comando para poder centrar la ventana en la mitad de la pantalla (En teoria)
#****************************************************************************************************************************************************************************************************************************************************************
screen_width =  Mymod.winfo_screenwidth()
screen_height = Mymod.winfo_screenheight()
x_cordinate =   int((screen_width/2) - (W/2))
y_cordinate =   int((screen_height/2) - (H/2)-25)
Mymod.geometry("{}x{}+{}+{}".format(W,H, x_cordinate, y_cordinate))
#****************************************************************************************************************************************************************************************************************************************************************
MyFondo=    PhotoImage(file=(user   + (str(lineas[4]))[:-1]))                                                                     # Variable para imagen del Logo del Icono a usar                                                                            *
lbl_lable=  ttk.Label(Mymod,image=MyFondo,border=0).place(x=-10,y=-10)                                                            # Se define en Lable para poder utilizar "mylogo" como fondo de ventana.                                                    *

#****************************************************************************************************************************************************************************************************************************************************************
# Definimos Variables para respuestas                                                                                                                                                                                                                           *
#****************************************************************************************************************************************************************************************************************************************************************
res_dispositivo=    StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Dispositivo"                                                                          *
res_ubi=            StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Ubicacion"                                                                            *
res_equipo=         StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Equipo"                                                                               *
res_nombre=         StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Nombre"                                                                               *
res_marca=          StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Marca"                                                                                *
res_modelo=         StringVar()
res_ip=             StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Ip"                                                                                   *
res_serial=         StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "serial"                                                                               *
res_usuario=        StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Usuario"                                                                              *
res_password=       StringVar()                                                                                                 # Variable para respuesta de Etiqueta de "Contraseña"                                                                           *
res_server=         StringVar()
#****************************************************************************************************************************************************************************************************************************************************************
mod_ubi=            StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "Ubicacion"                                                                            *
mod_equipo=         str()                                                                                                       # Variable para Modificar de Etiqueta de "Equipo"                                                                               *
mod_nombre=         StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "Nombre"                                                                               *
mod_marca=          StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "Marca"                                                                                *
mod_ip=             StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "Ip"                                                                                   *
mod_serial=         StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "serial"                                                                               *
mod_usuario=        StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "Usuario"                                                                              *
mod_password=       StringVar()                                                                                                 # Variable para Modificar de Etiqueta de "Contraseña"                                                                           *
mod_server=         StringVar()
#****************************************************************************************************************************************************************************************************************************************************************
simb_disp=          IntVar()                                                                                                    # Validador de Check Boton de "Dispositivo"                                                                                     *
simb_ubi=           IntVar()                                                                                                    # Validador de Check Boton de "Ubicacion"                                                                                       *
simb_equipo=        IntVar()                                                                                                    # Validador de Check Boton de "Equipo"                                                                                          *
simb_nombre=        IntVar()                                                                                                    # Validador de Check Boton de "Nombre"                                                                                          *
simb_marca=         IntVar()                                                                                                    # Validador de Check Boton de "Marca"                                                                                           *
simb_modelo=        IntVar()                                                                                                    # Valodador de Check Boton de "Modelo"                                                                                          *
simb_ip=            IntVar()                                                                                                    # Validador de Check Boton de "Ip"                                                                                              *
simb_serial=        IntVar()                                                                                                    # Validador de Check Boton de "serial"                                                                                          *
simb_user=          IntVar()                                                                                                    # Validador de Check Boton de "Usuario"                                                                                         *
simb_pass=          IntVar()                                                                                                    # Validador de Check Boton de "Contraseña"                                                                                      *
simb_server=        IntVar()
#****************************************************************************************************************************************************************************************************************************************************************
# Labels para mostrar información                                                                                                                                                                                                                               *    
#****************************************************************************************************************************************************************************************************************************************************************
lbl_Ubicacion=      tk.Label(Mymod,text="Ubicacion =>"          ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*2)
lbl_equipo=         tk.Label(Mymod,text="Equipo =>"             ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*3)
lbl_nombre=         tk.Label(Mymod,text="Nombre =>"             ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*4)
lbl_marca=          tk.Label(Mymod,text="Marca =>"              ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*5)
lbl_modelo=         tk.Label(Mymod,text="Modelo =>"             ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*6)
lbl_Serial=         tk.Label(Mymod,text="Serial =>"             ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*7)
lbl_ip=             tk.Label(Mymod,text="IP =>"                 ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*8)
lbl_usuario=        tk.Label(Mymod,text="Usuario =>"            ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*9)
lbl_Password=       tk.Label(Mymod,text="Password =>"           ,background=Fondo,font=("Arial",12),width=15)           .place(x=10,y=Renglon*10)

#****************************************************************************************************************************************************************************************************************************************************************                                                                                                                
lbl_res_ubicacion=  tk.Label(Mymod,textvariable=res_ubi         ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*2)
lbl_res_equipo=     tk.Label(Mymod,textvariable=res_equipo      ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*3)
lbl_res_nombre=     tk.Label(Mymod,textvariable=res_nombre      ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*4)
lbl_res_marca=      tk.Label(Mymod,textvariable=res_marca       ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*5)
lbl_res_Modelo=     tk.Label(Mymod,textvariable=res_modelo      ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*6)
lbl_res_Serial=     tk.Label(Mymod,textvariable=res_serial      ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*7)
lbl_res_Ip=         tk.Label(Mymod,textvariable=res_ip          ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*8)
lbl_res_Usuario=    tk.Label(Mymod,textvariable=res_usuario     ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*9)
lbl_res_Password=   tk.Label(Mymod,textvariable=res_password    ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*10)
lbl_res_Server=     tk.Label(Mymod,textvariable=res_server      ,font=Fuente,background="#AAAAAA",width=Ancho)    .place(x=Posicion,y=Renglon*11)
#****************************************************************************************************************************************************************************************************************************************************************                                                                                                                
modifi=450
#****************************************************************************************************************************************************************************************************************************************************************
E_mod_ubicacion=    ttk.Entry(Mymod,textvariable=res_ubi        ,font=Fuente,foreground="#FF0000",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*2)         # Defino el renglon de entrada de datos para comparar y Se define               *
E_mod_equipo=       ttk.Entry(Mymod,textvariable=res_equipo     ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*3)         #posición del renglon de "Entrada"                                              *
E_mod_nombre=       ttk.Entry(Mymod,textvariable=res_nombre     ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*4)
E_mod_marca=        ttk.Entry(Mymod,textvariable=res_marca      ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*5)
E_mod_modelo=       ttk.Entry(Mymod,textvariable=res_modelo     ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*6)
E_mod_Serial=       ttk.Entry(Mymod,textvariable=res_serial     ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*7)
E_mod_Ip=           ttk.Entry(Mymod,textvariable=res_ip         ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*8)
E_mod_Usuario=      ttk.Entry(Mymod,textvariable=res_usuario    ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*9)
E_mod_Password=     ttk.Entry(Mymod,textvariable=res_password   ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*10)
E_mod_Server=       ttk.Entry(Mymod,textvariable=res_server     ,font=Fuente,background="#AAAAAA",width=Ancho,justify='center')   .place(x=Posicion+modifi,y=Renglon*11)
#****************************************************************************************************************************************************************************************************************************************************************
#****************************************************************************************************************************************************************************************************************************************************************
# Funciones del Programa a utilizar                                                                                                                                                                                                                             *
#****************************************************************************************************************************************************************************************************************************************************************   

#****************************************************************************************************************************************************************************************************************************************************************                        
def Siguiente():
    
    global num,dispositivo,Wtree                                                                                                    # Tomo las Variables Globales y le permito el uso dentro del subprograma                                                    *
    
    num+=1                                                                                                                          # Aumento la variable "num"                                                                                                 *

    wb=openpyxl.load_workbook(str(lineas[0])[:-1])                                                                                  # Genero el vinculo y abro el Excel que esta definido dentro del String dentro del Path                                     *     
    ws=wb[Wtree]                                                                                                                    # Abro la Hoja (ws) CCTV para poder trabajar con esos datos
    dispositivo=ws.cell(row=num,column=4)                                                                                           # Defino la variable para poder compara con la variable "Entrada" teniendo en cuenta los parametros "num" como renglones    *
                                                                                                                                    #y como columna una sola definitiva para poder buscar solamente por nombre de equipo                                        *   
   
    if dispositivo.value=="":                                                                                                       # Verifico que la variable no este vacia                                                                                    *
        messagebox.showerror(message="Final de listado sin datos similares")                                                        # En caso de que se encuentra vacia, mostramos mensaje indicando que llego al final del listado de equipos                  *
    else:
        Imprimir()                                                                                                                  # En caso de que tengamos información dentro de ese renglon, pasamos los datos al sub Imprimi()                             *
#****************************************************************************************************************************************************************************************************************************************************************    
def Anterior():
    
    global num, dispositivo                                                                                                          # Tomo las Variables Globales y le permito el uso dentro del subprograma                                                    *
    
    if num>0:
        num-=1
        Imprimir()
    else:
        messagebox.showerror(message="No se encuentra valor menor")    
        
#****************************************************************************************************************************************************************************************************************************************************************
def m_ubi():
    global num,limite
    global dispositivo
        
    wb=openpyxl.load_workbook(str(lineas[0])[:-1])                                                                                  # Corresponde a la ruta en caso de tenerlo en la pc en forma local                                                          *
    ws=wb["CCTV"]

#****************************************************************************************************************************************************************************************************************************************************************
def Modific():
    
    global num,limite
    global dispositivo
        
    wb=openpyxl.load_workbook(str(lineas[0])[:-1])                                                                                  # Corresponde a la ruta en caso de tenerlo en la pc en forma local                                                          *
    ws=wb[Wtree]
    
    dispositivo=    ws.cell(row=num,column=4)
    nombre=         ws.cell(row=num,column=4)
    equipo=         ws.cell(row=num,column=2)
    ubicacion=      ws.cell(row=num,column=3)
    marca=          ws.cell(row=num,column=6)
    modelo=         ws.cell(row=num,column=7)
    ip=             ws.cell(row=num,column=11)
    serial=         ws.cell(row=num,column=8)
    usuario=        ws.cell(row=num,column=12)
    password=       ws.cell(row=num,column=13)
    server=         ws.cell(row=num,column=14)    
    
    equipo.value=   res_equipo  .get().upper()
    nombre.value=   res_nombre  .get().upper()
    ubicacion.value=res_ubi     .get().upper()
    marca.value=    res_marca   .get().upper()
    modelo.value=   res_modelo  .get().upper()
    ip.value=       res_ip      .get().upper()
    serial.value=   res_serial  .get().upper()
    usuario.value=  res_usuario .get().upper()
    password.value= res_password.get().upper()
    server.value=   res_server  .get().upper()
    
    wb.save(str(lineas[0])[:-1])
    
#****************************************************************************************************************************************************************************************************************************************************************        
def Imprimir():    
    
    global Dato,dispositivo, num, Wtree
    #from buscador import
    
    #dWtree="CCTV"
    print (Wtree)
    wb=openpyxl.load_workbook(str(lineas[0])[:-1])
    ws=wb[Wtree]
    
    Mensaje(str(Hini)+" => "+user[9:30] + " esta buscado el dato " +Dato+" que esta es el registro numero => "+str(num)+"\n")
    import WLog 
    
    dispositivo=    ws.cell(row=num,column=4)
    nombre=         ws.cell(row=num,column=4)
    equipo=         ws.cell(row=num,column=2)
    ubicacion=      ws.cell(row=num,column=3)
    marca=          ws.cell(row=num,column=6)
    modelo=         ws.cell(row=num,column=7)
    ip=             ws.cell(row=num,column=11)
    serial=         ws.cell(row=num,column=8)
    usuario=        ws.cell(row=num,column=12)
    password=       ws.cell(row=num,column=13)
    server=         ws.cell(row=num,column=14)    
           
    res_nombre  .set    (nombre.value)
    res_equipo  .set    (equipo.value)
    res_ubi     .set    (str(ubicacion.value))
    res_marca   .set    (marca.value)
    res_modelo  .set    (modelo.value)
    res_ip      .set    (ip.value)
    res_serial  .set    (serial.value)
    res_usuario .set    (usuario.value)
    res_password.set    (password.value)
    res_server  .set    (server.value)
    
#****************************************************************************************************************************************************************************************************************************************************************
def Conectar():

    wb=openpyxl.load_workbook(str(lineas[0])[:-1])
    ws=wb["CCTV"]
    server=ws.cell(row=num,column=14) 
    servidor='cmd /k "mstsc -v ' + server.value + ':4489'
    print(servidor)
    Salir()
    os.system(servidor)
    
    return 
#****************************************************************************************************************************************************************************************************************************************************************    
def Salir():
    
    Mensaje(str(Hini)+" => Se cierra aplicación "+user[9:]+"\n")
    import WLog
    Mymod.destroy()
    import buscador
#****************************************************************************************************************************************************************************************************************************************************************    
def Destroy():
    Mymod.destroy()
#****************************************************************************************************************************************************************************************************************************************************************
# Definimos los botoes a ver en la ventana inicial                                                                                                                                                                                                              *
#***************************************************************************************************************************************************************************************************************************************************************
altura=400
boton=      tk.Button(Mymod,text="Modificar",activebackground="#ABABAB",background="#838383",command=Modific,width=11,state='active')       .place(x=680,y=altura)                                     # Creo Boton "planilla" para procesar las plantillas Requeridas para informe.   *
salir=      tk.Button(Mymod,text="Salir"    ,activebackground="#BABABA",background="#838383",command=Salir,justify='center',width=23)       .place(x=790,y=altura+50)                                                       # Creo un Boton para cerrar la aplicación                                       *
bsiguiente= tk.Button(Mymod,text="Siguiente",activebackground="#ABABAB",background="#838383",command=Siguiente,width=11,state='active')     .place(x=780,y=altura)
banterios=  tk.Button(Mymod,text="Previo"   ,activebackground="#ABABAB",background="#838383",command=Anterior,width=11,state='active')      .place(x=880,y=altura)
#****************************************************************************************************************************************************************************************************************************************************************                     
imagen=     PhotoImage(file=(user+"/Desktop/Soft/flecha.png"))
#****************************************************************************************************************************************************************************************************************************************************************
#chk_simb_disp= tk.Button(myapp,activebackground=Fondo,background="#AAAAAA",image=imagen,height=bh).place(x=posicion+350,y=renglon*2)                                            # Boton para habilitar el cambio de Texto en Listado (Dispositivo)
chk_simb_ubi=   tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh,command=mod_ubi)                     .place(x=Posicion+350,y=Renglon*2)                                # Boton para habilitar el cambio de Texto en Listado (Ubicación)
chk_simb_equipo=tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_equipo)           .place(x=Posicion+350,y=Renglon*3)                  # Boton para habilitar el cambio de Texto en Listado (Equipo)
chk_simb_nombre=tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_nombre)           .place(x=Posicion+350,y=Renglon*4)                  # Boton para habilitar el cambio de Texto en Listado (Nombre)
chk_simb_marca= tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_marca)            .place(x=Posicion+350,y=Renglon*5)                    # Boton para habilitar el cambio de Texto en Listado (Marca)
chk_simb_modelo=tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_modelo)           .place(x=Posicion+350,y=Renglon*6)                  # Boton para habilitar el cambio de Texto en Listado (Modelo)
chk_simb_ip=    tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_ip)               .place(x=Posicion+350,y=Renglon*7)                          # Boton para habilitar el cambio de Texto en Listado (Ip)
chk_simb_serial=tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_serial)           .place(x=Posicion+350,y=Renglon*8)                  # Boton para habilitar el cambio de Texto en Listado (Serial)
chk_simb_user=  tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_user)             .place(x=Posicion+350,y=Renglon*9)                     # Boton para habilitar el cambio de Texto en Listado (Usuario)
chk_simb_pass=  tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,height=Bh, textvariable=simb_pass)             .place(x=Posicion+350,y=Renglon*10)                     # Boton para habilitar el cambio de Texto en Listado (Password)
chk_simb_server=tk.Button(Mymod,activebackground=Fondo,background="#AAAAAA",image=imagen,           textvariable=simb_server)           .place(x=Posicion+350,y=Renglon*11)                           # Boton para habilitar el cambio de Texto en Listado (Servidor)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
#****************************************************************************************************************************************************************************************************************************************************************
Imprimir()
#****************************************************************************************************************************************************************************************************************************************************************
log.close()
Mymod.mainloop()
#****************************************************************************************************************************************************************************************************************************************************************


