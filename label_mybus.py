import sys as sys

import tkinter as tk
from tkinter import *
from tkinter import Entry, Grid, Image, StringVar, Text, Variable, PhotoImage
from tkinter.constants import *
from tkinter.tix import *
from tkinter.ttk import *

from buscador   import *
from Variables  import *


def Labels ():
    
    Myimg=      PhotoImage(file=(user   + (str(lineas[3]))[:-1]))                                                                       # Variable para imagen del Boton de Busqueda. Se define la ruta en el programa.                                             *
    Mylogo=     PhotoImage(file=(user   + (str(lineas[4]))[:-1]))                                                                       # Variable para imagen del Logo del Icono a usar                                                                            *
    lbl_lable=  Label(mybus,image=Mylogo,border=0).place(x=-10,y=-10)                                                                   # Se define en Lable para poder utilizar "mylogo" como fondo de ventana.                                                    *

return