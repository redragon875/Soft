
#****************************************************************************************************************************************************************************************************************************************************************
import tkinter as tk
from tkinter import  messagebox
#****************************************************************************************************************************************************************************************************************************************************************
from Variables import *
from buscador import wbs


if wbs=="":
    messagebox.showerror(title="Error de Busqueda",message="Falta seleccionar dentro del arbol la subcategoria")                        # Comparo el Valor obtenido de Wbs y en caso de estar vacio, Muestro este mensjae de error                              *
else:
    wtree=wbs                                                                                                                           # Caso contrario, paso "wbs" a "wtree" que es un str para poder utilizarla en el programa.                              * 
#****************************************************************************************************************************************************************************************************************************************************************