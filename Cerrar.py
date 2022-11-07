
#****************************************************************************************************************************************************************************************************************************************************************
# Nombre: Cerrar.Py
# Descripcion: Consta de un simple programita que realizara el inicio de la aplicación. Muestra botones para poder ejecutar el Script correspondiente al buscador o al diario del programa
#****************************************************************************************************************************************************************************************************************************************************************
import os as os
import sys as sys
from sys import *
from os import replace, system as system
import inicio(user)

#****************************************************************************************************************************************************************************************************************************************************************
def cerrar():
    
    log=open(user  + (str(lineas[1])[:-1]),mode="a")
    registro=(str(Hini)+" => Inicio => "+user[9:30] + " Selecciona " +str(dato)+". Se ejecuta scrip y se cierra Inicio \n")
    print (registro)
    log.write(registro)
    log.close()
    myapp.destroy()