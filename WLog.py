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
#****************************************************************************************************************************************************************************************************************************************************************
import openpyxl
from openpyxl import Workbook,load_workbook
from Variables import *
#****************************************************************************************************************************************************************************************************************************************************************
log=open(user  + (str(lineas[1])[:-1]),mode="a")                                                                                # Abrimos el Archivo .txt que utilizamos para registrar los logs. Este esta enrutado en el Path                             *
log.write(Mensaje)
log.close()       
#****************************************************************************************************************************************************************************************************************************************************************