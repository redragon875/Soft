import cmd
import os as os
from select import select
import sys as sys
from sys import *
from os import replace, system as system
import datetime,time
#****************************************************************************************************************************************************************************************************************************************************************
import tkinter as tk
from tkinter import *
from tkinter import Entry, Grid, Image, StringVar, Text, Variable, messagebox, ttk, scrolledtext, simpledialog, tix, font, commondialog
from tkinter.ttk import Progressbar, Style, Treeview, setup_master,Sizegrip,Entry,Spinbox
from tkinter.tix import STATUS, LabelEntry,LabelFrame,Meter, ButtonBox,PhotoImage,ComboBox
from tkinter.constants import *
from turtle import color, delay, title, width
#****************************************************************************************************************************************************************************************************************************************************************
import openpyxl
from openpyxl import Workbook,load_workbook
#****************************************************************************************************************************************************************************************************************************************************************
from Variables import *
from buscador import Tree
global wtree

qst=messagebox.askokcancel(title="Modificar Parametros",message="Esta seguro que decea Modificar los Parametros???")
if qst==True:
    current_item = Tree.focus()    
    wbs = Tree.item(current_item,option='text') 
    wtree=str(wbs)
    import Destroy
    import Modificar
    
else:
    Mensaje="Se cancelo modificación solicitada"
    messagebox.showerror(title="Modificación",message=Mensaje)