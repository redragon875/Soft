def Cerrar_myapp():
    from Inicio     import myapp as inicio
    inicio.destroy()
    print("Se cierra Myapp")
    return

def Cerrar_mybus():
    from buscador   import mybus as bus
    bus.destroy()
    print("Se cierra Mybus")
    return