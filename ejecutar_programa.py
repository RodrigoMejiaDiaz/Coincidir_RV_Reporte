from tkinter import Tk
from main import ComprobarRegistroVentas  # Asegúrate de que 'main' es el nombre del archivo de tu clase

def start_program():
    root = Tk()
    ComprobarRegistroVentas(root)
    root.mainloop()

if __name__ == "__main__":
    start_program()