## Comprobar actualizaciones
from actualizar import obtener_ultima_version, actualizar_programa, obtener_version_actual_txt
from main import ComprobarRegistroVentas
from tkinter import *

def main():
    repo = "RodrigoMejiaDiaz/Coincidir_RV_Reporte"
    version_actual = obtener_version_actual_txt()

    print(f"Versión actual: {version_actual}")

    url_actualizacion, ultima_version = obtener_ultima_version(repo, version_actual)
    if url_actualizacion:
        print(f"Nueva versión disponible: {ultima_version}. Actualizando...")
        actualizar_programa(url_actualizacion)
    else:
        print("El programa está actualizado.")
        # Iniciar programa      
        root = Tk()
        ComprobarRegistroVentas(root)
        root.mainloop()

if __name__ == "__main__":
    main()