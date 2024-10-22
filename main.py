import codecs
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import threading
from pathlib import Path
import pandas as pd
from datetime import datetime

class ComprobarRegistroVentas:
    def __init__(self, root):
        self.root = root
        self.dataReporte = {}
        self.dataRV = {}
        # Variables que almacenan archivos
        self.archivoRV = ""
        self.archivoReporte = ""
        
        # Variable para almacenar el progreso de la barra de carga General
        self.progresoGeneral = 0
        
        ## Configuracion TKinter
        
        self.root.title("Coincidir Registro Ventas")
        self.root.resizable(FALSE, FALSE)
        
        # Especificar el tamaño de la ventana
        ancho_ventana = 400
        alto_ventana = 375

        # Configuración del mainframe
        mainframe = ttk.Frame(self.root, width=ancho_ventana, height=alto_ventana, padding="3 3 12 12")
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        
        # Añadir los widgets a utilizar y ubicarlos en el grid
        lblArchivoRV = ttk.Label(mainframe, text="Selecciona el Registro de Ventas:", wraplength=100)
        
        self.btnSeleccionarRV = ttk.Button(mainframe, text="Seleccionar Registro Ventas", command=self.seleccionarArchivoRV)
        
        # Variable para almacenar estado de archivo seleccionado
        self.estaSeleccionadoRV = StringVar()
        self.estaSeleccionadoRV.set("")
        self.lblCargadosRV = ttk.Label(mainframe, textvariable=self.estaSeleccionadoRV, wraplength=100)
        
        # Vincular función para verificar cambios
        self.estaSeleccionadoRV.trace_add("write", self.escuchar_cambios_seleccionado)
        
        lblArchivoReporte = ttk.Label(mainframe, text="Selecciona el Reporte:", wraplength=100)
        
        self.btnSeleccionarReporte = ttk.Button(mainframe, text="Seleccionar Reporte", command=self.seleccionarArchivoReporte)
        
        # Variable para almacenar estado de archivo seleccionado
        self.estaSeleccionadoReporte = StringVar()
        self.estaSeleccionadoReporte.set("")
        self.lblCargadosReporte = ttk.Label(mainframe, textvariable=self.estaSeleccionadoReporte, wraplength=125)
        
        # Vincular función para verificar cambios
        self.estaSeleccionadoReporte.trace_add("write", self.escuchar_cambios_seleccionado)
        
        self.btnCoincidir = ttk.Button(mainframe, text="Coincidir", command=self.btnCoincidir_handler)
        btnCerrar = ttk.Button(mainframe, text="Cerrar Programa", command=self.cerrarPrograma)
        
        # Configuracion grid
        lblArchivoRV.grid(column=0, row=0)
        self.btnSeleccionarRV.grid(column=1, row=0)
        self.lblCargadosRV.grid(column=2, row=0)
        lblArchivoReporte.grid(column=0, row=1)
        self.btnSeleccionarReporte.grid(column=1, row=1)
        self.lblCargadosReporte.grid(column=2, row=1)
        self.btnCoincidir.grid(column=1,row=2)
        btnCerrar.grid(column=1, row=3)
        
        # Cambiar estados a desabilitados
        self.btnCoincidir.state(['disabled'])
        
        # Barra de progreso
        self.progreso = ttk.Progressbar(mainframe, orient="horizontal", mode="determinate")
        self.progreso.grid(column=0, row=4, sticky="nswe", columnspan=3)

        # Etiqueta que muestra el progreso en texto
        self.label_progreso = ttk.Label(mainframe, text="Progreso: 0%")
        self.label_progreso.grid(column=1, row=5, sticky="ns")
        
        self.nroCoincidencias = StringVar()
        self.nroCoincidencias.set('0/0 Coincidencias')
        self.lblCoincidencias = ttk.Label(mainframe, textvariable=self.nroCoincidencias)
        self.lblCoincidencias.grid(column=1, row=6)
        
        # Centra ventana al abrir programa
        self.centrar_ventana(self.root, ancho_ventana, alto_ventana)
        
        # Configura el peso del programa padre para que se amplie con la ventana
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Asigna los pesos para que se amplien con la ventana
        mainframe.columnconfigure(0, weight=1, uniform="col")
        mainframe.columnconfigure(1, weight=1, uniform="col")
        mainframe.columnconfigure(2, weight=1, uniform="col")
        
        mainframe.rowconfigure(0, weight=1)
        mainframe.rowconfigure(1, weight=1)
        mainframe.rowconfigure(2, weight=1)
        mainframe.rowconfigure(3, weight=1)
        mainframe.rowconfigure(4, weight=1)
        mainframe.rowconfigure(5, weight=1)
        mainframe.rowconfigure(6, weight=1)
        
        # Loop para dar padding a todos los widgets hijos
        for child in mainframe.winfo_children(): 
            child.grid_configure(padx=5, pady=5)
    
    def cerrarPrograma(self):
        self.root.destroy() 
    
    # Función para que la ventana principal (self.root) se abra en el centro de la pantalla
    def centrar_ventana(self, ventana, ancho, alto):
        # Obtener el ancho y alto de la pantalla
        ancho_pantalla = ventana.winfo_screenwidth()
        alto_pantalla = ventana.winfo_screenheight()

        # Calcular las coordenadas x, y para centrar la ventana
        x = (ancho_pantalla // 2) - (ancho // 2)
        y = (alto_pantalla // 2) - (alto // 2)

        # Fijar las dimensiones y la posición de la ventana
        ventana.geometry(f'{ancho}x{alto}+{x}+{y}')
    
    # Verificar cambios self.estan_cagados
    def escuchar_cambios_seleccionado(self, *args):
        esta_SeleccionadoRV = self.estaSeleccionadoRV.get()
        esta_SeleccionadoReporte = self.estaSeleccionadoReporte.get()
        if esta_SeleccionadoRV != "" and esta_SeleccionadoRV != "No se seleccionó archivo" and esta_SeleccionadoReporte != "" and esta_SeleccionadoReporte != "No se seleccionó archivo":
            self.btnCoincidir.state(['!disabled'])
        else:
            self.btnCoincidir.state(['disabled'])
    
    # Función para abrir ventana emergente para seleccionar los archivos
    def seleccionarArchivoRV(self):
        # Abrir ventana emergente para seleccionar los archivos excel a subir
        self.archivoRV = filedialog.askopenfilename(filetypes=[("Archivos TXT", "*.txt")])
        
        # Comprobar si se han seleccionado archivos
        if self.archivoRV:
            # Cambiar color al label a negro en caso de que este en rojo
            self.lblCargadosRV.config(foreground="black")
            
            print("Archivos seleccionados:", self.archivoRV)
            
            
            # Extraer nombre del archivo Registro Ventas para conseguir fecha y mes
            archivoRV = Path(self.archivoRV)
            archivoRV = archivoRV.stem
            archivoRV = archivoRV[13:]
            self.fecha = archivoRV[:6]
            self.anio = self.fecha[:4]
            self.mes = self.fecha[4:]
            
            # Cambiar el valor de la etiqueta a 'estaSeleccionado'
            self.estaSeleccionadoRV.set(f"Mes: {self.mes} - Año: {self.anio}")
        
        else:
            self.estaSeleccionadoRV.set("No se seleccionó archivo")
            self.lblCargadosRV.config(foreground="red")
    
    # Función para abrir ventana emergente para seleccionar los archivos
    def seleccionarArchivoReporte(self):
        # Abrir ventana emergente para seleccionar los archivos excel a subir
        self.archivoReporte = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        archivoReporte = Path(self.archivoReporte)
        
        # Comprobar si se han seleccionado archivos
        if self.archivoReporte:
            # Cambiar color al label a negro en caso de que este en rojo
            self.lblCargadosReporte.config(foreground="black")
            
            print("Archivos seleccionados:", self.archivoReporte)
            
            # Cambiar el valor de la etiqueta a 'estaSeleccionado'
            self.estaSeleccionadoReporte.set(f"Archivo Seleccionado: {archivoReporte.stem} ")
        
        else:
            self.estaSeleccionadoReporte.set("No se seleccionó archivo")
            self.lblCargadosReporte.config(foreground="red")
    
    # Función para convertir fecha a dd/mm/yyyy y hora a h:m:s
    def convertirDatetimeString(self, fecha):
        if fecha != "0":
            fecha_formateada = fecha.strftime("%d/%m/%Y")
            return fecha_formateada
        
        
    def convertirStringDatetime(self, fecha):
        if fecha != "0":
            if isinstance(fecha, datetime):
                return fecha
            else:
                fecha_formateada = datetime.strptime(fecha,"%d/%m/%Y")
                return fecha_formateada
            
    def limpiar_numero(self, numero):
        # Convierte el número a cadena
        numero_str = str(numero)
        
        numero_str = numero_str.replace('.0', '')
        
        if numero_str == '':
            numero_str = '0'
            
        if numero_str == '10000000000000':
            numero_str = '0'
        
        if numero_str == None:
            numero_str = '0'
            
        if numero_str == '-':
            numero_str = '0'
        
        if numero_str.endswith('.'):
            numero_str = numero_str.rstrip('.')

        if numero_str == '5000':
            numero_str = '5.0'
        if numero_str == '3000':
            numero_str = '3.0'
        
        return numero_str
    
    # Función para actualizar barra de progreso    
    def actualizar_progresoGeneral(self):
        ## Actualizar barra de progreso
        # aumentar en 1 el progreso
        self.progresoGeneral += 1
        
        # Actualizar la barra de progreso
        self.progreso["value"] = self.progresoGeneral  # Actualizar el valor de la barra
        self.label_progreso.config(text=f"Progreso: {int((self.progresoGeneral) / 3 * 100)}%")

        # Actualizar la interfaz gráfica
        self.root.update_idletasks() 
    
    # Función para reiniciar el progreso de la barra General       
    def reiniciar_progresoGeneral(self):
        # Reiniciar la variable de progresoGeneral
        self.progresoGeneral = 0
        # Actualizar la barra de progreso
        self.progreso["value"] = self.progresoGeneral  # Actualizar el valor de la barra
        self.label_progreso.config(text="Progreso: 0%")
        self.progreso["maximum"] = 3
    
    def btnCoincidir_handler(self):
        # Limpiar diccionario
        self.dataReporte.clear()
        self.dataRV.clear()
        self.reiniciar_progresoGeneral()
        self.nroCoincidencias.set(f"0/0 Coincidencias")
        
        # Ejecutar la carga de data en un hilo separado
        threading.Thread(target=self.cargar_data).start()
        
    def cargar_data(self):
        
        self.leer_reporte(self.archivoReporte, self.anio, self.mes)
        
        self.actualizar_progresoGeneral()
        
        self.leer_txt(self.archivoRV)
        
        self.actualizar_progresoGeneral()
        
        coincidencias = f"{len(self.dataReporte)}"
        
        self.nroCoincidencias.set(f"0/{coincidencias} Coincidencias")
        
        self.coincidir_datas(self.dataRV, self.dataReporte)
        
        self.actualizar_progresoGeneral()
        
    
    def leer_reporte(self, reporte, anio, mes):
        # Leer el archivo Excel
        xls = pd.ExcelFile(reporte)
        
        # Iterar sobre las hojas del archivo
        for hoja in xls.sheet_names:
            if 'BOLETAS' in hoja.upper():
                # Leer la hoja que contiene 'Boletas' y columnas A:O
                df = pd.read_excel(xls, sheet_name=hoja, usecols="A:O")
                dfFiltrado = df[['Fecha', 'RUC', 'Tarifa', 'Boletas', 'Ticketera']]
                
        i = 0
        for _, row in dfFiltrado.iterrows():
            i += 1
            fecha = self.convertirStringDatetime(row['Fecha'])
            tarifa = self.limpiar_numero(row['Tarifa'])
            boleta = str(row["Boletas"])
            caseta = boleta[:3]
            numero = boleta[3:]
            boleta13 = f'{caseta}{int(numero):010d}'
            ruc = self.limpiar_numero(row['RUC'])
            if fecha.month == int(mes) and fecha.year == int(anio):
                key = (fecha, i)
                fecha = self.convertirDatetimeString(fecha)
                self.dataReporte[key] = {'fecha': fecha, 'ruc': ruc, 'monto': tarifa, 'boleta': boleta13, 'ticketera': row['Ticketera']}
                
    def leer_txt(self, ruta_archivo):
        i = 0
        f = codecs.open(ruta_archivo, "r", "ISO-8859-1")
        for line in f:
            i += 1
            # separate line into fields
            fields = line.split("|")
            
            fecha = self.convertirStringDatetime(fields[3])
            ticketera = fields[6]
            boleta = fields[7]
            
            if len(boleta) > 13:
                boleta = boleta[1:]
            
            ruc = self.limpiar_numero(fields[10])
            
            
            monto = fields[23]
            
            key = (fecha, i)
            
            fecha = self.convertirDatetimeString(fecha)
            
            self.dataRV[key] = {'fecha': fecha, 'ruc': ruc, 'monto': monto, 'boleta': boleta, 'ticketera': ticketera}
            
            
    def coincidir_datas(self, dataRV, dataReporte):
        # Inicializar el contador de coincidencias
        coincidencias = 0
        lista_coincidencias =[]
        # Iterar sobre los datos de dataRV
        for key_rv, data_dataRV in dataRV.items():
            fecha_rv = data_dataRV['fecha']
            ruc_rv = str(data_dataRV['ruc'])
            monto_rv = str(data_dataRV['monto'])
            boleta_rv = str(data_dataRV['boleta'])
            ticketera_rv = str(data_dataRV['ticketera'])
            # Iterar sobre los datos de dataReporte
            for key_reporte, data_dataReporte in dataReporte.copy().items():
                fecha_reporte = data_dataReporte['fecha']
                ruc_reporte = str(data_dataReporte['ruc'])
                monto_reporte = str(data_dataReporte['monto'])
                boleta_reporte = str(data_dataReporte['boleta'])
                ticketera_reporte = str(data_dataReporte['ticketera'])
                # Verificar si todos los campos coinciden
                if (fecha_rv == fecha_reporte and ruc_rv == ruc_reporte and
                    monto_rv == monto_reporte and
                    boleta_rv == boleta_reporte and
                    ticketera_rv == ticketera_reporte):
                    # Si hay coincidencia, aumentar el contador
                    coincidencias += 1
                    # Actualizar el número de coincidencias en el StringVar
                    total_items = len(dataReporte)
                    self.nroCoincidencias.set(f"{coincidencias}/{total_items} Coincidencias")
                    print(f"Coincidencia: {data_dataReporte}")
                    lista_coincidencias.append(key_reporte)
        
        if len(lista_coincidencias) != len(dataReporte):
            for key in list(dataReporte.keys()):
                if key in lista_coincidencias:
                    del dataReporte[key]
            
            
            print("\n")
            print(f"No se encontraron coincidencias de: ")
            for key, data_key in dataReporte.items():
                print(f"FECHA: {data_key['fecha']}")
                print(f"RUC: {data_key['ruc']}")
                print(f"MONTO: {data_key['monto']}")
                print(f"BOLETA: {data_key['boleta']}")
                print(f"TICKETERA: {data_key['ticketera']}")
                print("\n")
        else:
            print("\n")
            print("Todas coinciden!")
                    
# Iniciar programa                
                
root = Tk()
ComprobarRegistroVentas(root)
root.mainloop()