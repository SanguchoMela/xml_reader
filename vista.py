import tkinter as tk
from tkinter import filedialog

import lectorfact
from lectorfact import procesar_carpeta_facturas, exportar_facturas_excel, data_facturas

# Definicion de ventana principal
window = tk.Tk()
window.title("Lector xml")
window.geometry("450x300")

frame_superior = tk.Frame(window)
frame_superior.pack(fill=tk.BOTH, expand=True)
# frame_superior.pack(side=tk.TOP, fill=tk.X)

frame_facturas = tk.Frame(window)
frame_facturas.pack(fill=tk.BOTH, expand=True)
# frame_facturas.pack(side=tk.BOTTOM, fill=tk.X)

def limpiarFrame(frame):
    # Reiniciar variables
    global rutaCarpeta, rutaExcel, data_facturas
    rutaCarpeta = ""
    rutaExcel = ""
    data_facturas.clear()
    # Eliminar widgets
    for widget in frame_facturas.winfo_children():
        widget.destroy()

# Funcion para abrir el explorador de archivos
def abrirCarpeta():
    global rutaCarpeta
    rutaCarpeta = filedialog.askdirectory()
    if rutaCarpeta:
        labelCarpeta = tk.Label(frame_facturas, text="La carpeta que seleccionaste es: "+rutaCarpeta)
        labelCarpeta.pack()
        # print("Espera mientras se procesa la carpeta")
        procesar_carpeta_facturas(rutaCarpeta)
        labelGuardar = tk.Label(frame_facturas, text="Proceso completo, elige el lugar para guardar tu resumen")
        labelGuardar.pack()
        
        buttonSave = tk.Button(frame_facturas,text="Guardar como", width=20, command=guardarExcelResumen)
        buttonSave.pack()
    else:
        labelCarpeta = tk.Label(frame_facturas, text="No seleccionaste ninguna carpeta")
        labelCarpeta.pack()
        # print("No seleccionaste ninguna carpeta")

# Funcion para guardar los archivos xml ya procesados
def guardarExcelResumen():
    global rutaExcel
    rutaExcel = filedialog.asksaveasfilename()
    # print("Ruta excel: "+rutaExcel)
    if rutaExcel:
        exportar_facturas_excel(data_facturas, rutaExcel)
        labelGuardarExcel = tk.Label(frame_facturas, text="Tu resumen se guardará en: "+rutaExcel)
        labelGuardarExcel.pack()
        # print("Tu resumen se guardará en: "+rutaExcel)
        
        # Borrar el frame de facturas
        limpiarButton = tk.Button(frame_facturas, text="Limpiar", command=lambda:limpiarFrame(frame_facturas))
        limpiarButton.pack()
    else:
        labelGuardarExcel = tk.Label(frame_facturas, text="No seleccionaste ninguna lugar")
        labelGuardarExcel.pack()
        # print("No seleccionaste ninguna carpeta")

# Funcion para abrir un Toplevel al presionar un boton
def seccionFacturas():
    # print("Ventana facturas")
    # ventanaFact = Toplevel()
    # ventanaFact.title("Lector de facturas")
    # ventanaFact.geometry("400x250")
    
    labelSelect = tk.Label(frame_facturas,text="Selecciona la carpeta con facturas en xml")
    labelSelect.pack()
    
    buttonSelect = tk.Button(frame_facturas,text="Seleccionar", width=20, command=abrirCarpeta)
    buttonSelect.pack()


# Estilizado de la ventana principal
labelProgram = tk.Label(frame_superior,text="Elige el tipo de documento a leer")
labelProgram.pack()

botonFact = tk.Button(frame_superior, text="Facturas", width=25, command=seccionFacturas)
botonFact.pack()

# botonRet = tk.Button(window, text="Retenciones", width=25)
# botonRet.pack()

# Para mostrar la ventana principal
window.mainloop()
