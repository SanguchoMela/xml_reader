import tkinter as tk
from tkinter import filedialog

from lectorfact import procesar_carpeta_facturas, exportar_facturas_excel, data_facturas

from lectorRet import procesar_carpeta_retenciones, exportar_retenciones_excel, data_retenciones

# Definicion de ventana principal
window = tk.Tk()
window.title("Lector xml")
window.geometry("450x300")

frame_superior = tk.Frame(window)
frame_superior.pack(fill=tk.BOTH, expand=True)
# frame_superior.pack(side=tk.TOP, fill=tk.X)

frame_proceso = tk.Frame(window)
frame_proceso.pack(fill=tk.BOTH, expand=True)
# frame_proceso.pack(side=tk.BOTTOM, fill=tk.X)

def limpiarFrame(frame):
    # Reiniciar variables
    global rutaCarpeta, rutaExcel, data_facturas, data_retenciones
    rutaCarpeta = ""
    rutaExcel = ""
    data_facturas.clear()
    data_retenciones.clear()
    # Eliminar widgets
    for widget in frame_proceso.winfo_children():
        widget.destroy()

# Funcion para abrir el explorador de archivos
def abrirCarpeta():
    global rutaCarpeta
    rutaCarpeta = filedialog.askdirectory()
    if rutaCarpeta:
        labelCarpeta = tk.Label(frame_proceso, text="La carpeta que seleccionaste es: "+rutaCarpeta)
        labelCarpeta.pack()
        
        if isFactura: 
            existen = procesar_carpeta_facturas(rutaCarpeta)
            
            if existen: 
                labelGuardar = tk.Label(frame_proceso, text="Proceso completo, elige el lugar para guardar tu resumen")
                labelGuardar.pack()
                
                buttonSave = tk.Button(frame_proceso,text="Guardar como", width=20, command=guardarExcelResumen)
                buttonSave.pack()
            else:
                labelGuardar = tk.Label(frame_proceso, text="No se encontraron archivos en la carpeta.")
                labelGuardar.pack()
                
                limpiarButton = tk.Button(frame_proceso, text="Limpiar", command=lambda:limpiarFrame(frame_proceso))
                limpiarButton.pack()
                
        else:
            existen_retenciones = procesar_carpeta_retenciones(rutaCarpeta)
            
            if existen_retenciones: 
                labelGuardar = tk.Label(frame_proceso, text="Proceso completo, elige el lugar para guardar tu resumen")
                labelGuardar.pack()
                
                buttonSave = tk.Button(frame_proceso,text="Guardar como", width=20, command=guardarExcelResumen)
                buttonSave.pack()
            else:
                labelGuardar = tk.Label(frame_proceso, text="No se encontraron archivos en la carpeta.")
                labelGuardar.pack()
                
                limpiarButton = tk.Button(frame_proceso, text="Limpiar", command=lambda:limpiarFrame(frame_proceso))
                limpiarButton.pack()
            
    else:
        labelCarpeta = tk.Label(frame_proceso, text="No seleccionaste ninguna carpeta")
        labelCarpeta.pack()
        
        limpiarButton = tk.Button(frame_proceso, text="Limpiar", command=lambda:limpiarFrame(frame_proceso))
        limpiarButton.pack()
        # print("No seleccionaste ninguna carpeta")

# Funcion para guardar los archivos xml ya procesados
def guardarExcelResumen():
    global rutaExcel
    rutaExcel = filedialog.asksaveasfilename()
    # print("Ruta excel: "+rutaExcel)
    if rutaExcel:
        if isFactura:
            exportar_facturas_excel(data_facturas, rutaExcel)
        else:
            exportar_retenciones_excel(data_retenciones, rutaExcel)
            
        labelGuardarExcel = tk.Label(frame_proceso, text="Tu resumen se guardará en: "+rutaExcel)
        labelGuardarExcel.pack()
        # print("Tu resumen se guardará en: "+rutaExcel)
        
        # Borrar el frame de facturas
        limpiarButton = tk.Button(frame_proceso, text="Limpiar", command=lambda:limpiarFrame(frame_proceso))
        limpiarButton.pack()
    else:
        labelGuardarExcel = tk.Label(frame_proceso, text="No seleccionaste ningun lugar")
        labelGuardarExcel.pack()
        
        limpiarButton = tk.Button(frame_proceso, text="Limpiar", command=lambda:limpiarFrame(frame_proceso))
        limpiarButton.pack()
        # print("No seleccionaste ninguna carpeta")

# Funcion para abrir un Toplevel al presionar un boton
def seccionFacturas():
    global isFactura
    labelSelect = tk.Label(frame_proceso,text="Selecciona la carpeta con facturas en xml")
    labelSelect.pack()
    
    buttonSelect = tk.Button(frame_proceso,text="Seleccionar", width=20, command=abrirCarpeta)
    buttonSelect.pack()
    isFactura = True

def seccionRetenciones():
    global isFactura
    labelSelect = tk.Label(frame_proceso,text="Selecciona la carpeta con retenciones en xml")
    labelSelect.pack()
    
    buttonSelect = tk.Button(frame_proceso,text="Seleccionar", width=20, command=abrirCarpeta)
    buttonSelect.pack()
    isFactura = False


# Estilizado de la ventana principal
labelProgram = tk.Label(frame_superior,text="Elige el tipo de documento a leer")
labelProgram.pack()

botonFact = tk.Button(frame_superior, text="Facturas", width=25, command=seccionFacturas)
botonFact.pack()

botonFact = tk.Button(frame_superior, text="Retenciones", width=25, command=seccionRetenciones)
botonFact.pack()

# Para mostrar la ventana principal
window.mainloop()
