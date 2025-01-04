import os
import pandas as pd
import xml.etree.ElementTree as ET
from estilos_excel import dar_estilo

data_facturas = []
existen = None

# Funcion para procesar las etiquetas de cada archivo xml
def procesar_facturas(ruta_archivo):
    try:
        # Abrir el archivo XML
        arbol = ET.parse(ruta_archivo)
        # Obtener la raíz del árbol
        ruta = arbol.getroot()
        
        # obtener el comprobante dentro del CDATA
        comprobante_cdata = ruta.find('comprobante').text.strip()

        # parsear el comprobante xml contenido en el CDATA
        comprobante_root = ET.fromstring(comprobante_cdata)
        
        info_tributaria = comprobante_root.find('infoTributaria')
        if info_tributaria is not None:
            razon_social_vendedor = info_tributaria.find('razonSocial').text
            ruc_vendedor = info_tributaria.find('ruc').text
            cod_doc = info_tributaria.find('codDoc').text
            estab = info_tributaria.find('estab').text
            pto_emi = info_tributaria.find('ptoEmi').text
            secuencial = info_tributaria.find('secuencial').text
            
        # # Obtener información de la factura
        factura_info = comprobante_root.find('infoFactura')
        if factura_info is not None:
            fecha_emision_factura = factura_info.find('fechaEmision').text
            razon_social_comprador = factura_info.find('razonSocialComprador').text
            ruc_comprador = factura_info.find('identificacionComprador').text
            total_sin_impuestos = factura_info.find('totalSinImpuestos').text
            importe_total_factura = factura_info.find('importeTotal').text
            propina_elem = factura_info.find('propina')
            if propina_elem is not None:
                propina = propina_elem.text
            else:
                propina = "null"
            
        # Obtener información del total con impuestos
        total_con_impuesto_elem = factura_info.find('totalConImpuestos')
        
        # Inicializar las variables para evitar UnboundLocalError
        subtotal15 = ""
        iva15 = ""
        subtotal0 = ""
        otraTarifa = ""
        otroIVA = ""
        
        for total_impuesto in total_con_impuesto_elem.findall('totalImpuesto'):
            # Extraer los datos del elemento 'totalImpuesto'
            base_imponible = total_impuesto.find('baseImponible').text
            codigo_porcentaje = total_impuesto.find('codigoPorcentaje').text
            valor_iva = total_impuesto.find('valor').text
            
            if codigo_porcentaje == "4":
                subtotal15 = base_imponible
                iva15 = valor_iva
            elif codigo_porcentaje == '0':
                subtotal0 = base_imponible
            else:
                otraTarifa = base_imponible
                otroIVA = valor_iva
                
        # Obtener información de detalles
        detalles_elem = comprobante_root.find('detalles')
        descripciones = []
        
        for i, detalle in enumerate(detalles_elem.findall('detalle')):
            if i > 3:
                break
            # Extraer datos de cada elemento detalle
            descripcion = detalle.find('descripcion').text
            descripciones.append(descripcion)
            
        descripcion1 = descripciones[0] if len(descripciones) > 0 else ""
        descripcion2 = descripciones[1] if len(descripciones) > 1 else ""
        descripcion3 = descripciones[2] if len(descripciones) > 2 else ""

        # Agregar información a la lista
        data_facturas.append({
            "info_xml": {
                "Razón Social Comprador": razon_social_comprador,
                "RUC Comprador": ruc_comprador,
                "Razón Social Vendedor": razon_social_vendedor,
                "RUC Vendedor": ruc_vendedor,
                "Fecha de Emisión": fecha_emision_factura,
                "Numero de Factura": f"{cod_doc}-{estab}-{pto_emi}-{secuencial}",
                "Propina": propina,
                "Subtotal 0%": subtotal0,
                "Subtotal 15%": subtotal15,
                "Otra tarifa": otraTarifa,
                "Subtotal sin Impuestos": total_sin_impuestos,
                "IVA 15%": iva15,
                "Otro IVA": otroIVA,
                "Total": importe_total_factura,
                "Descripcion 1": descripcion1,
                "Descripcion 2": descripcion2,
                "Descripcion 3": descripcion3,
            }
        })
            
    except ET.ParseError as e:
        print(f"Error en archivo {ruta_archivo}: {e}")


def exportar_facturas_excel(data, excel_file):
    try:        
        # Crear DataFrame
        df = pd.DataFrame([fila["info_xml"] for fila in data_facturas])
        
        # Covertir varias columnas a numeros
        for column in ["Propina","Subtotal 0%", "Subtotal 15%", "Otra tarifa", "Subtotal sin Impuestos", "Otro IVA", "IVA 15%", "Total"]:
            df[column] = pd.to_numeric(df[column], errors='coerce')
            
        df.fillna(0, inplace=True)

        # Escribir DataFrame a Excel
        df.to_excel(excel_file, index=False)
        
        col_no_centrar = ["Razón Social Comprador", "Razón Social Vendedor", "Descripcion 1", "Descripcion 2", "Descripcion 3"]
            
        dar_estilo(excel_file, df, col_no_centrar)
        
        print(f"Datos exportados a '{excel_file}' correctamente.")
    except Exception as e:
        print(f"Error al exportar(): {e}")


# Funcion para crear un arreglo de archivos xml a partir de una carpeta
def procesar_carpeta_facturas(ruta_carpeta):
    global existen
    existen = False
    
    # listar archivos
    archivos = os.listdir(ruta_carpeta)

    # filtrar solo archivos xml
    archivos_xml = [archivo for archivo in archivos if archivo.endswith(".xml")]
    
    if archivos_xml:
        existen = True

    # procesar cada archivo xml
    for archivo_xml in archivos_xml:
        ruta_completa = os.path.join(ruta_carpeta,archivo_xml)
        procesar_facturas(ruta_completa)
        
    return existen


# # ruta de la carpeta con archivos xml
# carpeta_xml = "C:/Users/Hp/Downloads/drP_jun24"

# # procesar todos los xml en la carpeta
# procesar_carpeta_xml(carpeta_xml)

# # exportar datos a un archivo excel
# excel_ruta = "C:/Users/Hp/Escritorio/Conta/xmlReader/resumenExcel/pruebaDescr.xlsx"

# # print(exportar_a_excel(data_xml, excel_ruta))

# try:
#     exportar_a_excel = (data_xml, excel_ruta)
# except Exception as e:
#     print(f"Error al exportar: {e}")
