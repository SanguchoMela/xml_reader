import os
import pandas as pd
import xml.etree.ElementTree as ET

data_facturas = []

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
                "Subtotal sin Impuestos": total_sin_impuestos,
                "IVA 15%": iva15,
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
        # Separar los datos por diccionario
        arreglo_datos = []
        
        # Iterar sobre cada elemento en data
        for fila in data:
            # Obtener el diccionario de info_xml de cada fila
            info_xml = fila.get("info_xml",{})
            
            # Crear un nuevo diccionario para cada fila aplanada
            nueva_fila = {
                **info_xml,  # Agregar todos los elementos de info_xml como columnas
                "Razón Social Comprador": info_xml.get("Razón Social Comprador", ""),
                "RUC Comprador": info_xml.get("RUC Comprador", ""),
                "Razón Social Vendedor": info_xml.get("Razón Social Vendedor", ""),
                "RUC Vendedor": info_xml.get("RUC Vendedor", ""),
                "Fecha de Emisión": info_xml.get("Fecha de Emisión", ""),
                "Numero de Factura": info_xml.get("Numero de Factura", ""),
                "Propina": info_xml.get("Propina"),
                "Subtotal 0%": info_xml.get("Subtotal 0%"),
                "Subtotal 15%": info_xml.get("Subtotal 15%"),
                "Subtotal sin Impuestos": info_xml.get("Subtotal sin Impuestos", ""),
                "IVA 15%": info_xml.get("IVA 15%", ""),
                "Total": info_xml.get("Total", ""),
                "Descripcion 1": info_xml.get("Descripcion 1", ""),
                "Descripcion 2": info_xml.get("Descripcion 2", ""),
                "Descripcion 3": info_xml.get("Descripcion 3", ""),
            }
                
            arreglo_datos.append(nueva_fila) 
        
        # Crear DataFrame
        df = pd.DataFrame(arreglo_datos)
        
        # Covertir varias columnas a numeros
        for column in ["Propina","Subtotal 0%", "Subtotal 15%", "Subtotal sin Impuestos", "IVA 15%", "Total"]:
            df[column] = pd.to_numeric(df[column], errors='coerce')
            
        df.fillna(0, inplace=True)

        # Escribir DataFrame a Excel
        df.to_excel(excel_file, index=False)
        print(f"Datos exportados a '{excel_file}' correctamente.")
    except Exception as e:
        print(f"Error al exportar(): {e}")


# Funcion para crear un arreglo de archivos xml a partir de una carpeta
def procesar_carpeta_facturas(ruta_carpeta):
    # listar archivos
    archivos = os.listdir(ruta_carpeta)

    # filtrar solo archivos xml
    archivos_xml = [archivo for archivo in archivos if archivo.endswith(".xml")]

    # procesar cada archivo xml
    for archivo_xml in archivos_xml:
        ruta_completa = os.path.join(ruta_carpeta,archivo_xml)
        procesar_facturas(ruta_completa)

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
