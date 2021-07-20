from os import path, mkdir
from pathlib import Path
from datetime import date
from orden import Orden
import openpyxl as op


src_path = Path(__file__).parent
dowload_path = src_path.parent / 'download'
excel_path = src_path.parent / 'excel'
bot_path = src_path.parent / 'invoice bot'


# Verifica si existe el archivo excel 
def existe_archivo_excel(nombre_archivo):
    if path.exists(f'{excel_path}\\{nombre_archivo}.xlsx'):
        return True
    else: 
        raise Exception(f"Archivo no encontrado en la carpeta {excel_path}") 

def verificacion_carpetas():
    if path.exists(dowload_path) and path.exists(excel_path) and path.exists(bot_path):
        True
    else:
        print("Creacion de carpetas esenciales")
        if not path.exists(dowload_path):
            mkdir(f'{dowload_path}')
        if not path.exists(excel_path):
            mkdir(f'{excel_path}')
        print("Carpetas creadas")

def crear_carpeta():
    if path.exists(f'{dowload_path}\\invoices-{date.today()}'):
        print(f'Se encontro la carpeta .\invoices-{date.today()} en la carpeta .\download')
        ruta = f'{dowload_path}\\invoices-{date.today()}'
        return ruta
    else: 
        print(f'No se encontro la carpeta .\invoices-{date.today()} en la carpeta .\download. Se creara')
        mkdir(f'{dowload_path}\\invoices-{date.today()}')
        nueva_ruta = f'{dowload_path}\\invoices-{date.today()}'
        return nueva_ruta

def obtencion_columnas(archivo_excel):
    
    cabeceras = ["orden","user","password"] 
    nr_columnas = {}
    
    for cabecera in cabeceras:
        for columna in range(1, archivo_excel.max_column + 1):
            if cabecera == archivo_excel.cell(row = 1, column = columna).value :
                nr_columnas[cabecera]=columna

    return nr_columnas
    

def lectura_lista_orden(nombre_archivo):
    
    # Abro un archivo leo su hoja principal
    lista_invoice = op.load_workbook(f'{excel_path}\\{nombre_archivo}.xlsx')
    lista_invoice = lista_invoice.active

    columnas = obtencion_columnas(lista_invoice)

    invoices = []

    for orden in range(2, lista_invoice.max_row + 1):
        try:
            invoices.append(Orden(lista_invoice.cell(row=orden, column= columnas["user"]).value,
                                  lista_invoice.cell(row=orden, column = columnas["password"]).value,
                                 {"nombre":str(lista_invoice.cell(row = orden, column= columnas["orden"]).value).replace("|","").rstrip(),
                                  "link": lista_invoice.cell(row = orden, column= columnas["orden"]).hyperlink.target}))

        except Exception as ex:
            pass

    return invoices


