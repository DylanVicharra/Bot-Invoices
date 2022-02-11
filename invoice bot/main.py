import os
import webdriver as bot
import manejo_datos as md
import invoice as inv
from selenium.common.exceptions import TimeoutException, NoSuchWindowException, SessionNotCreatedException, WebDriverException


def secuencia_invoices(ejecutableChrome, carpeta_descarga, orden, tiempo_espera):

    try:
        driver = bot.crear_webdriver(ejecutableChrome, carpeta_descarga)
    except SessionNotCreatedException:
        print("La version de Chrome no corresponde con el webdriver que se utiliza." + "\n" +
              "Actualice su navegador Chrome a la ultima version disponible.")
        exit(1)

    funciones_invoices = {inv.pagina_orden:[driver, orden, tiempo_espera], inv.login_appleID:[driver, orden, tiempo_espera], inv.seleccion_invoice:[driver, orden, tiempo_espera]}

    for funcion in funciones_invoices:
        try:
            funcion(*funciones_invoices[funcion])
        except (NoSuchWindowException, WebDriverException):
            print("Ocurrio un error con la ventana del navegador.")
            print(f'No se pudo descargar la factura de la orden {orden.orden["nombre"]}')
            orden.estado = False 
            break    
        except TimeoutException:
            print('Se demoro en encontrar el boton o texto.')   
            print(f'No se pudo descargar la factura de la orden {orden.orden["nombre"]}')  
            orden.estado = False
            break
        except Exception as ex:
            error_message = '\n'.join(map(str, ex.args)).rstrip()
            print(f"{error_message}")
            print(f'No se pudo descargar la factura de la orden {orden.orden["nombre"]}')
            orden.estado = False 
            break    

    driver.quit()


def main():
    os.system('cls')
    print("                 ============ BOT - INVOICES APPLE ============                 ")
    print("Se usara como archivo predeterminado el 'BOT - INVOICES APPLE.xlsx' y la primera hoja disponible")

    # Nombres del archivo (para modificar mas facil)
    archivo_excel = 'BOT - INVOICES APPLE'.rstrip()
    # Tiempo maximo de espera 
    tiempo_espera = 8
    # Variable de errores
    errores = 0

    # Elimina el anterior log
    md.eliminar_archivo_texto()

    print("Verificacion de las carpetas download, excel e invoice bot")
    md.verificacion_carpetas()

    print("Instalacion del Webdriver de Chrome")
    try:
        ejecutableChrome = bot.instalar_webdriver()
    except:
        print("Archivo posiblemente dañado. Borrar la carpeta .wdm y volver a iniciar el programa")
        exit(1)
    print(f'Ruta: {ejecutableChrome}')

    print("Verificando la existencia de la carpeta de descarga")
    carpeta_descarga = md.crear_carpeta()

    # Creo un archivo excel donde guardar los serials
    nombre_archivo = "Informacion"
    archivo_general = md.crear_archivo(nombre_archivo)

    print(f"Lectura del archivo {archivo_excel}.xlsx")
    try: 
        md.existe_archivo_excel(archivo_excel)
    except Exception as ex: 
        print(f'{ex}' + '\n' + "Finalizando..." + '\n')
        exit(1)
        
    lista_orden = md.lectura_lista_orden(archivo_excel)
    
    if lista_orden:

        # Leo la lista de ordenes
        for orden in lista_orden:
            
            secuencia_invoices(ejecutableChrome, carpeta_descarga, orden, tiempo_espera)

            #Si hay un error lo guarda y despues lo escribe en un txt 
            if not orden.estado:
                md.escribir_errores(archivo_general, "Errores", orden.orden)
                errores+=1
            else: 
                # Secuencia de guardado de serial dentro del archivo anterior creado
                md.escribir_informacion(archivo_general, "Informacion", orden.orden, orden.email, orden.informacion["order_date"],orden.informacion["invoice_date"], orden.informacion["credit_card"], orden.informacion["number_card"], orden.informacion["total"])
                for producto in orden.serials:
                    md.escribir_serial(archivo_general,"Serials",orden.orden,  producto["nombre_producto"], producto["lista_serials"])

        # Guardo la informacion de los serial
        md.guardar_archivo_excel(archivo_general, nombre_archivo)
        # Finalizo la sesion de los excels usados
        archivo_general.close()
        
    else:
        print(f'El archivo {archivo_excel}.xlsx esta vacio. Ingresar datos al archivo.')
        # Finalizo la sesion de los excels usados
        archivo_general.close()
    
    if errores > 0:
        print("No se ha podido descargar ciertos invoices. Se detallan en la solapa 'Errores' del archivo excel")

    print("Finalizando BOT - INVOICES APPLE...")


if __name__ == "__main__":
    main()