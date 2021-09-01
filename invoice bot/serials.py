import json

def convertir_json(driver):
    #convierto en diccionario la informacion del invoice
    data = json.loads(driver.find_element_by_xpath('//script[@id="init_data"]').get_attribute('innerHTML'))
    return data

def obtener_items(data_items, lista_items, orden):

    for item in lista_items:
        orden.serials.append({"nombre_producto": data_items[item]["d"]["productName"].replace("-USA","").rstrip(),
                              "lista_serials": data_items[item]["d"]["lineItemSerialInfo"]})
        
               
def info_serial(driver, orden):
    # Convierto en diccionario la informacion del invoice
    data_script = convertir_json(driver)

    data_invoice = data_script["orderInvoices"]["c"]

    for invoice in data_invoice:
        obtener_items(data_script["orderInvoices"][invoice]["invoiceLineItems"], 
                      data_script["orderInvoices"][invoice]["invoiceLineItems"]["c"],
                      orden)
        


    


