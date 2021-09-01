
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchWindowException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import serials as serial


def pagina_orden(driver, orden, tiempo_espera):
    
    try:
        print(f'Se esta ingresando a la factura de la orden {orden.orden["nombre"]}  ')
        
        driver.get(orden.orden["link"])

        # Selecciona el boton sign in
        WebDriverWait(driver, tiempo_espera).until(EC.visibility_of_element_located((By.XPATH, '//a[@role="button"][@class="button form-button"]')))
        driver.execute_script("arguments[0].click();", driver.find_element_by_xpath('//a[@role="button"][@class="button form-button"]'))
        
    except (TimeoutException, WebDriverException):
        return Exception("Ocurrio un error con la pagina")


def login_appleID(driver, orden, tiempo_espera):
    
    procceded_login = False
    intentos = 0
    while not procceded_login:
        try:
            login_ready = EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//iframe[@id="aid-auth-widget-iFrame"]'))
            WebDriverWait(driver, tiempo_espera).until(login_ready)
            procceded_login = True  
        except (NoSuchWindowException, WebDriverException):
            procceded_login = True
            raise Exception("Se cerro la ventana del navegador")                  
        except TimeoutException: 
            intentos+=1 
            if intentos > 3:
                procceded_login = True
                raise Exception ("No cargo el elemento") 
    
    WebDriverWait(driver, tiempo_espera).until(EC.visibility_of_element_located((By.XPATH, '//input[@id="account_name_text_field"]')))
    driver.find_element_by_xpath('//input[@id="account_name_text_field"]').send_keys(orden.email)
    driver.find_element_by_xpath('//input[@id="account_name_text_field"]').send_keys(Keys.ENTER)

    WebDriverWait(driver, tiempo_espera).until(EC.visibility_of_element_located((By.XPATH, '//input[@id="password_text_field"]')))
    driver.find_element_by_xpath('//input[@id="password_text_field"]').send_keys(orden.password)
    driver.find_element_by_xpath('//input[@id="password_text_field"]').send_keys(Keys.ENTER)


def seleccion_invoice(driver, orden, tiempo_espera):

    WebDriverWait(driver, tiempo_espera).until(EC.title_is("Order Details - Apple"))

    
    WebDriverWait(driver, tiempo_espera).until(EC.visibility_of_element_located((By.XPATH, '//a[@data-metkey="viewinvoice"][@class="icon icon-after more"]')))
    driver.execute_script("arguments[0].click();", driver.find_element_by_xpath('//a[@data-metkey="viewinvoice"][@class="icon icon-after more"]'))

    original_window = driver.current_window_handle

    WebDriverWait(driver, tiempo_espera).until(EC.number_of_windows_to_be(2))

    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            break
    
    # Se obtiene informacion de la init_data de la pagina
    serial.info_serial(driver,orden)

    driver.execute_script('document.title="{}";'.format(f'{orden.orden["nombre"]}-invoice duplicate'))

    WebDriverWait(driver, tiempo_espera).until(EC.visibility_of_element_located((By.XPATH, '//button[@type="button"][@class="button rs-invoice-print"]')))
    driver.execute_script("arguments[0].click();", driver.find_element_by_xpath('//button[@type="button"][@class="button rs-invoice-print"]'))

    driver.close()

    orden.estado = True

    print(f'{orden.orden["nombre"]}-invoice duplicate.pdf. Descargado')
