import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import json


def opciones_chrome(carpeta_descarga):
    opciones = Options()
    opciones.add_argument("--disable-gpu")
    opciones.add_argument("--log-level=3")
    opciones.add_argument("--disable-popup-blocking")
    opciones.add_experimental_option("excludeSwitches", ['enable-automation','enable-logging'])

    # para usar el print-to-pdf de chrome
    settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
             'savefile.default_directory':carpeta_descarga}
             
    opciones.add_experimental_option('prefs', prefs)
    opciones.add_argument('--kiosk-printing')

    return opciones

def instalar_webdriver():
    os.environ['WDM_LOCAL'] = '1'
    os.environ['WDM_LOG_LEVEL'] = '0'
    executableChrome = ChromeDriverManager().install()
    return executableChrome

def crear_webdriver(executableChrome, carpeta_descarga):
    driver = webdriver.Chrome(executable_path=executableChrome, chrome_options=opciones_chrome(carpeta_descarga))
    return driver
