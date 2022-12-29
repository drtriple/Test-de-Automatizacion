#enconding: utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver. support.ui import WebDriverWait
from openpyxl import load_workbook
import time


driver = webdriver.Chrome("./D:\Datos\Desktop\BANCOLOMBIA PRUEBAS\Test de Conocimiento en Automatización\P1\chromedriver.exe")
driver.get("https://docs.google.com/forms/d/e/1FAIpQLSd8pYrym78Am_OtC7TeJ7igtixsW7eZPbRCAM6vbii3nS-0cA/viewform")
time.sleep(3)

#inputs = driver.find_elements_by_class_name("whsOnd")
#time.sleep(1)

filesheet = "./DatosPruebaPy.xlsx"
wb = load_workbook(filesheet)
hojas = wb.get_sheet_names()
print(hojas)
valores = wb.get_sheet_by_name('DatosPruebaPy')
wb.close()

for i in range(2,1001):
    identificador, Producto, Vendedor, Cantidad, Ubicacion, Categoria, PGanancia  = valores[f'A{i}:G{i}'][0]
    print(identificador.value, Producto.value, Vendedor.value, Cantidad.value, Ubicacion.value, Categoria.value, PGanancia.value)
    time.sleep(1)

    ## Nombre el Aspirante*
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys("Juan José Bedoya")
    ## Identificador
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(identificador.value)
    ## Producto
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Producto.value)
   ## Vendedor
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Vendedor.value)
    ## Cantidad
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Cantidad.value)
    ## Ubicación
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Ubicacion.value)
   ## Categoria
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Categoria.value)
    if PGanancia.value == None:
        PGanancia.value = "0"    
        ## Porcentaje de Ganancia
        driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[8]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(PGanancia.value)
    else:
        driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[2]/div[8]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(PGanancia.value)             
    
    ## enviar encuesta
    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()
    ##ENVIAR OTRO REGISTRO
    another_response = driver.find_element('xpath','/html/body/div[1]/div[2]/div[1]/div/div[4]/a')

    another_response.click()
    # CERRAR VENTANA
driver.close()