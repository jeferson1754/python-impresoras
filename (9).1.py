from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración del navegador
chrome_options = Options()
chrome_options.add_argument('--headless')  # Ejecutar en modo headless
chrome_options.add_argument('--disable-gpu')  # Deshabilitar GPU

# Inicializar el navegador
driver = webdriver.Chrome(service=Service(), options=chrome_options)

try:
    # Navegar a la URL
    driver.get("http://192.168.20.6/sws/index.sws")
    
    # Ajustar tamaño de ventana
    driver.set_window_size(1552, 832)
    
    # Cambiar al frame correcto
    driver.switch_to.frame(1)
    
    # Esperar hasta que los elementos estén presentes
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//tr[@id='1']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]"))
    )
    
    # Extraer información de los toners
    toner_negro = driver.find_element(By.XPATH, "//tr[@id='1']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
    toner_cian = driver.find_element(By.XPATH, "//tr[@id='2']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
    toner_magenta = driver.find_element(By.XPATH, "//tr[@id='3']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
    toner_amarillo = driver.find_element(By.XPATH, "//tr[@id='4']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
    
    # Extraer información de las unidades de imagen
    ui_negro = driver.find_element(By.XPATH, "(//tr[@id='1']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text
    ui_cian = driver.find_element(By.XPATH, "(//tr[@id='2']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text
    ui_magenta = driver.find_element(By.XPATH, "(//tr[@id='3']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text
    ui_amarillo = driver.find_element(By.XPATH, "(//tr[@id='4']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text
    
    # Imprimir los resultados
    print(f"Toner Negro Restante: {toner_negro}")
    print(f"Toner Cian Restante: {toner_cian}")
    print(f"Toner Magenta Restante: {toner_magenta}")
    print(f"Toner Amarillo Restante: {toner_amarillo}")
    print(f"Unidad de Imagen Negro Restante: {ui_negro}")
    print(f"Unidad de Imagen Cian Restante: {ui_cian}")
    print(f"Unidad de Imagen Magenta Restante: {ui_magenta}")
    print(f"Unidad de Imagen Amarillo Restante: {ui_amarillo}")

except Exception as e:
    print(f"Error: {e}")

finally:
    # Cerrar el navegador
    driver.quit()
