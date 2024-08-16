from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException

# Configuración del WebDriver
driver = webdriver.Chrome()

try:
    # Navegar a la página web
    driver.get("http://192.168.111.210/sws/index.sws")
    
    try:
        # Cambiar al iframe usando el nombre
        driver.switch_to.frame("ruifw_MainFrm")
        
        try:
            # Encontrar el elemento y obtener el texto
            black_and_white_counter = driver.find_element(By.CSS_SELECTOR, "td.tonervalue_number").text
            color_counter = driver.find_element(By.CSS_SELECTOR, "table#imagine_list td.tonervalue_number").text
            
            # Imprimir los resultados
            print("Toner Restante:", black_and_white_counter)
            print("Unidad de Imagen Restante:", color_counter)
        
        except NoSuchElementException:
            # Manejar el caso en que los elementos no se encuentren
            print("No se encontraron los elementos en el iframe.")
        
    except NoSuchElementException:
        # Manejar el caso en que el `iframe` no se pueda encontrar
        print("No se pudo encontrar el iframe con el nombre 'ruifw_MainFrm'.")
    except WebDriverException as e:
        # Manejar otros problemas con el WebDriver
        print("Error con el WebDriver:", e)
    
except (TimeoutException, WebDriverException) as e:
    # Manejar errores de acceso a la página (p. ej., si la página no responde)
    print("Impresora Fuera de Red")

finally:
    # Cerrar el navegador
    driver.quit()
