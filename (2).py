from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Lista de URLs o direcciones IP
urls = [
    "http://192.168.111.210/sws/index.sws",
    "http://192.168.111.217/sws/index.sws"  # Añadir más URLs aquí
    # "http://otra-direccion-ip/sws/index.sws",
]

# Configuración del WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

def fetch_data_from_url(url):
    try:
        # Navegar a la página web
        driver.get(url)
        
        try:
            # Cambiar al iframe usando el nombre
            driver.switch_to.frame("ruifw_MainFrm")
            
            try:
                # Encontrar el elemento y obtener el texto
                black_and_white_counter = driver.find_element(By.CSS_SELECTOR, "td.tonervalue_number").text
                color_counter = driver.find_element(By.CSS_SELECTOR, "table#imagine_list td.tonervalue_number").text
                
                # Imprimir los resultados
                print(f"URL: {url}")
                print("Toner Restante:", black_and_white_counter)
                print("Unidad de Imagen Restante:", color_counter)
                print()  # Línea en blanco para mejor legibilidad
            
            except NoSuchElementException:
                # Manejar el caso en que los elementos no se encuentren
                print(f"URL: {url}")
                print("No se encontraron los elementos en el iframe.")
                print()
            
        except NoSuchElementException:
            # Manejar el caso en que el `iframe` no se pueda encontrar
            print(f"URL: {url}")
            print("No se pudo encontrar el iframe con el nombre 'ruifw_MainFrm'.")
            print()
        except WebDriverException as e:
            # Manejar otros problemas con el WebDriver
            print(f"URL: {url}")
            print("Error con el WebDriver:", e)
            print()
        
    except (TimeoutException, WebDriverException) as e:
        # Manejar errores de acceso a la página (p. ej., si la página no responde)
        print(f"URL: {url}")
        print("Impresora Fuera de Red:", e)
        print()

# Iterar sobre todas las URLs
for url in urls:
    fetch_data_from_url(url)

# Cerrar el navegador
driver.quit()
