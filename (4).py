import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor
from webdriver_manager.chrome import ChromeDriverManager

# Lista de URLs o direcciones IP
urls = [
    "http://192.168.111.210/sws/index.sws",
    "http://192.168.111.217/sws/index.sws"  # Añadir más URLs aquí
]

# Lista para almacenar los resultados
results = []

def fetch_data_from_url(url):
    # Configuración del WebDriver para cada hilo
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Ejecutar en modo headless
    chrome_options.add_argument("--disable-gpu")  # Desactivar GPU para headless
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        # Navegar a la página web
        driver.get(url)
        
        try:
            # Esperar a que el `iframe` esté disponible y cambiar a él
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "ruifw_MainFrm")))
            
            try:
                # Esperar a que los elementos estén presentes
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "td.tonervalue_number")))
                black_and_white_counter = driver.find_element(By.CSS_SELECTOR, "td.tonervalue_number").text
                color_counter = driver.find_element(By.CSS_SELECTOR, "table#imagine_list td.tonervalue_number").text
                
                # Almacenar los resultados
                results.append({
                    "URL": url,
                    "Toner Restante": black_and_white_counter,
                    "Unidad de Imagen Restante": color_counter
                })
            
            except NoSuchElementException:
                # Manejar el caso en que los elementos no se encuentren
                results.append({
                    "URL": url,
                    "Toner Restante": "No se encontraron los elementos en el iframe.",
                    "Unidad de Imagen Restante": "No se encontraron los elementos en el iframe."
                })
            
        except NoSuchElementException:
            # Manejar el caso en que el `iframe` no se pueda encontrar
            results.append({
                "URL": url,
                "Toner Restante": "No se pudo encontrar el iframe con el nombre 'ruifw_MainFrm'.",
                "Unidad de Imagen Restante": "No se pudo encontrar el iframe con el nombre 'ruifw_MainFrm'."
            })
        except WebDriverException as e:
            # Manejar otros problemas con el WebDriver
            results.append({
                "URL": url,
                "Toner Restante": f"Error con el WebDriver: {e}",
                "Unidad de Imagen Restante": f"Error con el WebDriver: {e}"
            })
        
    except (TimeoutException, WebDriverException) as e:
        # Manejar errores de acceso a la página (p. ej., si la página no responde)
        results.append({
            "URL": url,
            "Toner Restante": f"Impresora Fuera de Red: {e}",
            "Unidad de Imagen Restante": f"Impresora Fuera de Red: {e}"
        })
    
    finally:
        driver.quit()

# Usar ThreadPoolExecutor para paralelizar la ejecución
with ThreadPoolExecutor(max_workers=4) as executor:
    executor.map(fetch_data_from_url, urls)

# Crear un DataFrame de pandas a partir de los resultados
df = pd.DataFrame(results)

# Guardar los resultados en un archivo Excel
df.to_excel("resultados.xlsx", index=False)

print("Resultados guardados en 'resultados.xlsx'.")
