import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Función para formatear direcciones IP
def format_ip(ip):
    # Eliminar cualquier carácter que no sea dígito
    ip = re.sub(r'\D', '', ip)
    
    # Asegurarse de que la longitud de IP sea correcta
    if len(ip) == 12:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:9]}.{ip[9:]}"
    elif len(ip) == 11:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:9]}.{ip[9:]}"
    elif len(ip) == 10:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:8]}.{ip[8:]}"
    elif len(ip) == 9:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:8]}.{ip[8:]}"
    elif len(ip) == 8:
            return f"{ip[:3]}.{ip[3:6]}.{ip[6:]}.."
    elif len(ip) == 7:
            return f"{ip[:3]}.{ip[3:6]}.{ip[6:]}."
    else:
            # Si la longitud es menor que 7, no tiene suficiente información para formatear
            return ip  # Retorna IP original si no cumple con las reglas

# Leer el archivo Excel
input_file = r'C:\Users\jvargas\Desktop\Impresoras.xlsx'
output_file = input_file  # Sobrescribir el archivo original

df_original = pd.read_excel(input_file)

# Limpiar y formatear datos IP
df_original['IP'] = df_original['IP'].astype(str).apply(format_ip)

# Convertir IP a URL
def ip_to_url(ip):
    return f"http://{ip}"

# Opciones de Chrome para modo headless
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def fetch_data_from_url(ip):
    url = ip_to_url(ip)
    print(f"Procesando URL: {url}")  # Añade un registro para URLs que se están procesando
    try:
        driver.get(url)
        driver.switch_to.frame("ruifw_MainFrm")
        black_and_white_counter = driver.find_element(By.CSS_SELECTOR, "table#toner_list td.tonervalue_number").text
        color_counter = driver.find_element(By.CSS_SELECTOR, "table#imagine_list td.tonervalue_number").text
        return {"IP": ip, "Toner Restante": black_and_white_counter, "Unidad de Imagen Restante": color_counter}
    except NoSuchElementException:
        return {"IP": ip, "Toner Restante": "Error", "Unidad de Imagen Restante": "Error"}
    except (TimeoutException, WebDriverException) as e:
        return {"IP": ip, "Toner Restante": "Fuera de Red", "Unidad de Imagen Restante": f"Fuera de Red"}

# Obtener resultados para cada IP
results = [fetch_data_from_url(ip) for ip in df_original['IP']]
df_results = pd.DataFrame(results)

# Fusionar datos actualizados en el DataFrame original
df_updated = df_original.merge(df_results, on='IP', how='left', suffixes=('', '_new'))

# Actualizar las columnas existentes
df_updated['Toner Restante'] = df_updated['Toner Restante_new']
df_updated['Unidad de Imagen Restante'] = df_updated['Unidad de Imagen Restante_new']

# Eliminar columnas auxiliares
df_updated = df_updated.drop(columns=['Toner Restante_new', 'Unidad de Imagen Restante_new'])

# Sobrescribir el archivo original con los resultados actualizados
df_updated.to_excel(output_file, index=False)

# Cerrar el driver
driver.quit()
