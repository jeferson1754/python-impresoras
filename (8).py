import pandas as pd
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# Obtener la fecha y hora actual
timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def format_ip(ip):
    if pd.isna(ip) or not ip.strip():
        return None  # Retorna None si la IP está vacía o solo tiene espacios

    ip = re.sub(r'\D', '', ip)
    
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
        return ip
    
def fetch_data_from_url(ip):
    url = f"http://{ip}" if ip else None
    if not url:
        return {"IP": ip, "Toner Restante": "", "Unidad de Imagen Restante": "", 'Estado': '', 'Marca de Tiempo': ""}

    print(f"Procesando URL: {url}")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        driver.get(url)
        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ruifw_MainFrm")))
        black_and_white_counter = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "table#toner_list td.tonervalue_number"))
        ).text
        color_counter = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "table#imagine_list td.tonervalue_number"))
        ).text
        return {"IP": ip, "Toner Restante": black_and_white_counter, "Unidad de Imagen Restante": color_counter, 'Estado': 'OK', 'Marca de Tiempo': timestamp}
    except (NoSuchElementException, TimeoutException):
        return {"IP": ip, "Toner Restante": "", "Unidad de Imagen Restante": "", 'Estado': 'No Disponible', 'Marca de Tiempo': timestamp}
    except WebDriverException:
        return {"IP": ip, "Toner Restante": "", "Unidad de Imagen Restante": "", 'Estado': 'Fuera de Red', 'Marca de Tiempo': timestamp}
    finally:
        driver.quit()

if __name__ == "__main__":
    input_file = r'C:\Users\jvargas\Desktop\Impresoras.xlsx'
    output_file = input_file

    df_original = pd.read_excel(input_file)
    df_original['IP'] = df_original['IP'].astype(str).apply(lambda x: format_ip(x) if pd.notna(x) else x)
    df_filtered = df_original[df_original['IP'].notna()]

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_ip = {executor.submit(fetch_data_from_url, ip): ip for ip in df_filtered['IP']}
        results = [future.result() for future in as_completed(future_to_ip)]

    df_results = pd.DataFrame(results)
    df_updated = df_original.merge(df_results, on='IP', how='left', suffixes=('', '_new'))

    columns_to_update = ['Toner Restante', 'Unidad de Imagen Restante', 'Estado', 'Marca de Tiempo']
    for col in columns_to_update:
        df_updated[col] = df_updated[f'{col}_new'].fillna(df_updated[col])
    
    df_updated = df_updated.drop(columns=[f'{col}_new' for col in columns_to_update])
    df_updated.to_excel(output_file, index=False)