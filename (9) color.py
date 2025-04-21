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
    
    # Validar longitud y formatear la IP
    if len(ip) == 12:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:9]}.{ip[9:]}"
    elif len(ip) == 11:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:9]}.{ip[9:]}"
    elif len(ip) == 10:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:8]}.{ip[8:]}"
    elif len(ip) == 9:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:8]}.{ip[8:]}"
    elif len(ip) == 8:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:]}.0"
    elif len(ip) == 7:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:]}.0"
    else:
        return ip

def clean_percentage(value):
    try:
        # Si el valor es una cadena, quita el símbolo '%' y convierte a float
        if isinstance(value, str):
            value = value.replace('%', '').strip()
        return float(value)
    except ValueError:
        return None

def fetch_data_from_url(ip, options):
    url = f"http://{ip}" if ip else None
    if not url:
        return {"IP": ip, "Toner Negro": "", "UI Negro": "" ,"Toner Cian": "", "UI Cian": "" ,"Toner Magenta": "", "UI Magenta": "" ,"Toner Amarillo": "", "UI Amarillo": "" ,'Estado': '', 'Marca de Tiempo': ""}
    
    print(f"Procesando URL: {url}")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        driver.get(url)
        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ruifw_MainFrm")))

        toner_negro = driver.find_element(By.XPATH, "//tr[@id='1']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
        ui_negro = driver.find_element(By.XPATH, "(//tr[@id='1']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text

        toner_cian = driver.find_element(By.XPATH, "//tr[@id='2']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
        ui_cian = driver.find_element(By.XPATH, "(//tr[@id='2']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text

        toner_magenta = driver.find_element(By.XPATH, "//tr[@id='3']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
        ui_magenta = driver.find_element(By.XPATH, "(//tr[@id='3']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text

        toner_amarillo = driver.find_element(By.XPATH, "//tr[@id='4']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
        ui_amarillo = driver.find_element(By.XPATH, "(//tr[@id='4']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text

        return {
            "IP": ip,
            "Toner Negro": clean_percentage(toner_negro),
            "UI Negro": clean_percentage(ui_negro),
            "Toner Cian": clean_percentage(toner_cian),
            "UI Cian": clean_percentage(ui_cian),
            "Toner Magenta": clean_percentage(toner_magenta),
            "UI Magenta": clean_percentage(ui_magenta),
            "Toner Amarillo": clean_percentage(toner_amarillo),
            "UI Amarillo": clean_percentage(ui_amarillo),
            'Estado': 'OK',
            'Marca de Tiempo': timestamp
        }
    except (NoSuchElementException, TimeoutException):
        return {"IP": ip, "Toner Negro": "", "UI Negro": "" ,"Toner Cian": "", "UI Cian": "","Toner Magenta": "", "UI Magenta": "","Toner Amarillo": "", "UI Amarillo": "" , 'Estado': 'No Disponible', 'Marca de Tiempo': timestamp}
    except WebDriverException:
        return {"IP": ip, "Toner Negro": "", "UI Negro": "" ,"Toner Cian": "", "UI Cian": "","Toner Magenta": "", "UI Magenta": "","Toner Amarillo": "", "UI Amarillo": "" , 'Estado': 'Fuera de Red', 'Marca de Tiempo': timestamp}
    finally:
        driver.quit()

if __name__ == "__main__":
    input_file = r'C:\Users\jvargas\Desktop\Impresoras-color.xlsx'
    output_file = input_file

    df_original = pd.read_excel(input_file)
    df_original['IP'] = df_original['IP'].astype(str).apply(lambda x: format_ip(x) if pd.notna(x) else x)
    df_filtered = df_original[df_original['IP'].notna()]

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_ip = {executor.submit(fetch_data_from_url, ip, options): ip for ip in df_filtered['IP']}
        results = [future.result() for future in as_completed(future_to_ip)]

    df_results = pd.DataFrame(results)

    df_updated = df_original.merge(df_results, on='IP', how='left', suffixes=('', '_new'))

    # Actualizar solo si el estado es 'OK'
    mask_ok = df_updated['Estado_new'] == 'OK'
    
    # Actualizar las columnas con los datos nuevos solo si el estado es 'OK'
    for col in ['Toner Negro', 'UI Negro', 'Toner Cian', 'UI Cian', 'Toner Magenta', 'UI Magenta', 'Toner Amarillo', 'UI Amarillo']:
        df_updated.loc[mask_ok, col] = df_updated.loc[mask_ok, f'{col}_new'].apply(clean_percentage)
    
    # Actualizar Estado y Marca de Tiempo para todos los casos
    df_updated['Estado'] = df_updated['Estado_new'].fillna(df_updated['Estado'])
    df_updated['Marca de Tiempo'] = df_updated['Marca de Tiempo_new'].fillna(df_updated['Marca de Tiempo'])

    # Eliminar columnas auxiliares si existen
    columns_to_drop = [f'{col}_new' for col in ['Toner Negro', 'UI Negro', 'Toner Cian', 'UI Cian', 'Toner Magenta', 'UI Magenta', 'Toner Amarillo', 'UI Amarillo', 'Estado', 'Marca de Tiempo']]
    columns_to_drop = [col for col in columns_to_drop if col in df_updated.columns]
    df_updated = df_updated.drop(columns=columns_to_drop)

    df_updated.to_excel(output_file, index=False)
