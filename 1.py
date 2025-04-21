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
from openpyxl import load_workbook
from openpyxl.styles import Font

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
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:]}"
    elif len(ip) == 7:
        return f"{ip[:3]}.{ip[3:6]}.{ip[6:]}"
    else:
        return ip


def get_column_widths(ws):
    widths = {}
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        widths[column] = max_length + 2  # Adding a bit more space
    return widths


def preserve_formulas_and_formats(input_file):
    wb = load_workbook(input_file, data_only=False)
    formulas = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_formulas = {cell.coordinate: cell.formula if cell.data_type ==
                          'f' else cell.value for row in ws.iter_rows() for cell in row}
        formulas[sheet_name] = sheet_formulas
    return formulas


def apply_formulas_and_formats(output_file, formulas):
    wb = load_workbook(output_file, data_only=False)
    for sheet_name, sheet_formulas in formulas.items():
        ws = wb[sheet_name]
        for cell_coord, formula in sheet_formulas.items():
            cell = ws[cell_coord]
            if formula is not None:
                if isinstance(formula, str) and formula.startswith('='):
                    cell.formula = formula
                else:
                    cell.value = formula
    wb.save(output_file)


def procesar_impresoras_normales(file_path):

    def fetch_data_from_url(ip):
        url = f"http://{ip}" if ip else None
        if not url:
            return {"IP": ip, "Toner Negro": "", "UI Negro": "", 'Estado': '', 'Marca de Tiempo': ""}

        print(f"Procesando URL: {url}")
        driver = webdriver.Chrome(service=Service(
            ChromeDriverManager().install()), options=options)
        try:
            driver.get(url)
            WebDriverWait(driver, 5).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ruifw_MainFrm")))

            # Intentar obtener el valor del contador negro y blanco
            try:
                black_and_white_counter = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "table#toner_list td.tonervalue_number"))
                ).text
            except (NoSuchElementException, TimeoutException):
                try:
                    # Buscar el segundo fallback: el <td> que contiene "0%" directamente
                    fallback_element = WebDriverWait(driver, 3).until(
                        EC.visibility_of_element_located(
                            (By.XPATH,
                             '//table[@width="100%" and @border="0"]//td[contains(text(), "0%")]')
                        )
                    )
                    black_and_white_counter = fallback_element.text
                except (NoSuchElementException, TimeoutException):
                    black_and_white_counter = "0%"

            # Intentar obtener el valor del contador de color
            try:
                color_counter = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "table#imagine_list td.tonervalue_number"))
                ).text
            except (NoSuchElementException, TimeoutException):
                try:
                    # Buscar el segundo fallback: el <td> que contiene "0%" directamente
                    fallback_element = WebDriverWait(driver, 3).until(
                        EC.visibility_of_element_located(
                            (By.XPATH,
                             '//table[@width="100%" and @border="0"]//td[contains(text(), "0%")]')
                        )
                    )
                    color_counter = fallback_element.text
                except (NoSuchElementException, TimeoutException):
                    color_counter = "0%"

            # Determinar el estado basado en los contadores
            estado = 'OK'
            if color_counter == "0%":
                estado = 'Sin UI'
            if black_and_white_counter == "0%":
                estado = 'Sin Toner'
            if color_counter == "0%" and black_and_white_counter == "0%":
                estado = 'Sin UI y Sin Toner'

            return {"IP": ip, "Toner Negro": black_and_white_counter, "UI Negro": color_counter, 'Estado': estado, 'Marca de Tiempo': timestamp}
        except (NoSuchElementException, TimeoutException):
            return {"IP": ip, "Toner Negro": "", "UI Negro": "", 'Estado': 'No Disponible', 'Marca de Tiempo': timestamp}
        except WebDriverException:
            return {"IP": ip, "Toner Negro": "", "UI Negro": "", 'Estado': 'Fuera de Red', 'Marca de Tiempo': timestamp}
        finally:
            driver.quit()

    def convert_to_percentage(value):
        if pd.isna(value) or value is None:
            return ""  # Deja la celda vacía
        elif isinstance(value, (int, float)):
            return f"{int(value * 100)}%"
        elif isinstance(value, str):
            try:
                # Convertir de cadena a número, manejando coma como punto decimal
                number = float(value.replace(',', '.'))
                return f"{int(number * 100)}%"
            except ValueError:
                return value
        else:
            return value

    # Leer solo la hoja "Impresoras Normales"
    df_original = pd.read_excel(file_path, sheet_name='Impresoras Normales')

    # Formatear las IPs
    df_original['IP'] = df_original['IP'].astype(str).apply(
        lambda x: format_ip(x) if pd.notna(x) else x)
    df_filtered = df_original[df_original['IP'].notna()]

    # Configuración de Selenium
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    # Ejecutar la función en paralelo
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_ip = {executor.submit(
            fetch_data_from_url, ip): ip for ip in df_filtered['IP']}
        results = [future.result() for future in as_completed(future_to_ip)]

    # Crear un DataFrame con los resultados obtenidos
    df_results = pd.DataFrame(results)

    # Combinar los resultados con el DataFrame original
    df_updated = df_original.merge(
        df_results, on='IP', how='left', suffixes=('', '_new'))

    # Asegurarse de que las columnas estén en el tipo correcto antes de asignar
    df_updated['Toner Negro'] = df_updated['Toner Negro'].fillna(
        '').astype(str)
    df_updated['UI Negro'] = df_updated['UI Negro'].fillna('').astype(str)

    # Actualizar 'Toner Negro' y 'UI Negro' solo si el estado es 'OK'
    mask_ok = df_updated['Estado_new'] == 'OK'
    df_updated.loc[mask_ok,
                   'Toner Negro'] = df_updated.loc[mask_ok, 'Toner Negro_new']
    df_updated.loc[mask_ok,
                   'UI Negro'] = df_updated.loc[mask_ok, 'UI Negro_new']

    # Actualizar 'Estado' y 'Marca de Tiempo' para todos los casos
    df_updated['Estado'] = df_updated['Estado_new'].fillna(
        df_updated['Estado'])
    df_updated['Marca de Tiempo'] = df_updated['Marca de Tiempo_new'].fillna(
        df_updated['Marca de Tiempo'])

    # Eliminar columnas auxiliares
    columns_to_drop = ['Toner Negro_new', 'UI Negro_new',
                       'Estado_new', 'Marca de Tiempo_new']
    df_updated = df_updated.drop(columns=columns_to_drop)

    # Convertir valores decimales a porcentajes
    df_updated['Toner Negro'] = df_updated['Toner Negro'].apply(
        convert_to_percentage)
    df_updated['UI Negro'] = df_updated['UI Negro'].apply(
        convert_to_percentage)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_updated.to_excel(
            writer, sheet_name='Impresoras Normales', index=False)

    # Aplicar fórmulas y formatos preservados
    formulas = preserve_formulas_and_formats(file_path)
    apply_formulas_and_formats(file_path, formulas)


def procesar_impresoras_colores(file_path, output_file):

    def clean_percentage(value: str) -> str:
        try:
            if isinstance(value, str):
                value = value.replace('%', '').strip()
            return f"{int(round(float(value)))}%"
        except ValueError:
            return ""

    def fetch_data_from_url(ip: str, options: webdriver.ChromeOptions) -> dict:
        url = f"http://{ip}" if ip else None
        if not url:
            return {"IP": ip, "Toner Negro": "", "UI Negro": "", "Toner Cian": "", "UI Cian": "", "Toner Magenta": "", "UI Magenta": "", "Toner Amarillo": "", "UI Amarillo": "", 'Estado': '', 'Marca de Tiempo': ""}

        print(f"Procesando URL: {url}")
        driver = webdriver.Chrome(service=Service(
            ChromeDriverManager().install()), options=options)
        try:
            driver.get(url)
            WebDriverWait(driver, 5).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ruifw_MainFrm")))

            data = {}
            for idx, color in enumerate(['Negro', 'Cian', 'Magenta', 'Amarillo'], start=1):
                data[f"Toner {color}"] = driver.find_element(
                    By.XPATH, f"//tr[@id='{idx}']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]").text
                data[f"UI {color}"] = driver.find_element(
                    By.XPATH, f"(//tr[@id='{idx}']/td[2]/table/tbody/tr/td/table/tbody/tr/td[2])[2]").text

            return {**{"IP": ip, 'Estado': 'OK', 'Marca de Tiempo': timestamp}, **{key: clean_percentage(value) for key, value in data.items()}}
        except (NoSuchElementException, TimeoutException):
            return {"IP": ip, "Toner Negro": "", "UI Negro": "", "Toner Cian": "", "UI Cian": "", "Toner Magenta": "", "UI Magenta": "", "Toner Amarillo": "", "UI Amarillo": "", 'Estado': 'No Disponible', 'Marca de Tiempo': timestamp}
        except WebDriverException:
            return {"IP": ip, "Toner Negro": "", "UI Negro": "", "Toner Cian": "", "UI Cian": "", "Toner Magenta": "", "UI Magenta": "", "Toner Amarillo": "", "UI Amarillo": "", 'Estado': 'Fuera de Red', 'Marca de Tiempo': timestamp}
        finally:
            driver.quit()

    # Leer las hojas del archivo Excel
    sheets = pd.read_excel(file_path, sheet_name=None)
    wb = load_workbook(file_path)
    column_widths = {sheet_name: get_column_widths(
        wb[sheet_name]) for sheet_name in wb.sheetnames}

    # Procesar la hoja 'Impresoras a Color'
    df_original = sheets['Impresoras a Color']
    df_original['IP'] = df_original['IP'].astype(str).apply(format_ip)
    df_filtered = df_original[df_original['IP'].notna()]

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_ip = {executor.submit(
            fetch_data_from_url, ip, options): ip for ip in df_filtered['IP']}
        results = [future.result() for future in as_completed(future_to_ip)]

    df_results = pd.DataFrame(results)
    df_updated = df_original.merge(
        df_results, on='IP', how='left', suffixes=('', '_new'))

    mask_ok = df_updated['Estado_new'] == 'OK'
    columns = ['Toner Negro', 'UI Negro', 'Toner Cian', 'UI Cian',
               'Toner Magenta', 'UI Magenta', 'Toner Amarillo', 'UI Amarillo']
    for col in columns:
        df_updated.loc[mask_ok, col] = df_updated.loc[mask_ok,
                                                      f'{col}_new'].apply(clean_percentage)

    df_updated['Estado'] = df_updated['Estado_new'].fillna(
        df_updated['Estado'])
    df_updated['Marca de Tiempo'] = df_updated['Marca de Tiempo_new'].fillna(
        df_updated['Marca de Tiempo'])

    columns_to_drop = [f'{col}_new' for col in columns +
                       ['Estado', 'Marca de Tiempo']]
    df_updated.drop(columns=[
                    col for col in columns_to_drop if col in df_updated.columns], inplace=True)

    # Guardar el DataFrame actualizado
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_updated.to_excel(
            writer, sheet_name='Impresoras a Color', index=False)
        for sheet_name, df_sheet in sheets.items():
            if sheet_name != 'Impresoras a Color':
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

    # Aplicar fórmulas y formatos preservados
    formulas = preserve_formulas_and_formats(file_path)
    apply_formulas_and_formats(output_file, formulas)


def procesar_impresoras_colores_clx(file_path, output_file):

    def clean_percentage(value: str) -> str:
        try:
            if isinstance(value, str):
                value = value.replace('%', '').strip()
            return f"{int(round(float(value)))}%"
        except ValueError:
            return ""

    def fetch_data_from_url(ip, options):
        url = f"http://{ip}" if ip else None
        if not url:
            return {"IP": ip, "Toner Negro": "", "Toner Cian": "", "Toner Magenta": "", "Toner Amarillo": "", 'Estado': '', 'Marca de Tiempo': ""}

        print(f"Procesando URL: {url}")
        driver = webdriver.Chrome(service=Service(
            ChromeDriverManager().install()), options=options)

        try:

            driver.get(url)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, ".x-grid3-row:nth-child(1) .x-column:nth-child(2)"))
            )

            toner_negro = driver.find_element(
                By.CSS_SELECTOR, ".x-grid3-row:nth-child(1) .x-column:nth-child(2)").text
            # print(f"Toner Negro encontrado: {toner_negro}")

            toner_cian = driver.find_element(
                By.CSS_SELECTOR, ".x-grid3-row:nth-child(2) .x-column:nth-child(2)").text
            # print(f"Toner Cian encontrado: {toner_cian}")

            toner_magenta = driver.find_element(
                By.CSS_SELECTOR, ".x-grid3-row:nth-child(3) .x-column:nth-child(2)").text
            # print(f"Toner Magenta encontrado: {toner_magenta}")

            toner_amarillo = driver.find_element(
                By.CSS_SELECTOR, ".x-grid3-row:nth-child(4) .x-column:nth-child(2)").text
            # print(f"Toner Amarillo encontrado: {toner_amarillo}")

            return {
                "IP": ip,  # Asegurando que 'IP' esté en el resultado
                "Toner Negro": toner_negro,
                "Toner Cian": toner_cian,
                "Toner Magenta": toner_magenta,
                "Toner Amarillo": toner_amarillo,
                'Estado': 'OK' if toner_negro or toner_cian or toner_magenta or toner_amarillo else 'No disponible',
                'Marca de Tiempo': timestamp
            }
        except (NoSuchElementException):
            return {"IP": ip, "Toner Negro": "", "Toner Cian": "", "Toner Magenta": "", "Toner Amarillo": "", 'Estado': 'No Disponible', 'Marca de Tiempo': timestamp}
        except TimeoutException:
            return {"IP": ip, "Toner Negro": "", "Toner Cian": "", "Toner Magenta": "", "Toner Amarillo": "", 'Estado': 'No Disponible', 'Marca de Tiempo': timestamp}
        except WebDriverException:
            print(f"Timeout al intentar conectar con {url}")
            return {"IP": ip, "Toner Negro": "", "Toner Cian": "", "Toner Magenta": "", "Toner Amarillo": "", 'Estado': 'Fuera de Red', 'Marca de Tiempo': timestamp}
        finally:
            driver.quit()

    # Leer las hojas del archivo Excel
    sheets = pd.read_excel(file_path, sheet_name=None)
    wb = load_workbook(file_path)
    column_widths = {sheet_name: get_column_widths(
        wb[sheet_name]) for sheet_name in wb.sheetnames}

    # Procesar la hoja 'Impresoras a Color'
    df_original = sheets['Impresora CLX-6260']
    df_original['IP'] = df_original['IP'].astype(str).apply(format_ip)
    df_filtered = df_original[df_original['IP'].notna()]

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_ip = {executor.submit(
            fetch_data_from_url, ip, options): ip for ip in df_filtered['IP']}
        results = [future.result() for future in as_completed(future_to_ip)]

    df_results = pd.DataFrame(results)

    # Verificar que las columnas 'IP' existan antes de hacer el merge
    if 'IP' not in df_original.columns or 'IP' not in df_results.columns:
        raise KeyError("'IP' column is missing in one of the DataFrames.")

    # Fusionar los resultados
    df_updated = df_original.merge(
        df_results, on='IP', how='left', suffixes=('', '_new')
    )

    # Restablecer columnas NaN
    columns = ['Toner Negro', 'Toner Cian', 'Toner Magenta',
               'Toner Amarillo', 'Estado', 'Marca de Tiempo']
    df_updated[columns] = df_updated[columns].fillna('')

    mask_ok = df_updated['Estado_new'] == 'OK'
    columns = ['Toner Negro', 'Toner Cian',
               'Toner Magenta', 'Toner Amarillo']
    for col in columns:
        df_updated[col] = df_updated[col].astype(str)
        df_updated.loc[mask_ok, col] = df_updated.loc[mask_ok,
                                                      f'{col}_new'].apply(clean_percentage)

    df_updated['Estado'] = df_updated['Estado_new'].fillna(
        df_updated['Estado'])
    df_updated['Marca de Tiempo'] = df_updated['Marca de Tiempo_new'].fillna(
        df_updated['Marca de Tiempo'])

    columns_to_drop = [f'{col}_new' for col in columns +
                       ['Estado', 'Marca de Tiempo']]
    df_updated.drop(columns=[
                    col for col in columns_to_drop if col in df_updated.columns], inplace=True)

    # Guardar el DataFrame actualizado
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_updated.to_excel(
            writer, sheet_name='Impresora CLX-6260', index=False)
        for sheet_name, df_sheet in sheets.items():
            if sheet_name != 'Impresora CLX-6260':
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

    # Aplicar fórmulas y formatos preservados
    formulas = preserve_formulas_and_formats(file_path)
    apply_formulas_and_formats(output_file, formulas)


def format_excel_sheets(file_path):
    wb = load_workbook(file_path)
    red_font = Font(color="FF0000")
    orange_font = Font(color="ff6f00")

    print("Aplicando formato a todas las hojas:")
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"Procesando hoja: {sheet_name}")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell_value = str(cell.value).strip()
              # Comprobar si el valor es '0%' o '0' para aplicar el texto rojo
                if cell_value in ['0%', '0']:
                    cell.font = red_font
                    print(
                        f"Formato aplicado a celda: {cell.coordinate}, Valor: '{cell_value}' en hoja {sheet_name}")

                # Comprobar si el valor es menor al 5% para aplicar el texto naranja
                elif cell_value.endswith('%') and float(cell_value[:-1]) < 5:
                    cell.font = orange_font
                    print(
                        f"Formato aplicado a celda: {cell.coordinate}, Valor: '{cell_value}' en hoja {sheet_name}")

        # Ajustar el ancho de las columnas para cada hoja
        column_widths = get_column_widths(ws)
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

    # Asegurar que "Impresoras Normales" esté al principio
    if "Impresoras Normales" in wb.sheetnames:
        wb.move_sheet("Impresoras Normales", offset=-
                      wb.index(wb["Impresoras Normales"]))

    wb.save(file_path)
    print("Formato aplicado y archivo guardado.")


input_file = r'G:\Unidades compartidas\Informática\Impresoras - final.xlsx'

# procesar_impresoras_colores_clx(input_file, input_file)
# procesar_impresoras_colores(input_file, input_file)
procesar_impresoras_normales(input_file)
format_excel_sheets(input_file)
