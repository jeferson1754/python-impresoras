import pandas as pd
import numpy as np
from datetime import timedelta
from sklearn.linear_model import LinearRegression

# --- CONFIGURACIÓN ---
# Archivo con hoja 'Histórico'
INPUT_FILE = r"G:\Unidades compartidas\Informática\Impresoras - final.xlsx"
OUTPUT_FILE = "predicciones_toner.xlsx"

# Consumibles a revisar
TONER_COLUMNS = ["Toner Negro", "Toner Cian",
                 "Toner Magenta", "Toner Amarillo"]
KITS_COLUMNS = ["Kit Mant.", "Kit Alim."]
CONSUMIBLES = TONER_COLUMNS + KITS_COLUMNS

ESTADO_VALIDO = "OK"

# --- CARGA DE DATOS ---
try:
    df = pd.read_excel(INPUT_FILE, sheet_name="Histórico")
except Exception as e:
    raise Exception(f"Error al leer el archivo histórico: {e}")

# Normalizar nombres de columnas y fechas
df.columns = [col.strip() for col in df.columns]
df["Fecha de registro"] = pd.to_datetime(
    df["Marca de Tiempo"], errors="coerce")

# Filtrar registros válidos
df = df[df["Estado"] == ESTADO_VALIDO]

# --- FUNCIÓN DE PREDICCIÓN ---


def predecir_consumible(sub_df, consumible):
    sub_df = sub_df.sort_values("Fecha de registro")
    sub_df = sub_df[["Fecha de registro", consumible]].dropna()
    sub_df[consumible] = pd.to_numeric(sub_df[consumible].astype(
        str).str.replace("%", ""), errors="coerce")
    sub_df = sub_df.dropna()

    if len(sub_df) < 2:
        return np.nan, np.nan, np.nan, np.nan, "❌ Muy pocos datos"

    # Crear eje temporal en días
    sub_df["Días"] = (sub_df["Fecha de registro"] -
                      sub_df["Fecha de registro"].min()).dt.total_seconds() / (24*3600)
    X = sub_df[["Días"]].values
    y = sub_df[consumible].values

    # Elegir método según número de registros
    if len(sub_df) >= 3:
        model = LinearRegression()
        model.fit(X, y)
        consumo_diario = -model.coef_[0]  # Pendiente negativa
        metodo = "📈 Regresión lineal"
    else:
        delta_pct = y[-2] - y[-1]
        delta_days = sub_df["Días"].iloc[-1] - sub_df["Días"].iloc[-2]
        consumo_diario = delta_pct / delta_days if delta_days > 0 else np.nan
        metodo = "⚙️ Promedio simple"

    porcentaje_actual = y[-1]
    if consumo_diario <= 0 or np.isnan(consumo_diario):
        return porcentaje_actual, 0, np.nan, np.nan, metodo

    dias_restantes = porcentaje_actual / consumo_diario
    fecha_agotamiento = sub_df["Fecha de registro"].iloc[-1] + \
        timedelta(days=dias_restantes)

    return round(porcentaje_actual, 1), round(consumo_diario, 2), round(dias_restantes, 1), fecha_agotamiento, metodo


def predecir_consumible_promedio():
# --- GENERAR PREDICCIONES ---
    resultados = []

    for (ip, modelo), grupo in df.groupby(["IP", "Modelo"], dropna=False):
        # Definir nombre si existe columna, si no usar IP
        nombre = grupo["Nombre"].iloc[0] if "Nombre" in grupo.columns else ip

        for consumible in CONSUMIBLES:
            if consumible not in grupo.columns or grupo[consumible].dropna().empty:
                continue
            pct, consumo, dias, fecha_fin, metodo = predecir_consumible(grupo, consumible)
            resultados.append({
                "Nombre": nombre,
                "IP": ip,
                "Modelo": modelo,
                "Consumible": consumible,
                "Porcentaje actual": pct,
                "Consumo diario (%)": consumo,
                "Días restantes estimados": dias,
                "Fecha estimada de agotamiento": fecha_fin,
                "Método": metodo
            })

    df_pred = pd.DataFrame(resultados)

    # --- AGREGAR ALERTAS ---
    df_pred["Alerta"] = np.where(
        df_pred["Días restantes estimados"] <= 3, "⚠️ Reemplazar pronto", "OK")

    # --- GUARDAR RESULTADOS ---
    df_pred.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Predicciones guardadas en: {OUTPUT_FILE}")


