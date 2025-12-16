import os
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

from datetime import timedelta
from sklearn.linear_model import LinearRegression

# --------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------
INPUT_FILE = r"G:\Unidades compartidas\Informática\Impresoras - final.xlsx"
OUTPUT_FILE = "predicciones_toner_ema.xlsx"

TONER_COLUMNS = ["Toner Negro", "Toner Cian",
                 "Toner Magenta", "Toner Amarillo"]
KITS_COLUMNS = ["Kit Mant.", "Kit Alim."]
CONSUMIBLES = TONER_COLUMNS + KITS_COLUMNS

ESTADO_VALIDO = "OK"
VENTANA_EMA = 10
DIAS_ALERTA_CRITICA = 3
DIAS_ALERTA_MEDIA = 7
MAX_DIAS_PREDICCION = 365 * 2  # máximo 2 años


# --------------------------------------------------
# CARGA Y LIMPIEZA DE DATOS
# --------------------------------------------------
df = pd.read_excel(INPUT_FILE, sheet_name="Histórico")
df.columns = df.columns.str.strip()

df["Fecha de registro"] = pd.to_datetime(
    df["Marca de Tiempo"], errors="coerce")
df = df[df["Estado"].str.strip() == ESTADO_VALIDO].copy()

for col in CONSUMIBLES:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace("", np.nan)
        .astype(float)
    )

df.sort_values("Fecha de registro", ascending=False, inplace=True)
df.drop_duplicates(subset=["IP", "Marca de Tiempo"],
                   keep="first", inplace=True)

# --------------------------------------------------
# FUNCIÓN DE PREDICCIÓN (EMA + FALLBACK)
# --------------------------------------------------




# --------------------------------------------------
# GENERAR PREDICCIONES
# --------------------------------------------------
