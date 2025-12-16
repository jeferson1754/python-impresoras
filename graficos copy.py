import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

INPUT_FILE = r"G:\Unidades compartidas\Informática\Impresoras - final.xlsx"
TONER_COLUMNS = ["Toner Negro", "Toner Cian",
                 "Toner Magenta", "Toner Amarillo"]
OUTPUT_DIR = "graficos_toner/"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Cargar datos históricos ---
df = pd.read_excel(INPUT_FILE, sheet_name="Histórico")
df["Fecha de registro"] = pd.to_datetime(
    df["Marca de Tiempo"], errors="coerce")
df = df[df["Estado"] == "OK"]  # solo registros válidos
ALERTA_PCT = 10  # porcentaje crítico de tóner
OUTPUT_DIR = "graficos_toner/"
os.makedirs(OUTPUT_DIR, exist_ok=True)

for (ip, modelo), grupo in df.groupby(["IP", "Modelo"], dropna=False):
    # Obtener el nombre de la impresora si existe, si no usar modelo
    nombre_impresora = grupo["Nombre"].iloc[0] if "Nombre" in grupo.columns else modelo

    plt.figure(figsize=(10, 6))
    for color in TONER_COLUMNS:
        if color not in grupo.columns:
            continue
        serie = grupo.sort_values("Fecha de registro")[
            ["Fecha de registro", color]].dropna()
        if serie.empty:
            continue
        # Convertir porcentaje a numérico
        serie[color] = pd.to_numeric(serie[color].astype(
            str).str.replace("%", ""), errors="coerce")
        plt.plot(serie["Fecha de registro"],
                 serie[color], marker='o', label=color)

    # Línea de alerta
    plt.axhline(ALERTA_PCT, color='red', linestyle='--',
                linewidth=1.5, label=f'Alerta {ALERTA_PCT}%')

    plt.title(f"Consumo de Tóner - {nombre_impresora} ({ip})")
    plt.xlabel("Fecha")
    plt.ylabel("Porcentaje restante (%)")
    plt.ylim(0, 110)
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()

    # Guardar gráfico
    nombre_archivo = f"{OUTPUT_DIR}{ip}_{nombre_impresora.replace(' ', '_')}.png"
    plt.savefig(nombre_archivo)
    plt.close()
    print(f"✅ Gráfico guardado: {nombre_archivo}")
    # --- Ejecutar app ---
    

