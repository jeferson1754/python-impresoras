"""
Sistema de Monitoreo y Predicción de Consumo de Tóner
- Funcionalidades:
    1. Dashboard interactivo (Dash)
    2. Gráficos de consumo por impresora con alerta
    3. Comparativa de consumo entre impresoras
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Librerías Dash
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px

# ---------------- CONFIGURACIÓN ---------------- #
INPUT_FILE = r"C:\Users\jvargas\Phyton\python-impresoras\predicciones_toner.xlsx"
OUTPUT_DIR_GRAFICOS = "graficos_toner_prediccion"
ALERTA_PCT = 10  # porcentaje crítico de tóner

os.makedirs(OUTPUT_DIR_GRAFICOS, exist_ok=True)

# Cargar datos
df = pd.read_excel(INPUT_FILE)
df["Fecha estimada de agotamiento"] = pd.to_datetime(
    df["Fecha estimada de agotamiento"], errors="coerce")

TONER_COLUMNS = ["Toner Negro", "Toner Cian",
                 "Toner Magenta", "Toner Amarillo"]

# --------------------------------------------- #
# Función 1: Generar gráficos por impresora con alerta

def graficos_consumo_por_impresora():
   # Agrupar por IP y Modelo para generar una gráfica por impresora
    for (ip, modelo), grupo in df.groupby(["IP", "Modelo"], dropna=False):
        nombre_impresora = grupo["Nombre"].iloc[0] if "Nombre" in grupo.columns else modelo

        plt.figure(figsize=(10, 6))
        datos_encontrados = False

        # === SOLUCIÓN: Agrupar por Consumible para trazar las líneas ===
        # Iterar sobre cada tipo de tóner dentro de esa impresora
        for consumible, grupo_consumible in grupo.groupby("Consumible", dropna=False):

            # Las columnas a seleccionar ahora son la Fecha y el Porcentaje actual
            required_cols = [
                "Fecha estimada de agotamiento", "Porcentaje actual"]

            if "Porcentaje actual" not in grupo_consumible.columns:
                continue

            # Seleccionar y limpiar datos para este consumible específico
            serie = grupo_consumible.sort_values("Fecha estimada de agotamiento")[
                required_cols].dropna(subset=["Porcentaje actual"])

            if serie.empty:
                continue

            # Convertir la columna de porcentaje a numérico
            serie.loc[:, "Porcentaje actual"] = pd.to_numeric(
                serie["Porcentaje actual"].astype(
                    str).str.replace("%", "").str.strip(),
                errors="coerce"
            )

            # Quitar filas donde la conversión a numérico falló
            serie.dropna(subset=["Porcentaje actual"], inplace=True)

            if serie.empty:
                continue

            # Graficar usando el nombre del consumible como etiqueta
            plt.plot(serie["Fecha estimada de agotamiento"],
                     serie["Porcentaje actual"],
                     marker='o',
                     # La etiqueta es el valor (ej. 'Toner Amarillo')
                     label=consumible)
            datos_encontrados = True

        if not datos_encontrados:
            plt.title(
                f"Sin datos de consumo válidos para {nombre_impresora} ({ip})")
            plt.text(0.5, 0.5, 'No hay datos válidos para graficar.',
                     horizontalalignment='center', verticalalignment='center',
                     transform=plt.gca().transAxes)

        else:
            # Diseño y Visualización
            plt.axhline(ALERTA_PCT, color='red', linestyle='--',
                        linewidth=1.5, label=f'Alerta {ALERTA_PCT}%')
            plt.title(f"Consumo de Tóner - {nombre_impresora} ({ip})")
            plt.xlabel("Fecha")
            plt.ylabel("Porcentaje restante (%)")
            plt.ylim(0, 100)
            plt.grid(True, linestyle="--", alpha=0.5)
            plt.legend(loc='lower left')

            # Generar el nombre de archivo limpio
            nombre_archivo = f"{OUTPUT_DIR_GRAFICOS}/{ip}_{nombre_impresora.replace(' ', '_').replace('/', '_')}.png"
            plt.savefig(nombre_archivo)
            print(f"✅ Gráfico guardado: {nombre_archivo}")

        # Acción de limpieza
        plt.tight_layout()
        plt.close()

    print("Proceso de generación de gráficos por impresora completado.")


# --------------------------------------------- #
# Función 3: Comparativa entre impresoras
def comparativa_consumo_impresoras():
    df_toner = df[df["Consumible"].isin(TONER_COLUMNS)]
    df_toner["Porcentaje actual"] = pd.to_numeric(
        df_toner["Porcentaje actual"], errors="coerce")
    df_toner["Consumo diario (%)"] = pd.to_numeric(
        df_toner["Consumo diario (%)"], errors="coerce")

    consumo_promedio = (
        df_toner.groupby(["Nombre", "Modelo", "Consumible"])
        .agg({"Consumo diario (%)": "mean", "Días restantes estimados": "mean"})
        .reset_index()
    )

    consumo_promedio.sort_values(
        by="Consumo diario (%)", ascending=False, inplace=True)

    # Mostrar tabla
    print("\n=== Consumo promedio diario por impresora y consumible ===")
    print(consumo_promedio.head(20))

    # Gráfico comparativo
    plt.figure(figsize=(12, 6))
    sns.barplot(
        data=consumo_promedio,
        x="Nombre",
        y="Consumo diario (%)",
        hue="Consumible"
    )
    plt.title("Comparativa de consumo diario de tóner por impresora")
    plt.xlabel("Impresora")
    plt.ylabel("Consumo diario (%)")
    plt.xticks(rotation=45)
    plt.legend(title="Consumible")
    plt.tight_layout()
    plt.show()


# --------------------------------------------- #
# Menú principal
def menu():
    while True:
        print("\n=== Sistema de Monitoreo de Tóner ===")
        print("1. Generar gráficos por impresora con alerta")
        print("2. Comparativa de consumo entre impresoras")
        print("0. Salir")
        opcion = input("Selecciona una opción: ")

        if opcion == "1":
            graficos_consumo_por_impresora()
        elif opcion == "2":
            comparativa_consumo_impresoras()
        elif opcion == "0":
            print("Saliendo...")
            break
        else:
            print("Opción no válida, intenta de nuevo.")


# --------------------------------------------- #
if __name__ == "__main__":
    menu()
