from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import PolynomialFeatures
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os
from datetime import timedelta
from sklearn.linear_model import LinearRegression

# --- CONFIGURACIÓN ---
INPUT_FILE = r"G:\Unidades compartidas\Informática\Impresoras - final.xlsx"
# Nuevo nombre para diferenciar la versión EMA
OUTPUT_FILE = "predicciones_toner_ema.xlsx"

TONER_COLUMNS = ["Toner Negro", "Toner Cian",
                 "Toner Magenta", "Toner Amarillo"]
KITS_COLUMNS = ["Kit Mant.", "Kit Alim."]
CONSUMIBLES = TONER_COLUMNS + KITS_COLUMNS

ESTADO_VALIDO = "OK"
VENTANA_EMA = 10  # Parámetro para suavizado EMA (span)
DIAS_ALERTA_CRITICA = 3
DIAS_ALERTA_MEDIA = 7

# --- CARGA Y PRE-PROCESAMIENTO ---
try:
    df = pd.read_excel(INPUT_FILE, sheet_name="Histórico")
except Exception as e:
    raise Exception(f"Error al leer el archivo histórico: {e}")

df.columns = [col.strip() for col in df.columns]
df["Fecha de registro"] = pd.to_datetime(
    df["Marca de Tiempo"], errors="coerce")
df = df[df["Estado"].str.strip() == ESTADO_VALIDO].copy()

# Convertir columnas de consumibles a numérico
for col in CONSUMIBLES:
    df[col] = pd.to_numeric(
        df[col].astype(str).str.replace("%", "").str.strip(), errors="coerce"
    )

# Eliminar duplicados de marca de tiempo, manteniendo el más reciente
df.sort_values("Fecha de registro", ascending=False, inplace=True)
df.drop_duplicates(subset=["IP", "Marca de Tiempo"],
                   keep="first", inplace=True)


# --- FUNCIÓN DE PREDICCIÓN (USANDO EMA) ---
def predecir_consumible(sub_df, consumible):
    sub_df = sub_df.sort_values("Fecha de registro")
    sub_df = sub_df[["Fecha de registro", consumible]].dropna()
    sub_df = sub_df.dropna()

    if len(sub_df) < 2:
        return np.nan, np.nan, np.nan, np.nan, "❌ Muy pocos datos"

    # Crear eje temporal en días
    sub_df["Días"] = (sub_df["Fecha de registro"] -
                      sub_df["Fecha de registro"].min()).dt.total_seconds() / (24*3600)

    y = sub_df[consumible].values
    porcentaje_actual = y[-1]

    # 1. Calcular la Tasa de Consumo Instantánea (Delta %)
    sub_df['Delta_Pct'] = sub_df[consumible].diff().abs()
    sub_df['Delta_Dias'] = sub_df["Días"].diff()
    # Tasa instantánea de consumo: % / día
    sub_df['Tasa_Consumo_Inst'] = sub_df['Delta_Pct'] / sub_df['Delta_Dias']

    # 2. Aplicar el Promedio Móvil Exponencial (EMA) para suavizado
    # Esto da más peso a los datos recientes
    consumo_diario_ema = (
        sub_df['Tasa_Consumo_Inst']
        .ewm(span=VENTANA_EMA, adjust=False)
        .mean()
        .iloc[-1]
    )
    metodo = f"⭐ EMA (span={VENTANA_EMA})"

    consumo_diario = consumo_diario_ema

    # Fallback a Regresión Lineal si EMA es inválida (por ejemplo, en el primer punto)
    if np.isnan(consumo_diario) or consumo_diario <= 0:
        if len(sub_df) >= 3:
            X = sub_df[["Días"]].values
            model = LinearRegression()
            model.fit(X, y)
            consumo_diario = -model.coef_[0]
            metodo = "📈 Regresión Lineal (Fallback)"
        else:
            # Fallback a promedio simple de los dos últimos puntos
            delta_pct = y[-2] - y[-1]
            delta_days = sub_df["Días"].iloc[-1] - sub_df["Días"].iloc[-2]
            consumo_diario = delta_pct / delta_days if delta_days > 0 else np.nan
            metodo = "⚙️ Promedio simple"

    if consumo_diario <= 0 or np.isnan(consumo_diario):
        return porcentaje_actual, 0, np.nan, np.nan, f"{metodo} - Pendiente Inválida"

    dias_restantes = porcentaje_actual / consumo_diario
    fecha_agotamiento = sub_df["Fecha de registro"].iloc[-1] + \
        timedelta(days=dias_restantes)

    return round(porcentaje_actual, 1), round(consumo_diario, 4), round(dias_restantes, 1), fecha_agotamiento, metodo


# --- GENERAR PREDICCIONES ---
resultados = []

for (ip, modelo), grupo in df.groupby(["IP", "Modelo"], dropna=False, observed=True):
    nombre = grupo["Nombre"].iloc[0] if "Nombre" in grupo.columns and not grupo["Nombre"].empty else ip

    for consumible in CONSUMIBLES:
        if consumible not in grupo.columns or grupo[consumible].dropna().empty:
            continue

        # Uso de la función de predicción basada en EMA/Regresión
        pct, consumo, dias, fecha_fin, metodo = predecir_consumible(
            grupo, consumible)

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

# --- AGREGAR ALERTAS MEJORADAS ---


def generar_alerta(dias):
    if pd.isna(dias):
        return "❓ Datos insuficientes"
    elif dias <= DIAS_ALERTA_CRITICA:
        return "🚨 REEMPLAZAR URGENTE"
    elif dias <= DIAS_ALERTA_MEDIA:
        return "⚠️ Reemplazar pronto"
    elif dias <= 15:
        return "🔔 Bajo stock (2 semanas)"
    else:
        return "🟢 OK"


df_pred["Alerta"] = df_pred["Días restantes estimados"].apply(generar_alerta)


# --- 7. FUNCIÓN DE GRÁFICOS (Ahora usando Regresión Lineal para visualización) ---

def generar_graficos(df_historico, df_predicciones, consumibles_list):
    GRAFICOS_FOLDER = "graficos_prediccion_ema"  # Carpeta diferente
    os.makedirs(GRAFICOS_FOLDER, exist_ok=True)

    print("\nGenerando gráficos de predicción individual...")

    df_historico_clean = df_historico.copy()
    df_historico_clean = df_historico_clean.dropna(
        subset=["Fecha de registro"])
    for col in consumibles_list:
        df_historico_clean[col] = pd.to_numeric(
            df_historico_clean[col].astype(str).str.replace("%", "").str.strip(), errors="coerce")

    # Solo graficar los que tienen alerta URGENTE o PRONTO
    top_alertas = df_predicciones[
        (df_predicciones['Alerta'] == "🚨 REEMPLAZAR URGENTE") |
        (df_predicciones['Alerta'] == "⚠️ Reemplazar pronto") |
        (df_predicciones['Alerta'] == "🔔 Bajo stock (2 semanas)")
    ].sort_values("Días restantes estimados", na_position='last')

    if top_alertas.empty:
        print("No hay alertas críticas para graficar individualmente.")

    for index, row in top_alertas.iterrows():
        ip = row['IP']
        consumible = row['Consumible']
        alerta = row['Alerta']

        grupo = df_historico_clean[(df_historico_clean["IP"] == ip)].copy()
        grupo = grupo.dropna(subset=[consumible])

        if len(grupo) < 2:
            continue

        # Usar Regresión Lineal para una visualización estable (Evita la curva Polinomial que subía)
        grupo["Días"] = (grupo["Fecha de registro"] -
                         grupo["Fecha de registro"].min()).dt.total_seconds() / (24*3600)

        X_hist = grupo[["Días"]].values
        y_hist = grupo[consumible].values

        try:
            # Usar LinearRegression (Grado 1) para la visualización
            modelo_lineal = LinearRegression()
            modelo_lineal.fit(X_hist, y_hist)

            dias_fin_pred = row['Días restantes estimados']

            # Ajustar rango de predicción
            if pd.isna(dias_fin_pred) or dias_fin_pred < 0:
                dias_futuro = grupo["Días"].max() + 30
            else:
                # 5 días extra de margen
                dias_futuro = grupo["Días"].max() + dias_fin_pred + 5

            X_pred = np.arange(grupo["Días"].min(),
                               dias_futuro).reshape(-1, 1)
            y_pred = modelo_lineal.predict(X_pred)

            fecha_inicio = grupo["Fecha de registro"].min()
            fechas_pred = fecha_inicio + \
                pd.to_timedelta(X_pred.flatten(), unit='D')

        except Exception:
            fechas_pred, y_pred = [], []
            pass

        # Configuración del gráfico
        plt.figure(figsize=(10, 6))

        plt.scatter(grupo["Fecha de registro"], y_hist,
                    color='darkblue', s=50, label='Histórico de %')

        if len(fechas_pred) > 0:
            plt.plot(fechas_pred, y_pred, color='red', linestyle='--',
                     linewidth=2, label='Tendencia Lineal de Predicción')

            if not pd.isna(row['Fecha estimada de agotamiento']):
                plt.axvline(x=row['Fecha estimada de agotamiento'], color='darkorange',
                            linestyle=':', linewidth=2, label='Fecha Agotamiento Est.')
                plt.text(row['Fecha estimada de agotamiento'], 5, f"{row['Fecha estimada de agotamiento'].strftime('%Y-%m-%d')}",
                         rotation=90, verticalalignment='bottom')

        plt.title(
            f"Consumo Histórico y Predicción: {row['Nombre']} ({ip})\nConsumible: {consumible} | Alerta: {alerta}", fontsize=14)
        plt.xlabel("Fecha de Registro")
        plt.ylabel(f"Porcentaje de Consumible (%)")
        plt.ylim(0, 105)
        plt.grid(axis='y', linestyle='--')
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        nombre_archivo = f"{GRAFICOS_FOLDER}/{ip}_{consumible.replace(' ', '_')}.png"
        plt.savefig(nombre_archivo)
        plt.close()

    print(
        f"✅ Gráficos de tendencia guardados en la carpeta: {GRAFICOS_FOLDER}")

    # -----------------------------------------------------
    # GRÁFICO 2: RESUMEN DE ALERTAS (General)
    # -----------------------------------------------------

    conteo_alertas = df_predicciones.groupby(
        'Alerta').size().reset_index(name='Cantidad')

    orden_alertas = ["🚨 REEMPLAZAR URGENTE", "⚠️ Reemplazar pronto",
                     "🔔 Bajo stock (2 semanas)", "🟢 OK", "❓ Datos insuficientes"]
    conteo_alertas['Alerta'] = pd.Categorical(
        conteo_alertas['Alerta'], categories=orden_alertas, ordered=True)
    conteo_alertas = conteo_alertas.sort_values('Alerta')

    plt.figure(figsize=(12, 7))
    sns.barplot(
        x='Alerta',
        y='Cantidad',
        data=conteo_alertas,
        palette=['red', 'orange', 'gold', 'green', 'gray']
    )

    for index, row in conteo_alertas.iterrows():
        plt.text(index, row['Cantidad'] + 0.1, str(row['Cantidad']),
                 ha='center', va='bottom', fontsize=12)

    plt.title("Resumen de Predicciones por Nivel de Alerta", fontsize=16)
    plt.xlabel("Nivel de Alerta")
    plt.ylabel("Cantidad de Consumibles/Impresoras")
    plt.xticks(rotation=15)
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()

    plt.savefig(f"{GRAFICOS_FOLDER}/resumen_alertas_global.png")
    plt.close()

    print("✅ Gráfico de resumen de alertas guardado.")


# --- LLAMADA A LA FUNCIÓN DE GRÁFICOS ---
generar_graficos(df, df_pred, CONSUMIBLES)

# --- GUARDAR RESULTADOS ---
df_pred.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Predicciones guardadas en: {OUTPUT_FILE}")
