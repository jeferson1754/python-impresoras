import pandas as pd
import dash
from dash import dcc, html
import plotly.express as px
from dash.dependencies import Input, Output

# --- Configuración ---
# Archivo con predicciones (IMPORTANTE: La ruta local debe ser correcta)
INPUT_FILE = r"C:\Users\jvargas\Phyton\python-impresoras\predicciones_toner.xlsx"

# --- Cargar datos con manejo de errores ---
try:
    df = pd.read_excel(INPUT_FILE)
except FileNotFoundError:
    print(
        f"ERROR: No se encontró el archivo de datos en la ruta: {INPUT_FILE}")
    print("Por favor, verifica la ruta de INPUT_FILE.")
    # Crea un DataFrame vacío para que la app no falle al arrancar
    df = pd.DataFrame(columns=["Fecha estimada de agotamiento", "Consumible",
                      "Modelo", "Porcentaje actual", "IP", "Días restantes estimados", "Método"])

if df.empty:
    # Si el archivo está vacío o el loading falló, se usan listas vacías.
    CONSUMIBLES = []
    MODELOS = []
else:
    df["Fecha estimada de agotamiento"] = pd.to_datetime(
        df["Fecha estimada de agotamiento"], errors="coerce"
    )
    # Rellenar cualquier valor nulo de la fecha para evitar problemas en Plotly
    df.dropna(subset=['Fecha estimada de agotamiento'], inplace=True)
    CONSUMIBLES = df["Consumible"].unique()
    MODELOS = df["Modelo"].unique()


# --- Crear app Dash ---
app = dash.Dash(__name__)

app.layout = html.Div([
    # Título centrado
    html.H1("Dashboard de Consumo de Tóner", style={'textAlign': 'center'}),

    # Contenedores de Dropdown (filtros)
    html.Div([
        html.Div([
            html.Label("Selecciona impresora:"),
            dcc.Dropdown(
                id='dropdown-modelo',
                options=[{"label": m, "value": m} for m in MODELOS],
                multi=True,
                placeholder="Todos los modelos"
            ),
        ], style={"width": "48%", "display": "inline-block"}),

        html.Div([
            html.Label("Selecciona consumible:"),
            dcc.Dropdown(
                id='dropdown-consumible',
                options=[{"label": c, "value": c} for c in CONSUMIBLES],
                multi=True,
                placeholder="Todos los consumibles"
            ),
        ], style={"width": "48%", "display": "inline-block", "float": "right"}),
    ], style={'padding': '10px 50px'}),  # Separación para los filtros

    dcc.Graph(id='graph-toner')
])

# --- Callback para actualizar gráfico ---


@app.callback(
    Output('graph-toner', 'figure'),
    Input('dropdown-modelo', 'value'),
    Input('dropdown-consumible', 'value')
)
def update_graph(selected_modelos, selected_consumibles):
    filtered_df = df.copy()

    # 1. Aplicar filtros
    if selected_modelos:
        filtered_df = filtered_df[filtered_df["Modelo"].isin(selected_modelos)]
    if selected_consumibles:
        filtered_df = filtered_df[filtered_df["Consumible"].isin(
            selected_consumibles)]

    # 2. Manejar caso sin datos
    if filtered_df.empty:
        return px.line(title="No hay datos para mostrar")

    # 3. Crear el gráfico con mejoras visuales
    fig = px.line(
        filtered_df,
        x="Fecha estimada de agotamiento",
        y="Porcentaje actual",
        color="Consumible",
        line_group="IP",
        markers=True,  # Añadir puntos para mayor claridad
        hover_data=["IP", "Nombre", "Modelo", "Días restantes estimados"]
    )

    # 4. Ajustar el layout y añadir línea de umbral
    fig.update_layout(
        title="Proyección de Consumo de Tóner",
        xaxis_title="Fecha de Proyección de Agotamiento",
        yaxis_title="Porcentaje restante (%)",
        yaxis=dict(range=[0, 100]),
        # === LÍNEA MODIFICADA AL 10% ===
        shapes=[
            dict(
                type='line',
                xref='paper', yref='y',
                x0=0, x1=1, y0=10, y1=10,  # ¡CAMBIO AQUÍ: y0=10 y y1=10!
                line=dict(color='red', width=2, dash='dash')
            )
        ]
    )
    return fig


# --- Ejecutar app ---
if __name__ == "__main__":
    app.run(debug=True)
