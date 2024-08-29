import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.widgets import CheckButtons

# Función para leer los datos desde un archivo de texto
def leer_datos_desde_txt(archivo):
    with open(archivo, 'r') as file:
        lines = file.readlines()

    data = {'Fecha y Hora': [], 'Posición 1': [], 'Posición 2': [], 'Posición 3': []}
    current_date = ''

    for line in lines:
        line = line.strip()
        if '-' in line and len(line.split()) == 1:  # Detectar la fecha
            current_date = line
        elif line:  # Detectar líneas con hora y valores
            time, values = line.split(' ', 1)
            position_values = [int(x.strip()) for x in values.split(',')]
            data['Fecha y Hora'].append(f'{current_date} {time}')
            data['Posición 1'].append(position_values[0])
            data['Posición 2'].append(position_values[1])
            data['Posición 3'].append(position_values[2])

    return pd.DataFrame(data)

# Leer los datos desde el archivo de texto
archivo = r"C:\Users\jvargas\Documents\data.txt"  # Cambia esto por la ruta real de tu archivo de texto
df = leer_datos_desde_txt(archivo)

# Convertir la columna de fecha y hora a tipo datetime
df['Fecha y Hora'] = pd.to_datetime(df['Fecha y Hora'])

# Crear el gráfico
fig, ax = plt.subplots(figsize=(12, 6))
line1, = ax.plot(df['Fecha y Hora'], df['Posición 1'], label='Activos', marker='o')
line2, = ax.plot(df['Fecha y Hora'], df['Posición 2'], label='Inactivos', marker='o')
line3, = ax.plot(df['Fecha y Hora'], df['Posición 3'], label='Desconocidos', marker='o')

plt.xlabel('Fecha y Hora')
plt.ylabel('Valores')
plt.title('Equipos en la red')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()

# Crear la anotación que aparecerá al pasar el mouse
annot = ax.annotate("", xy=(0,0), xytext=(20,20),
                    textcoords="offset points",
                    bbox=dict(boxstyle="round", fc="w"),
                    arrowprops=dict(arrowstyle="->"))
annot.set_visible(False)

# Función para actualizar la anotación
def update_annot(line, ind):
    x, y = line.get_data()
    annot.xy = (x[ind["ind"][0]], y[ind["ind"][0]])
    text = f"{line.get_label()}\n{df['Fecha y Hora'].iloc[ind['ind'][0]]}\n{y[ind['ind'][0]]}"
    annot.set_text(text)
    annot.get_bbox_patch().set_alpha(0.8)

# Función para manejar los eventos del mouse
def hover(event):
    vis = annot.get_visible()
    if event.inaxes == ax:
        for line in [line1, line2, line3]:
            cont, ind = line.contains(event)
            if cont:
                update_annot(line, ind)
                annot.set_visible(True)
                fig.canvas.draw_idle()
                return
    if vis:
        annot.set_visible(False)
        fig.canvas.draw_idle()

fig.canvas.mpl_connect("motion_notify_event", hover)

# Configurar los botones de selección
rax = plt.axes([0.4, 0.80, 0.2, 0.1])  # Posición en la parte superior del gráfico
labels = ['Activos', 'Inactivos', 'Desconocidos']
visibility = [line1.get_visible(), line2.get_visible(), line3.get_visible()]
check = CheckButtons(rax, labels, visibility)

def func(label):
    if label == 'Activos':
        line1.set_visible(not line1.get_visible())
    elif label == 'Inactivos':
        line2.set_visible(not line2.get_visible())
    elif label == 'Desconocidos':
        line3.set_visible(not line3.get_visible())
    plt.draw()

check.on_clicked(func)

# Mostrar el gráfico
plt.show()
