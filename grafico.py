import matplotlib.pyplot as plt
import pandas as pd

# Crear un DataFrame con los datos
data = {
    'Fecha y Hora': [
        '21-08-2024 10:16', '21-08-2024 11:43', '21-08-2024 13:30', '21-08-2024 15:15', '21-08-2024 15:30',
        '22-08-2024 08:00', '22-08-2024 13:33', '22-08-2024 15:45', '22-08-2024 16:52'
    ],
    'Posición 1': [131, 149, 25, 24, 137, 100, 171, 153, 148],
    'Posición 2': [71, 54, 179, 180, 150, 183, 112, 133, 139],
    'Posición 3': [53, 52, 51, 51, 478, 482, 482, 734, 733]
}

df = pd.DataFrame(data)
df['Fecha y Hora'] = pd.to_datetime(df['Fecha y Hora'])

# Crear el gráfico
plt.figure(figsize=(12, 6))
plt.plot(df['Fecha y Hora'], df['Posición 1'], label='Activo', marker='o')
plt.plot(df['Fecha y Hora'], df['Posición 2'], label='Inactivo', marker='o')
plt.plot(df['Fecha y Hora'], df['Posición 3'], label='Desconocido', marker='o')

plt.xlabel('Fecha y Hora')
plt.ylabel('Valores')
plt.title('Valores de Posiciones a lo largo del Tiempo')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()

# Mostrar el gráfico
plt.show()
