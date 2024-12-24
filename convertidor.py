import pandas as pd

# Cargar el archivo JSON en un DataFrame de pandas
df = pd.read_json("nbot.json")

# Guardar el DataFrame como un archivo CSV
df.to_csv("nbot.csv", index=False)
