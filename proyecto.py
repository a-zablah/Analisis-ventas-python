import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_csv("sample.csv", encoding="latin1")
print(df.head())

#Cantidad de filas y columnas
print("Forma del dataset:", df.shape) 

#Tipos de datos
print("\nTipos de datos:") 
print(df.info())

#Valores nulos por columna
print("\nValores nulos por columna:") 
print(df.isnull().sum())

#Resumen estadístico por columnas
print("\nResumen estadístico:") 
print(df.describe())

#Producto más vendido
producto_mas_vendido = df.groupby("Product Name")["Quantity"].sum().idxmax()

#Ciudad con más ventas
ciudad_mas_ventas = df.groupby("City")["Quantity"].sum().idxmax()

#Convertir  Order Date a tipo de fecha
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")

#Crear columna con el mes
df["Mes"] = df["Order Date"].dt.month

#Mes con más ingresos
mes_mas_ingresos = df.groupby("Mes")["Sales"].sum().idxmax()

print(f"\nProducto más vendido: {producto_mas_vendido}")
print(f"Ciudad con más ventas: {ciudad_mas_ventas}")
print(f"Mes con más ingresos: {mes_mas_ingresos}")

resultados = {
    "Métrica": ["Producto más vendido", "Ciudad con más ventas", "Mes con más ingresos"],
    "Valor": ["Staples", "New York City", "Noviembre"]
}

df_resultados = pd.DataFrame(resultados)

df_resultados.to_excel("analisis_ventas.xlsx", index=False)

print("Archivo Excel exportado correctamente.")