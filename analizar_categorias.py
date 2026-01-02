import pandas as pd

# Leer el archivo Excel consolidado
df = pd.read_excel('REQCONS_CONSOLIDADO.xlsx')

print("=" * 70)
print("ANÁLISIS DE CATEGORÍAS EN LA COLUMNA 'FAMILIA'")
print("=" * 70)

# Ver las primeras filas
print("\nPrimeras 5 filas del archivo:")
print(df.head())

# Ver todas las categorías únicas en la columna Familia
print("\n" + "=" * 70)
print("CATEGORÍAS ÚNICAS ENCONTRADAS:")
print("=" * 70)

categorias_unicas = df['Familia'].unique()
for i, categoria in enumerate(categorias_unicas, 1):
    print(f"{i}. {categoria}")

# Contar cuántos bienes hay por categoría
print("\n" + "=" * 70)
print("DISTRIBUCIÓN DE BIENES POR CATEGORÍA:")
print("=" * 70)

conteo_por_familia = df.groupby('Familia').agg({
    'Cantidad': 'sum',
    'Nombre': 'count'
}).rename(columns={'Nombre': 'Numero_de_Lotes'})

print(conteo_por_familia)

print("\n" + "=" * 70)
print(f"Total de categorías: {len(categorias_unicas)}")
print(f"Total de lotes: {len(df)}")
print("=" * 70)