import pdfplumber
import pandas as pd
from pathlib import Path
import re

def extraer_tablas_pdf(ruta_pdf):
    """
    Extrae todas las tablas del PDF y las convierte en un DataFrame
    
    Args:
        ruta_pdf: Ruta del archivo PDF a procesar
    
    Returns:
        DataFrame con todos los datos extraídos
    """
    print(f"📄 Abriendo PDF: {ruta_pdf}")
    print("⏳ Extrayendo tablas... (esto puede tardar unos minutos)")
    
    todas_las_filas = []
    
    with pdfplumber.open(ruta_pdf) as pdf:
        total_paginas = len(pdf.pages)
        print(f"📊 Total de páginas: {total_paginas}")
        
        for num_pagina, pagina in enumerate(pdf.pages, start=1):
            # Mostrar progreso cada 10 páginas
            if num_pagina % 10 == 0 or num_pagina == 1:
                print(f"   Procesando página {num_pagina}/{total_paginas}...")
            
            # Extraer tablas de la página
            tablas = pagina.extract_tables()
            
            for tabla in tablas:
                if tabla and len(tabla) > 1:  # Verificar que la tabla tenga contenido
                    # La primera fila suele ser encabezados
                    encabezados = tabla[0]
                    
                    # Procesar el resto de filas
                    for fila in tabla[1:]:
                        if fila and any(fila):  # Verificar que la fila no esté vacía
                            # Verificar que tenga al menos los campos básicos
                            if len(fila) >= 4:  # Cantidad, Nombre, Descripción, Familia mínimo
                                todas_las_filas.append(fila)
    
    print(f"✅ Extracción completada: {len(todas_las_filas)} filas encontradas")
    
    # Crear DataFrame con nombres de columnas estándar
    columnas = ['Cantidad', 'Nombre', 'Descripcion', 'Familia', 'Unidad', 'Observaciones']
    
    # Ajustar según el número de columnas que realmente tenga
    if todas_las_filas:
        num_columnas = len(todas_las_filas[0])
        if num_columnas < len(columnas):
            columnas = columnas[:num_columnas]
        elif num_columnas > len(columnas):
            # Agregar columnas extra si hay más
            for i in range(len(columnas), num_columnas):
                columnas.append(f'Columna_Extra_{i+1}')
    
    df = pd.DataFrame(todas_las_filas, columns=columnas)
    
    return df

def limpiar_datos(df):
    """
    Limpia y prepara los datos para su procesamiento
    
    Args:
        df: DataFrame con datos crudos
    
    Returns:
        DataFrame limpio
    """
    print("\n🧹 Limpiando datos...")
    
    # Crear una copia para no modificar el original
    df_limpio = df.copy()
    
    # Eliminar filas completamente vacías
    df_limpio = df_limpio.dropna(how='all')
    
    # Eliminar filas donde Nombre y Descripción estén vacíos
    df_limpio = df_limpio[
        (df_limpio['Nombre'].notna()) & 
        (df_limpio['Nombre'].astype(str).str.strip() != '')
    ]
    
    # Limpiar espacios en blanco
    columnas_texto = ['Nombre', 'Descripcion', 'Familia', 'Unidad', 'Observaciones']
    for col in columnas_texto:
        if col in df_limpio.columns:
            df_limpio[col] = df_limpio[col].astype(str).str.strip()
    
    # Convertir Cantidad a numérico
    df_limpio['Cantidad'] = pd.to_numeric(df_limpio['Cantidad'], errors='coerce')
    
    # Eliminar filas donde la cantidad no sea válida
    df_limpio = df_limpio[df_limpio['Cantidad'].notna()]
    
    print(f"✅ Datos limpios: {len(df_limpio)} filas válidas")
    
    return df_limpio

def consolidar_lotes(df):
    """
    Agrupa y consolida lotes repetidos basándose en Nombre + Descripción
    
    Args:
        df: DataFrame con datos limpios
    
    Returns:
        DataFrame consolidado con cantidades sumadas
    """
    print("\n🔄 Consolidando lotes repetidos...")
    
    # Agrupar por Nombre + Descripción y sumar cantidades
    df_consolidado = df.groupby(['Nombre', 'Descripcion'], as_index=False).agg({
        'Cantidad': 'sum',
        'Familia': 'first',  # Tomar el primer valor
        'Unidad': 'first',
        'Observaciones': lambda x: ' | '.join(x.dropna().astype(str).unique())  # Concatenar observaciones únicas
    })
    
    # Contar cuántas veces se repitió cada lote
    conteo_repeticiones = df.groupby(['Nombre', 'Descripcion']).size().reset_index(name='Veces_Repetido')
    
    # Agregar columna de repeticiones
    df_consolidado = df_consolidado.merge(conteo_repeticiones, on=['Nombre', 'Descripcion'])
    
    # Reordenar columnas
    columnas_ordenadas = ['Cantidad', 'Nombre', 'Descripcion', 'Familia', 'Unidad', 'Veces_Repetido', 'Observaciones']
    df_consolidado = df_consolidado[columnas_ordenadas]
    
    # Ordenar por cantidad descendente
    df_consolidado = df_consolidado.sort_values('Cantidad', ascending=False)
    
    print(f"✅ Consolidación completada:")
    print(f"   - Filas originales: {len(df)}")
    print(f"   - Filas consolidadas: {len(df_consolidado)}")
    print(f"   - Lotes únicos: {len(df_consolidado)}")
    
    return df_consolidado

def generar_reportes(df_original, df_consolidado, nombre_base):
    """
    Genera archivos Excel con los resultados
    
    Args:
        df_original: DataFrame con datos originales
        df_consolidado: DataFrame consolidado
        nombre_base: Nombre base para los archivos de salida
    """
    print("\n💾 Generando archivos Excel...")
    
    # Archivo consolidado
    archivo_consolidado = f"{nombre_base}_CONSOLIDADO.xlsx"
    df_consolidado.to_excel(archivo_consolidado, index=False, sheet_name='Lotes Consolidados')
    print(f"✅ Archivo consolidado creado: {archivo_consolidado}")
    
    # Archivo con datos originales (para verificación)
    archivo_original = f"{nombre_base}_DATOS_ORIGINALES.xlsx"
    df_original.to_excel(archivo_original, index=False, sheet_name='Datos Extraídos')
    print(f"✅ Archivo original creado: {archivo_original}")
    
    # Generar reporte de resumen
    print("\n📊 RESUMEN:")
    print(f"   - Total de lotes únicos: {len(df_consolidado)}")
    print(f"   - Lotes que se repitieron: {len(df_consolidado[df_consolidado['Veces_Repetido'] > 1])}")
    print(f"   - Cantidad total de artículos: {df_consolidado['Cantidad'].sum():.0f}")
    
    # Mostrar los 5 lotes más repetidos
    mas_repetidos = df_consolidado.nlargest(5, 'Veces_Repetido')[['Nombre', 'Cantidad', 'Veces_Repetido']]
    print("\n🔝 Top 5 lotes más repetidos:")
    print(mas_repetidos.to_string(index=False))

def main():
    """
    Función principal que ejecuta todo el proceso
    """
    print("=" * 70)
    print("🚀 CONSOLIDADOR DE LICITACIÓN - INICIO")
    print("=" * 70)
    
    # Solicitar ruta del PDF
    print("\n📁 Ingresa la ruta completa del archivo PDF:")
    print("   Ejemplo: C:\\Users\\TuUsuario\\Documents\\licitacion.pdf")
    ruta_pdf = input("   Ruta: ").strip().strip('"')  # Eliminar comillas si las hay
    
    # Verificar que el archivo existe
    if not Path(ruta_pdf).exists():
        print(f"❌ ERROR: El archivo no existe: {ruta_pdf}")
        return
    
    try:
        # Paso 1: Extraer tablas
        df_crudo = extraer_tablas_pdf(ruta_pdf)
        
        if len(df_crudo) == 0:
            print("❌ ERROR: No se encontraron tablas en el PDF")
            return
        
        # Paso 2: Limpiar datos
        df_limpio = limpiar_datos(df_crudo)
        
        if len(df_limpio) == 0:
            print("❌ ERROR: No quedaron datos válidos después de la limpieza")
            return
        
        # Paso 3: Consolidar lotes
        df_consolidado = consolidar_lotes(df_limpio)
        
        # Paso 4: Generar reportes
        nombre_base = Path(ruta_pdf).stem  # Nombre del archivo sin extensión
        generar_reportes(df_limpio, df_consolidado, nombre_base)
        
        print("\n" + "=" * 70)
        print("✅ PROCESO COMPLETADO EXITOSAMENTE")
        print("=" * 70)
        print("\n📂 Los archivos se guardaron en la misma carpeta que este script")
        
    except Exception as e:
        print(f"\n❌ ERROR durante el procesamiento:")
        print(f"   {str(e)}")
        print("\nSi el error persiste, contacta con tu desarrollador.")

if __name__ == "__main__":
    main()