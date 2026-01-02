import pandas as pd
from pathlib import Path

def clasificar_familia(familia):
    """
    Clasifica cada familia en una macro-categoría
    
    Args:
        familia: Texto de la columna Familia
    
    Returns:
        Nombre de la macro-categoría
    """
    familia_upper = str(familia).upper()
    
    # Palabras clave para MOBILIARIO
    mobiliario_keywords = [
        'SILLA', 'MESA', 'ESCRITORIO', 'PUPITRE', 'ARCHIVERO',
        'BANCA', 'BANCO', 'ANAQUEL', 'LIBRERO', 'LOCKER',
        'CASILLERO', 'GABINETE UNIVERSAL', 'CREDENZA', 'SEMI EJECUTIVO'
    ]
    
    # Palabras clave para EQUIPO DE CÓMPUTO
    computo_keywords = [
        'COMPUTADORA', 'COMPUTACION', 'LAPTOP', 'MONITOR',
        'IMPRESORA', 'MULTIFUNCIONAL', 'PROYECTOR', 'NO-BREAK',
        'SCANNER', 'DIGITALIZADOR', 'SWITCH', 'SERVIDOR',
        'ALMACENAMIENTO', 'CAMARA DE VIDEO', 'MULTIMEDIA',
        'DISPOSITIVO CONTROLADOR', 'RELEVADOR'
    ]
    
    # Palabras clave para EQUIPO MÉDICO
    medico_keywords = [
        'MEDICO', 'QUIRURGICO', 'CLINICA', 'ESTETOSCOPIO',
        'BAUMANOMETRO', 'CAMA CLINICA', 'CAMA PEDIATRICA',
        'CARRO CAMILLA', 'CARRO CURACIONES', 'CARRO CUNA',
        'CARRO ROJO', 'CARRO MONITOR', 'GLUCOMETRO', 'OXIMETRO',
        'INCUBADORA', 'NEGATOSCOPIO', 'ASPIRADOR', 'RESUCITADOR',
        'ELECTROCARDIOGRAFO', 'ELECTROTERAPIA', 'ULTRASONIDO',
        'ORTOPANTOMOGRAFO', 'LAVABO CIRUJANO', 'MESA INSTRUMENTAL',
        'ORINAL', 'SILLA DE RUEDAS', 'BASCULA', 'CUNA DE CALOR',
        'LAMPARA QUIRURGICA', 'LAMPARA DE FOTOTERAPIA', 'DOPPLER',
        'HIELERA VACUNAS', 'REFRIGERADOR', 'PORTA LEBRILLO',
        'ELECTROTERAPIA', 'REHABILITACION', 'LASER TERAPEUTICO',
        'MESA PEDIATRICA', 'MESA EXPLORACION', 'ELECTROCIRUGIA',
        'DESFIBRILADOR', 'MONITOR DE PRESION', 'MONITOR DE SIGNOS',
        'BOMBA DE INFUSION', 'FOTOTERAPIA', 'SUTURA', 'CIRUGIA',
        'DISECCION', 'BOTIQUIN', 'MANDIL EMPLOMADO', 'RADIACION',
        'EXAMINACION CLINICA', 'DIAGNOSTICO BASICO', 'ESTACION DE LAVAMANOS',
        'ESCANER DE CAMA', 'INFANTOMETRO', 'ESTADIMETRO'
    ]
    
    # Palabras clave para INSTRUMENTAL CIENTÍFICO
    cientifico_keywords = [
        'MICROSCOPIO', 'DENSIMETRO', 'BRUJULA', 'ANTROPOMETRO',
        'PLICOMETRO', 'CALIBRADOR', 'MANOMETRO', 'PENETROMETRO',
        'FOTOMETRO', 'ESPECTRO-FOTOMETRO', 'EVAPORADOR',
        'AGITADOR', 'BALANZA', 'TRIPIE', 'INSTRUMENTO CIENTIFICO',
        'APARATO CIENTIFICO', 'SENSOR', 'REGULADOR OXIGENO',
        'MEDIDOR DEL PH', 'SISMOGRAFO', 'TURBIDIMETRO',
        'MEDIDOR DE OXIGENO', 'LENTE PROTECTOR', 'APARATO EXTRACCION',
        'APARATO DISPERSION'
    ]
    
    # Clasificación por prioridad
    if any(keyword in familia_upper for keyword in computo_keywords):
        return 'EQUIPO_DE_COMPUTO'
    elif any(keyword in familia_upper for keyword in medico_keywords):
        return 'EQUIPO_MEDICO'
    elif any(keyword in familia_upper for keyword in cientifico_keywords):
        return 'INSTRUMENTAL_CIENTIFICO'
    elif any(keyword in familia_upper for keyword in mobiliario_keywords):
        return 'MOBILIARIO'
    else:
        return 'OTROS'

def separar_por_categorias(archivo_consolidado):
    """
    Lee el archivo consolidado y lo separa en múltiples archivos por categoría
    
    Args:
        archivo_consolidado: Ruta del archivo Excel consolidado
    """
    print("=" * 70)
    print("SEPARADOR DE CATEGORÍAS - INICIO")
    print("=" * 70)
    
    # Leer el archivo consolidado
    print(f"\nLeyendo archivo: {archivo_consolidado}")
    df = pd.read_excel(archivo_consolidado)
    
    print(f"Total de lotes a clasificar: {len(df)}")
    
    # Aplicar clasificación
    print("\nClasificando lotes por categoría...")
    df['Categoria'] = df['Familia'].apply(clasificar_familia)
    
    # Contar por categoría
    conteo_categorias = df['Categoria'].value_counts()
    print("\nDistribución por categoría:")
    for categoria, cantidad in conteo_categorias.items():
        print(f"  - {categoria}: {cantidad} lotes")
    
    # Crear carpeta para los archivos separados
    carpeta_salida = Path('CATEGORIAS_SEPARADAS')
    carpeta_salida.mkdir(exist_ok=True)
    print(f"\nCreando archivos en la carpeta: {carpeta_salida}")
    
    # Separar y guardar cada categoría
    archivos_creados = []
    
    for categoria in df['Categoria'].unique():
        # Filtrar datos de esta categoría
        df_categoria = df[df['Categoria'] == categoria].copy()
        
        # Eliminar la columna auxiliar 'Categoria' antes de guardar
        df_categoria_final = df_categoria.drop(columns=['Categoria'])
        
        # Ordenar por cantidad descendente
        df_categoria_final = df_categoria_final.sort_values('Cantidad', ascending=False)
        
        # Nombre del archivo
        nombre_archivo = carpeta_salida / f"{categoria}.xlsx"
        
        # Guardar archivo
        df_categoria_final.to_excel(nombre_archivo, index=False, sheet_name=categoria)
        
        archivos_creados.append({
            'categoria': categoria,
            'archivo': nombre_archivo,
            'lotes': len(df_categoria_final),
            'cantidad_total': df_categoria_final['Cantidad'].sum()
        })
        
        print(f"  Archivo creado: {nombre_archivo}")
    
    # Resumen final
    print("\n" + "=" * 70)
    print("RESUMEN DE ARCHIVOS CREADOS")
    print("=" * 70)
    
    for info in archivos_creados:
        print(f"\nCategoría: {info['categoria']}")
        print(f"  Archivo: {info['archivo']}")
        print(f"  Lotes únicos: {info['lotes']}")
        print(f"  Cantidad total: {info['cantidad_total']:.0f}")
    
    # Crear archivo de resumen
    df_resumen = pd.DataFrame(archivos_creados)
    archivo_resumen = carpeta_salida / "RESUMEN_CATEGORIAS.xlsx"
    df_resumen.to_excel(archivo_resumen, index=False, sheet_name='Resumen')
    
    print(f"\nArchivo de resumen creado: {archivo_resumen}")
    
    print("\n" + "=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print(f"\nTodos los archivos se guardaron en: {carpeta_salida.absolute()}")
    print("\nPuedes enviar cada archivo al proveedor correspondiente:")
    print("  - MOBILIARIO.xlsx -> Proveedores de mobiliario")
    print("  - EQUIPO_DE_COMPUTO.xlsx -> Proveedores de tecnología")
    print("  - EQUIPO_MEDICO.xlsx -> Proveedores de equipo médico")
    print("  - INSTRUMENTAL_CIENTIFICO.xlsx -> Proveedores de laboratorio")
    
    if 'OTROS' in df['Categoria'].values:
        print("\nNOTA: Se creó un archivo OTROS.xlsx con elementos no clasificados.")
        print("Revisa este archivo para determinar a qué proveedor enviarlo.")

def main():
    """
    Función principal
    """
    print("\n" + "=" * 70)
    print("SEPARADOR AUTOMÁTICO POR CATEGORÍAS")
    print("=" * 70)
    
    # Buscar el archivo consolidado
    archivo_default = 'REQCONS_CONSOLIDADO.xlsx'
    
    if Path(archivo_default).exists():
        print(f"\nSe encontró el archivo: {archivo_default}")
        usar_default = input("¿Deseas usar este archivo? (S/N): ").strip().upper()
        
        if usar_default == 'S':
            archivo_consolidado = archivo_default
        else:
            print("\nIngresa la ruta del archivo consolidado:")
            archivo_consolidado = input("Ruta: ").strip().strip('"')
    else:
        print("\nIngresa la ruta del archivo consolidado:")
        archivo_consolidado = input("Ruta: ").strip().strip('"')
    
    # Verificar que existe
    if not Path(archivo_consolidado).exists():
        print(f"\nERROR: El archivo no existe: {archivo_consolidado}")
        return
    
    try:
        separar_por_categorias(archivo_consolidado)
    except Exception as e:
        print(f"\nERROR durante el procesamiento:")
        print(f"  {str(e)}")

if __name__ == "__main__":
    main()