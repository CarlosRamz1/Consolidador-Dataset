# Consolidador-Dataset

Sistema automatizado para consolidar y analizar requisiciones de licitaciones gubernamentales en formato PDF.

## Descripción

Consolidador-Dataset es una herramienta desarrollada en Python que extrae, procesa y consolida información de documentos PDF de licitaciones gubernamentales. El sistema identifica automáticamente lotes repetidos en múltiples partidas y genera reportes consolidados en formato Excel, facilitando el proceso de cotización y análisis.

## Características principales

- Extracción automática de tablas desde archivos PDF de múltiples páginas
- Identificación inteligente de lotes duplicados basada en nombre y descripción
- Consolidación automática de cantidades para artículos repetidos
- Generación de reportes en formato Excel
- Limpieza y validación de datos
- Análisis estadístico de repeticiones
- Manejo robusto de errores

## Requisitos del sistema

- Python 3.8 o superior
- Windows, macOS o Linux
- 100 MB de espacio en disco
- 2 GB de RAM (recomendado para documentos grandes)

## Instalación

### 1. Clonar el repositorio

git clone https://github.com/CarlosRamz1/Consolidador-Dataset.git
cd Consolidador-Dataset

### 2. Crear entorno virtual (recomendado)

python -m venv venv

### 3. Activar el entorno virtual

venv\Scripts\activate

### 4. Instalar dependencias

pip install -r requirements.txt

## Uso

### Ejecución básica

python consolidador.py

El programa solicitará la ruta del archivo PDF a procesar.

### Ejemplo de flujo de trabajo

1. Ejecutar el programa
2. Ingresar la ruta completa del archivo PDF
3. Esperar el procesamiento (el tiempo depende del tamaño del documento)
4. Revisar los archivos Excel generados en la misma carpeta

### Archivos de salida

El programa genera dos archivos Excel:

- `[nombre]_CONSOLIDADO.xlsx`: Contiene los lotes agrupados con cantidades sumadas
- `[nombre]_DATOS_ORIGINALES.xlsx`: Contiene todos los datos extraídos sin procesar

## Estructura de datos

### Entrada esperada

El programa espera tablas PDF con la siguiente estructura:

| Cantidad | Nombre | Descripción | Familia | Unidad | Observaciones |
|----------|--------|-------------|---------|--------|---------------|

### Salida consolidada

| Cantidad | Nombre | Descripción | Familia | Unidad | Veces_Repetido | Observaciones |
|----------|--------|-------------|---------|--------|----------------|---------------|

## Funcionalidades técnicas

### Extracción de datos
- Utiliza pdfplumber para extraer tablas con precisión
- Ignora contenido no tabular (encabezados, firmas, texto libre)
- Procesa múltiples tablas por página

### Limpieza de datos
- Eliminación de filas vacías o inválidas
- Normalización de espacios en blanco
- Conversión de tipos de datos
- Validación de cantidades numéricas

### Consolidación
- Agrupación por nombre y descripción exacta
- Suma automática de cantidades
- Concatenación de observaciones únicas
- Conteo de repeticiones

## Casos de uso

- Análisis de requisiciones para licitaciones gubernamentales
- Consolidación de compras departamentales
- Preparación de cotizaciones unificadas
- Auditoría de solicitudes de compra
- Optimización de procesos de adquisición

## Tecnologías utilizadas

- **Python 3.x**: Lenguaje de programación principal
- **pdfplumber**: Extracción de tablas desde PDF
- **pandas**: Manipulación y análisis de datos
- **openpyxl**: Generación de archivos Excel

## Limitaciones conocidas

- El PDF debe contener tablas con estructura definida
- Los nombres de columnas deben seguir el formato esperado
- Archivos extremadamente grandes (>1000 páginas) pueden requerir mayor tiempo de procesamiento

## Historial de versiones

### v1.0.0 (2025-12-20)
- Lanzamiento inicial
- Extracción básica de tablas PDF
- Consolidación de lotes duplicados
- Generación de reportes Excel

## 

Desarrollado para optimizar procesos de licitaciones gubernamentales y facilitar el trabajo de áreas de compras y adquisiciones.