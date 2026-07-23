# MEPyD - Análisis de Ejecución de Proyectos de Inversión

Aplicación Streamlit para analizar el dataset oficial del Ministerio de Economía, Planificación y Desarrollo (MEPyD): "Estadísticas Informe de Ejecución de Proyectos de Inversión, 2018 - 2024".

## Archivos

- `mepyd_streamlit_app.py`: aplicación principal de Streamlit.
- `requirements.txt`: dependencias Python necesarias.
- `Dockerfile`: imagen para ejecutar la app en contenedor.

## Requisitos

- Python 3.9+ instalado
- `pip` instalado
- Opcional: Docker si se desea ejecutar en contenedor

## Instalación local

1. Instalar dependencias:

```bash
pip install -r requirements.txt
```

2. Ejecutar la aplicación:

```bash
streamlit run mepyd_streamlit_app.py
```

3. Abrir en el navegador:

```text
http://localhost:8501
```

## Uso

- La aplicación descarga automáticamente el dataset desde el sitio del MEPyD.
- Si no hay conexión, también permite subir un archivo CSV o Excel manualmente.
- Ofrece filtros por periodo, institución ejecutora, región/provincia y ODS.
- Presenta métricas generales, análisis de brechas, inversión territorial y un explorador exportable.

## Uso con Docker

1. Construir la imagen:

```bash
docker build -t mepyd-streamlit .
```

2. Ejecutar el contenedor:

```bash
docker run -p 8501:8501 mepyd-streamlit
```

3. Abrir en el navegador:

```text
http://localhost:8501
```

## Notas

- El dataset se carga con cache de Streamlit para no volver a descargarlo en cada interacción.
- Las columnas financieras se limpian y normalizan automáticamente.
- Se generaron columnas calculadas de `% Ejecución Financiera` y `Subejecución / Brecha`.
